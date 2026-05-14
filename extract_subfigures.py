"""
Extract individual sub-figures (a, b, c, ...) from composite figure images.
Main paper: split composite images by detecting white-space separators.
SI paper: extract individual image blocks from each page.
Uses Pillow for image processing (numpy-free).
"""
import fitz
import json
import os
import sys
import re
from io import BytesIO
from PIL import Image

MAIN_PDF = "D:/firstcc/论文正文.pdf"
SI_PDF = "D:/firstcc/论文si.pdf"
OUTPUT_DIR = "D:/firstcc/images"
RENDER_DPI = 300

os.makedirs(OUTPUT_DIR, exist_ok=True)


def detect_grid_layout(img):
    """Detect sub-figure grid by finding major white separator rows/columns."""
    w, h = img.size
    if w < 50 or h < 50:
        return [(0, 0, w, h)]

    gray = img.convert("L")
    pixels = gray.load()

    min_gap_px = 15
    span_threshold = 0.82
    white_threshold = 0.95

    # Horizontal gaps: rows that span >82% of image width
    h_gaps = []
    in_gap = False
    gap_start = 0
    for y in range(h):
        white_count = 0
        for x in range(w):
            if pixels[x, y] >= 248:
                white_count += 1
        if white_count / w >= span_threshold:
            if not in_gap:
                gap_start = y
                in_gap = True
        else:
            if in_gap and (y - gap_start) >= min_gap_px:
                gap_white = 0
                gap_total = (y - gap_start) * w
                for gy in range(gap_start, y):
                    for gx in range(w):
                        if pixels[gx, gy] >= 248:
                            gap_white += 1
                if gap_total > 0 and gap_white / gap_total >= white_threshold:
                    h_gaps.append((gap_start, y))
            in_gap = False

    # Vertical gaps
    v_gaps = []
    in_gap = False
    gap_start = 0
    for x in range(w):
        white_count = 0
        for y in range(h):
            if pixels[x, y] >= 248:
                white_count += 1
        if white_count / h >= span_threshold:
            if not in_gap:
                gap_start = x
                in_gap = True
        else:
            if in_gap and (x - gap_start) >= min_gap_px:
                gap_white = 0
                gap_total = (x - gap_start) * h
                for gy in range(h):
                    for gx in range(gap_start, x):
                        if pixels[gx, gy] >= 248:
                            gap_white += 1
                if gap_total > 0 and gap_white / gap_total >= white_threshold:
                    v_gaps.append((gap_start, x))
            in_gap = False

    if not h_gaps and not v_gaps:
        return [(0, 0, w, h)]

    h_cuts = [0] + [(gs + ge) // 2 for gs, ge in h_gaps] + [h]
    v_cuts = [0] + [(gs + ge) // 2 for gs, ge in v_gaps] + [w]

    cells = []
    for yi in range(len(h_cuts) - 1):
        for xi in range(len(v_cuts) - 1):
            x0, x1 = v_cuts[xi], v_cuts[xi + 1]
            y0, y1 = h_cuts[yi], h_cuts[yi + 1]
            cw, ch = x1 - x0, y1 - y0
            if cw < 30 or ch < 30:
                continue
            cell_white = 0
            cell_total = cw * ch
            for cy in range(y0, min(y1, h)):
                for cx in range(x0, min(x1, w)):
                    if pixels[cx, cy] >= 248:
                        cell_white += 1
            if cell_total > 0 and cell_white / cell_total > 0.98:
                continue
            cells.append((x0, y0, x1, y1))

    return cells if cells else [(0, 0, w, h)]


def trim_white_borders(img, padding=5):
    """Trim white borders from an image."""
    w, h = img.size
    gray = img.convert("L")
    pixels = gray.load()

    non_white_x = []
    non_white_y = []
    for y in range(h):
        for x in range(w):
            if pixels[x, y] < 240:
                non_white_x.append(x)
                non_white_y.append(y)

    if not non_white_x:
        return img

    x0 = max(min(non_white_x) - padding, 0)
    x1 = min(max(non_white_x) + padding, w)
    y0 = max(min(non_white_y) - padding, 0)
    y1 = min(max(non_white_y) + padding, h)

    return img.crop((x0, y0, x1, y1))


def render_and_split(pdf_path, page_0idx, region_bbox, fig_label, dpi=RENDER_DPI):
    """Render a region of a PDF page, split into sub-figures by grid detection."""
    doc = fitz.open(pdf_path)
    page = doc[page_0idx]

    margin = 5
    x0 = max(region_bbox[0] - margin, 0)
    y0 = max(region_bbox[1] - margin, 0)
    x1 = min(region_bbox[2] + margin, page.rect.width)
    y1 = min(region_bbox[3] + margin, page.rect.height)

    mat = fitz.Matrix(dpi / 72, dpi / 72)
    clip = fitz.Rect(x0, y0, x1, y1)
    pix = page.get_pixmap(matrix=mat, clip=clip)
    doc.close()

    img = Image.open(BytesIO(pix.tobytes("png")))
    cells = detect_grid_layout(img)

    saved = []
    labels = "abcdefghijklmnopqrstuvwxyz"

    for ci, (cx0, cy0, cx1, cy1) in enumerate(cells):
        sub_img = img.crop((cx0, cy0, cx1, cy1))
        sub_img = trim_white_borders(sub_img)

        if sub_img.size[0] < 15 or sub_img.size[1] < 15:
            continue

        label = labels[ci] if ci < len(labels) else str(ci)
        fname = f"{fig_label}{label}.png"
        path = os.path.join(OUTPUT_DIR, fname)
        sub_img.save(path)
        saved.append(fname)

    return saved


def extract_si_blocks(pdf_path, page_0idx, fig_label, dpi=RENDER_DPI):
    """Extract individual image blocks from an SI page."""
    doc = fitz.open(pdf_path)
    page = doc[page_0idx]

    blocks = page.get_text("dict")["blocks"]
    img_blocks = []
    for b in blocks:
        if b["type"] == 1:
            w = b["bbox"][2] - b["bbox"][0]
            h = b["bbox"][3] - b["bbox"][1]
            if w > 20 and h > 10:
                img_blocks.append(list(b["bbox"]))

    if not img_blocks:
        doc.close()
        return render_and_split(
            pdf_path, page_0idx,
            [50, 50, page.rect.width - 50, page.rect.height - 50],
            fig_label, dpi
        )

    img_blocks.sort(key=lambda b: (round(b[1], -1), b[0]))

    saved = []
    labels = "abcdefghijklmnopqrstuvwxyz"

    for idx, bbox in enumerate(img_blocks):
        margin = 3
        x0 = max(bbox[0] - margin, 0)
        y0 = max(bbox[1] - margin, 0)
        x1 = min(bbox[2] + margin, page.rect.width)
        y1 = min(bbox[3] + margin, page.rect.height)

        mat = fitz.Matrix(dpi / 72, dpi / 72)
        clip = fitz.Rect(x0, y0, x1, y1)
        pix = page.get_pixmap(matrix=mat, clip=clip)

        img = Image.open(BytesIO(pix.tobytes("png")))
        img = trim_white_borders(img)

        if img.size[0] < 10 or img.size[1] < 10:
            continue

        label = labels[idx] if idx < len(labels) else str(idx)
        fname = f"{fig_label}{label}.png"
        path = os.path.join(OUTPUT_DIR, fname)
        img.save(path)
        saved.append(fname)

    doc.close()
    return saved


def main():
    print("=" * 60)
    print("Extracting sub-figures from main paper and SI")
    print("=" * 60)

    all_figures = {}

    # ============================================================
    # Main paper figures
    # ============================================================
    print("\n--- Main Paper Figures ---")

    doc = fitz.open(MAIN_PDF)
    main_figure_pages = {
        2: "Figure1",
        3: "Figure2",
        4: "Figure3",
        5: "Figure4",
        6: "Figure5",
    }

    for page_1idx, fig_label in main_figure_pages.items():
        page_0idx = page_1idx - 1
        page = doc[page_0idx]

        blocks = page.get_text("dict")["blocks"]
        img_bbox = None
        for b in blocks:
            if b["type"] == 1:
                w = b["bbox"][2] - b["bbox"][0]
                h = b["bbox"][3] - b["bbox"][1]
                if w > 100 and h > 100:
                    img_bbox = list(b["bbox"])
                    break

        if img_bbox is None:
            print(f"  {fig_label}: No large image found on page {page_1idx}")
            continue

        print(f"  {fig_label}: bbox={[round(v,1) for v in img_bbox]}")
        saved = render_and_split(MAIN_PDF, page_0idx, img_bbox, fig_label)
        all_figures[fig_label] = saved
        for s in saved:
            print(f"    -> {s}")

    doc.close()

    # ============================================================
    # SI figures
    # ============================================================
    print("\n--- SI Figures ---")

    doc = fitz.open(SI_PDF)
    extracted_si = set()

    for page_0idx in range(doc.page_count):
        page = doc[page_0idx]
        text = page.get_text("text")

        blocks = page.get_text("dict")["blocks"]
        has_images = any(
            b["type"] == 1 and (b["bbox"][2] - b["bbox"][0]) > 20
            and (b["bbox"][3] - b["bbox"][1]) > 10
            for b in blocks
        )

        # Figures
        for m in re.finditer(r'Figure\s+(S\d+)', text):
            fig_label = f"Figure_{m.group(1)}"  # Figure_S1, Figure_S2, etc.

            if fig_label in extracted_si:
                continue
            if not has_images:
                continue  # cross-reference on a non-figure page

            print(f"  {fig_label} (page {page_0idx+1})")
            saved = extract_si_blocks(SI_PDF, page_0idx, fig_label)
            all_figures[fig_label] = saved
            extracted_si.add(fig_label)
            for s in saved:
                print(f"    -> {s}")
            break

        # Tables
        for m in re.finditer(r'Table\s+(S\d+)', text):
            table_label = f"Table_{m.group(1)}"

            if table_label in extracted_si:
                continue

            if not has_images:
                print(f"  {table_label} (page {page_0idx+1}, full page)")
                mat = fitz.Matrix(200 / 72, 200 / 72)
                pix = page.get_pixmap(matrix=mat)
                img = Image.open(BytesIO(pix.tobytes("png")))
                fname = f"{table_label}.png"
                path = os.path.join(OUTPUT_DIR, fname)
                img.save(path)
                all_figures[table_label] = [fname]
                extracted_si.add(table_label)
                print(f"    -> {fname}")
                break
            break

    doc.close()

    # Save manifest
    print(f"\n=== Summary ===")
    total = sum(len(v) for v in all_figures.values())
    print(f"Figures/tables: {len(all_figures)}")
    print(f"Sub-figures: {total}")

    manifest_path = "D:/firstcc/subfigure_manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(all_figures, f, ensure_ascii=False, indent=2)
    print(f"Manifest: {manifest_path}")


if __name__ == "__main__":
    main()
