"""
Generate an HTML preview page for all extracted sub-figures.
"""
import json
import os

MANIFEST = "D:/firstcc/subfigure_manifest.json"
IMAGES_DIR = "D:/firstcc/images"
OUTPUT = "D:/firstcc/figure_preview.html"

with open(MANIFEST, "r", encoding="utf-8") as f:
    manifest = json.load(f)

# Separate main, SI figures, tables
main_figs = {}
si_figs = {}
tables = {}
for k, v in manifest.items():
    if k.startswith("Figure") and not k.startswith("Figure_S"):
        main_figs[k] = v
    elif k.startswith("Figure_S"):
        si_figs[k] = v
    elif k.startswith("Table"):
        tables[k] = v


def render_section(title, items):
    rows = ""
    for fig_id in sorted(items.keys(), key=lambda x: (x.split("_S")[0] if "_S" in x else x,
        int(x.split("_S")[1]) if "_S" in x and x.split("_S")[1].isdigit() else 0)
        if "_S" not in x else (x, 0)):
        sub_figs = items[fig_id]
        imgs_html = ""
        for sf in sub_figs:
            img_path = f"images/{sf}"
            full = f"D:/firstcc/{img_path}"
            if os.path.exists(full):
                size_kb = os.path.getsize(full) / 1024
                imgs_html += (
                    f'<div class="img-wrap">'
                    f'<img src="{img_path}" alt="{sf}" loading="lazy">'
                    f'<div class="label">{sf} ({size_kb:.0f}KB)</div>'
                    f'</div>'
                )
            else:
                imgs_html += f'<div class="img-wrap missing"><div class="label">{sf} (missing)</div></div>'

        rows += (
            f'<div class="fig-row">'
            f'<h3>{fig_id} <span class="meta">({len(sub_figs)} panels)</span></h3>'
            f'<div class="img-grid">{imgs_html}</div>'
            f'</div>'
        )
    return f'<h2>{title} ({len(items)})</h2>{rows}'

html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Figure Review - Ni/MoO2 Paper</title>
<style>
  body {{ font-family: 'Microsoft YaHei', sans-serif; background: #f0f0f0; margin: 0; padding: 20px; }}
  h1 {{ color: #1B3A5C; font-size: 20px; }}
  h2 {{ color: #333; border-bottom: 2px solid #1B3A5C; padding-bottom: 5px; margin-top: 30px; font-size: 16px; }}
  h3 {{ color: #555; margin: 10px 0 5px; font-size: 14px; }}
  .meta {{ font-weight: normal; font-size: 0.85em; color: #999; }}
  .img-grid {{ display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }}
  .img-wrap {{ background: white; padding: 8px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
  .img-wrap img {{ max-width: 400px; max-height: 300px; display: block; }}
  .img-wrap.missing {{ background: #fff3f3; border: 2px dashed #c00; min-width: 100px; min-height: 60px; display: flex; align-items: center; justify-content: center; }}
  .label {{ font-size: 0.75em; color: #666; margin-top: 4px; }}
  .missing .label {{ color: #c00; }}
  .nav {{ position: fixed; top: 10px; right: 10px; background: white; padding: 10px 15px; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); font-size: 12px; }}
  .nav a {{ color: #1B3A5C; text-decoration: none; display: block; }}
  .nav a:hover {{ text-decoration: underline; }}
</style>
</head>
<body>
<div class="nav">
  <a href="#main">Main Figures</a>
  <a href="#si">SI Figures</a>
  <a href="#tables">Tables</a>
</div>
<h1>Figure Review - Ni/MoO2 Electrode Paper</h1>
<p>Review all {sum(len(v) for v in manifest.values())} extracted sub-figures across {len(manifest)} figures.</p>
<div id="main">{render_section('Main Paper Figures', main_figs)}</div>
<div id="si">{render_section('Supporting Information Figures', si_figs)}</div>
<div id="tables">{render_section('Tables', tables)}</div>
</body>
</html>"""

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(html)
print(f"Preview saved to: {OUTPUT}")
print(f"Total: {len(manifest)} figures, {sum(len(v) for v in manifest.values())} sub-figures")
