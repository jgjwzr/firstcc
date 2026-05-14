"""Entry point for Zahner Plotter v2."""

import tkinter as tk
from zahner_plotter.model import Model
from zahner_plotter.view import View
from zahner_plotter.controller import Controller


def main():
    root = tk.Tk()
    model = Model()
    view = View(root, model)
    ctrl = Controller(root, model, view)
    ctrl.run()


if __name__ == "__main__":
    main()
