"""
ZiiPOS Menu Converter V4 -- Entry Point
"""
import sys
import os
import tkinter as tk

if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LIB_DIR = os.path.join(BASE_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from app import App

# Ensure PyInstaller bundles openpyxl and its submodules
import openpyxl  # noqa: F401
import openpyxl.cell._writer  # noqa: F401


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
