import tkinter as tk
from tkinter import ttk

from excel_checker.ui.app import ExcelMandatoryCheckerApp


def main():
    root = tk.Tk()
    style = ttk.Style()

    try:
        style.theme_use("clam")
    except Exception:
        pass

    ExcelMandatoryCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
