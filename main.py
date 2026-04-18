import os
import sys
import tkinter as tk
from tkinter import messagebox

from app.catalog import EquipmentCatalog
from app.ui_editor import StartWindow


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


CATALOG_PATH = resource_path(os.path.join("data", "equipment_catalog.xlsx"))


def main():
    if not os.path.exists(CATALOG_PATH):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Нет справочника",
            f"Файл справочника не найден:\n{CATALOG_PATH}"
        )
        root.destroy()
        return

    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    root = tk.Tk()
    StartWindow(root, catalog)
    root.mainloop()


if __name__ == "__main__":
    main()