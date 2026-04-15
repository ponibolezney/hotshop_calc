import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlsxwriter


POSITION_OPTIONS = ["У стены", "Свободно стоящее"]


class EquipmentBlock:
    def __init__(self, parent, room_block, number, data=None):
        self.parent = parent
        self.room_block = room_block
        self.number = number
        self.entries = {}

        self.frame = tk.Frame(parent, bd=3, relief="solid", padx=10, pady=10)
        self.frame.pack(anchor="w", fill="x", padx=20, pady=8)

        default_name = f"Оборудование{number}"
        if data and data.get("name"):
            default_name = data["name"]

        self.name_var = tk.StringVar(value=default_name)
        name_entry = tk.Entry(self.frame, textvariable=self.name_var, width=30, font=("Arial", 12))
        name_entry.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.entries["name"] = name_entry

        fields = [
            ("quantity", "Количество"),
            ("qy", "Qy"),
            ("ka", "Ka"),
            ("kz", "Kz"),
            ("length", "Длина"),
            ("width", "Ширина"),
            ("height", "Высота"),
        ]

        for i, (key, label) in enumerate(fields, start=1):
            tk.Label(self.frame, text=label, font=("Arial", 11)).grid(
                row=i, column=0, sticky="w", padx=(0, 10), pady=4
            )
            entry = tk.Entry(self.frame, width=18, font=("Arial", 11))
            if data and key in data:
                entry.insert(0, str(data[key]))
            entry.grid(row=i, column=1, sticky="w", pady=4)
            self.entries[key] = entry

        position_row = len(fields) + 1
        tk.Label(self.frame, text="Положение", font=("Arial", 11)).grid(
            row=position_row, column=0, sticky="w", padx=(0, 10), pady=4
        )

        self.position_var = tk.StringVar()
        if data and data.get("position") in POSITION_OPTIONS:
            self.position_var.set(data["position"])
        else:
            self.position_var.set("")

        position_combo = ttk.Combobox(
            self.frame,
            textvariable=self.position_var,
            values=POSITION_OPTIONS,
            state="readonly",
            width=16,
            font=("Arial", 11)
        )
        position_combo.grid(row=position_row, column=1, sticky="w", pady=4)
        self.entries["position"] = position_combo

        self.delete_button = tk.Button(
            self.frame,
            text="Удалить оборудование",
            font=("Arial", 11),
            command=self.delete_self
        )
        self.delete_button.grid(row=position_row + 1, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def delete_self(self):
        if messagebox.askyesno("Подтверждение", "Удалить это оборудование?"):
            self.frame.destroy()
            self.room_block.remove_equipment(self)

    def reset_highlight(self):
        for widget in self.entries.values():
            if isinstance(widget, ttk.Combobox):
                try:
                    widget.configure(style="TCombobox")
                except Exception:
                    pass
            else:
                widget.config(bg="white")

    def mark_invalid(self, key):
        widget = self.entries[key]
        if isinstance(widget, ttk.Combobox):
            try:
                widget.configure(style="Invalid.TCombobox")
            except Exception:
                pass
        else:
            widget.config(bg="#ffb3b3")

    def validate(self):
        self.reset_highlight()
        ok = True

        required_keys = ["name", "quantity", "qy", "ka", "kz", "position", "length", "width", "height"]

        for key in required_keys:
            value = self.get_widget_value(key)
            if not value:
                self.mark_invalid(key)
                ok = False

        if not ok:
            return False

        if not self.is_positive_int(self.get_widget_value("quantity")):
            self.mark_invalid("quantity")
            ok = False

        for key in ["qy", "ka", "kz", "length", "width", "height"]:
            if not self.is_number(self.get_widget_value(key)):
                self.mark_invalid(key)
                ok = False

        return ok

    def get_widget_value(self, key):
        widget = self.entries[key]
        return widget.get().strip()

    @staticmethod
    def is_positive_int(value):
        try:
            number = int(value)
            return number >= 1
        except Exception:
            return False

    @staticmethod
    def is_number(value):
        try:
            float(value)
            return True
        except Exception:
            return False

    def get_data(self):
        return {
            "name": self.get_widget_value("name"),
            "quantity": int(self.get_widget_value("quantity")),
            "qy": float(self.get_widget_value("qy")),
            "ka": float(self.get_widget_value("ka")),
            "kz": float(self.get_widget_value("kz")),
            "position": self.get_widget_value("position"),
            "length": float(self.get_widget_value("length")),
            "width": float(self.get_widget_value("width")),
            "height": float(self.get_widget_value("height")),
        }


class RoomBlock:
    def __init__(self, parent, app, number, data=None):
        self.parent = parent
        self.app = app
        self.number = number
        self.equipment_blocks = []

        self.frame = tk.Frame(parent, bd=3, relief="solid", padx=10, pady=10)
        self.frame.pack(anchor="w", fill="x", padx=20, pady=12)

        top_row = tk.Frame(self.frame)
        top_row.pack(anchor="w", fill="x", pady=(0, 10))

        default_name = f"Помещение{number}"
        if data and data.get("room_name"):
            default_name = data["room_name"]

        self.name_var = tk.StringVar(value=default_name)
        self.name_entry = tk.Entry(top_row, textvariable=self.name_var, width=28, font=("Arial", 12))
        self.name_entry.pack(side="left")

        self.delete_room_button = tk.Button(
            top_row,
            text="Удалить помещение",
            font=("Arial", 10),
            command=self.delete_self
        )
        self.delete_room_button.pack(side="left", padx=(12, 0))

        self.equipment_container = tk.Frame(self.frame)
        self.equipment_container.pack(anchor="w", fill="x")

        self.add_equipment_button = tk.Button(
            self.frame,
            text="Добавить оборудование",
            font=("Arial", 11),
            command=self.add_equipment
        )
        self.add_equipment_button.pack(anchor="w", pady=(10, 0))

        if data and "equipment" in data:
            for eq_data in data["equipment"]:
                self.add_equipment(eq_data)

    def delete_self(self):
        if messagebox.askyesno("Подтверждение", "Удалить это помещение со всем оборудованием?"):
            self.frame.destroy()
            self.app.remove_room(self)

    def add_equipment(self, data=None):
        equipment_number = len(self.equipment_blocks) + 1
        block = EquipmentBlock(self.equipment_container, self, equipment_number, data=data)
        self.equipment_blocks.append(block)

    def remove_equipment(self, equipment_block):
        if equipment_block in self.equipment_blocks:
            self.equipment_blocks.remove(equipment_block)

    def reset_highlight(self):
        self.name_entry.config(bg="white")
        self.add_equipment_button.config(bg="SystemButtonFace")
        for eq in self.equipment_blocks:
            eq.reset_highlight()

    def validate(self):
        self.reset_highlight()
        ok = True

        if not self.name_entry.get().strip():
            self.name_entry.config(bg="#ffb3b3")
            ok = False

        if not self.equipment_blocks:
            self.add_equipment_button.config(bg="#ffb3b3")
            ok = False

        for eq in self.equipment_blocks:
            if not eq.validate():
                ok = False

        return ok

    def get_data(self):
        return {
            "room_name": self.name_entry.get().strip(),
            "equipment": [eq.get_data() for eq in self.equipment_blocks]
        }


class NewCalculationWindow:
    def __init__(self, root, loaded_data=None):
        self.root = root
        self.root.title("Новый расчёт")
        self.root.geometry("1150x800")
        self.root.minsize(950, 650)

        self.room_blocks = []

        self.setup_styles()

        self.main_frame = tk.Frame(root)
        self.main_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)

        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.bottom_frame = tk.Frame(root, pady=10)
        self.bottom_frame.pack(fill="x", side="bottom")

        self.constants_entries = {}
        self.build_constants_block()

        self.rooms_container = tk.Frame(self.scrollable_frame)
        self.rooms_container.pack(anchor="w", fill="x")

        self.add_room_button = tk.Button(
            self.scrollable_frame,
            text="Добавить помещение",
            font=("Arial", 12),
            command=self.add_room
        )
        self.add_room_button.pack(anchor="w", padx=20, pady=20)

        self.save_json_button = tk.Button(
            self.bottom_frame,
            text="Сохранить JSON...",
            font=("Arial", 12),
            command=self.save_json
        )
        self.save_json_button.pack(side="right", padx=(10, 20))

        self.save_excel_button = tk.Button(
            self.bottom_frame,
            text="Сохранить Excel...",
            font=("Arial", 12),
            command=self.save_excel
        )
        self.save_excel_button.pack(side="right", padx=(10, 0))

        self.cancel_button = tk.Button(
            self.bottom_frame,
            text="Отмена",
            font=("Arial", 12),
            command=self.root.destroy
        )
        self.cancel_button.pack(side="right", padx=(0, 10))

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        if loaded_data:
            self.load_into_form(loaded_data)

    def setup_styles(self):
        style = ttk.Style()
        style.configure("TCombobox", fieldbackground="white")
        style.configure("Invalid.TCombobox", fieldbackground="#ffb3b3")

    def build_constants_block(self):
        block = tk.Frame(self.scrollable_frame, bd=3, relief="solid", padx=10, pady=10)
        block.pack(anchor="w", fill="x", padx=20, pady=(20, 10))

        title = tk.Label(block, text="Константы проекта", font=("Arial", 13, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        fields = [
            ("k1", "K1"),
            ("k2", "K2"),
            ("distance_to_hood", "Расстояние до местного отсоса"),
            ("mobility_coefficient", "Поправочный коэффициент"),
            ("hood_efficiency", "Коэффициент эффективности"),
        ]

        defaults = {
            "k1": "",
            "k2": "",
            "distance_to_hood": "",
            "mobility_coefficient": "",
            "hood_efficiency": "",
        }

        for i, (key, label) in enumerate(fields, start=1):
            tk.Label(block, text=label, font=("Arial", 11)).grid(
                row=i, column=0, sticky="w", padx=(0, 10), pady=4
            )
            entry = tk.Entry(block, width=20, font=("Arial", 11))
            entry.insert(0, defaults[key])
            entry.grid(row=i, column=1, sticky="w", pady=4)
            self.constants_entries[key] = entry

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_room(self, data=None):
        room_number = len(self.room_blocks) + 1
        block = RoomBlock(self.rooms_container, self, room_number, data=data)
        self.room_blocks.append(block)

    def remove_room(self, room_block):
        if room_block in self.room_blocks:
            self.room_blocks.remove(room_block)

    def load_into_form(self, loaded_data):
        constants = loaded_data.get("constants", {})
        for key, entry in self.constants_entries.items():
            entry.delete(0, tk.END)
            if key in constants:
                entry.insert(0, str(constants[key]))

        rooms = loaded_data.get("rooms", [])
        for room_data in rooms:
            self.add_room(data=room_data)

    def reset_constants_highlight(self):
        for entry in self.constants_entries.values():
            entry.config(bg="white")

    def validate_constants(self):
        self.reset_constants_highlight()
        ok = True

        for key, entry in self.constants_entries.items():
            value = entry.get().strip()
            if not value:
                entry.config(bg="#ffb3b3")
                ok = False
                continue
            try:
                float(value)
            except Exception:
                entry.config(bg="#ffb3b3")
                ok = False

        return ok

    def validate_all(self):
        ok = True

        if not self.validate_constants():
            ok = False

        if not self.room_blocks:
            self.add_room_button.config(bg="#ffb3b3")
            ok = False
        else:
            self.add_room_button.config(bg="SystemButtonFace")

        for room in self.room_blocks:
            if not room.validate():
                ok = False

        return ok

    def collect_data(self):
        constants = {
            key: float(entry.get().strip())
            for key, entry in self.constants_entries.items()
        }

        return {
            "constants": constants,
            "rooms": [room.get_data() for room in self.room_blocks]
        }

    def save_json(self):
        if not self.validate_all():
            messagebox.showerror("Ошибка", "Не все параметры заполнены или заполнены неверно!")
            return

        data = self.collect_data()

        file_path = filedialog.asksaveasfilename(
            title="Сохранить JSON",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Успех", f"JSON успешно сохранён:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить JSON:\n{e}")

    def save_excel(self):
        if not self.validate_all():
            messagebox.showerror("Ошибка", "Не все параметры заполнены или заполнены неверно!")
            return

        data = self.collect_data()

        file_path = filedialog.asksaveasfilename(
            title="Сохранить Excel-файл",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:
            return

        try:
            export_to_excel(data, file_path)
            messagebox.showinfo("Успех", f"Excel успешно сохранён:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить Excel:\n{e}")


def export_to_excel(data, file_path):
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Данные")

    header_format = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "border": 1
    })

    cell_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    room_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "bold": True
    })

    const_title_format = workbook.add_format({
        "bold": True,
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    const_cell_format = workbook.add_format({
        "border": 1,
        "align": "left",
        "valign": "vcenter"
    })

    worksheet.set_column(0, 0, 28)
    worksheet.set_column(1, 5, 18)
    worksheet.set_column(6, 8, 12)

    row = 0

    worksheet.merge_range(row, 0, row, 1, "Константы проекта", const_title_format)
    row += 1

    constants_labels = [
        ("k1", "K1"),
        ("k2", "K2"),
        ("distance_to_hood", "Расстояние до местного отсоса"),
        ("mobility_coefficient", "Поправочный коэффициент"),
        ("hood_efficiency", "Коэффициент эффективности"),
    ]

    for key, label in constants_labels:
        worksheet.write(row, 0, label, const_cell_format)
        worksheet.write(row, 1, data["constants"][key], const_cell_format)
        row += 1

    row += 1

    headers = [
        "Наименование",
        "Количество",
        "Qy, кВт",
        "Ka, Вт/кВт",
        "Kz",
        "Расположение",
        "Длина",
        "Ширина",
        "Высота"
    ]

    for col, header in enumerate(headers):
        worksheet.write(row, col, header, header_format)
    row += 1

    for room in data["rooms"]:
        room_name = room["room_name"]
        equipment_list = room["equipment"]

        worksheet.merge_range(row, 0, row, 8, room_name, room_format)
        row += 1

        for eq in equipment_list:
            worksheet.write(row, 0, eq["name"], cell_format)
            worksheet.write(row, 1, eq["quantity"], cell_format)
            worksheet.write(row, 2, eq["qy"], cell_format)
            worksheet.write(row, 3, eq["ka"], cell_format)
            worksheet.write(row, 4, eq["kz"], cell_format)
            worksheet.write(row, 5, eq["position"], cell_format)
            worksheet.write(row, 6, eq["length"], cell_format)
            worksheet.write(row, 7, eq["width"], cell_format)
            worksheet.write(row, 8, eq["height"], cell_format)
            row += 1

    workbook.close()


class StartWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Модуль ввода данных")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        container = tk.Frame(root)
        container.place(relx=0.5, rely=0.5, anchor="center")

        self.new_button = tk.Button(
            container,
            text="Создать новый расчёт",
            font=("Arial", 16),
            width=24,
            height=2,
            command=self.create_new_calculation
        )
        self.new_button.pack(pady=15)

        self.load_button = tk.Button(
            container,
            text="Загрузить расчётные данные",
            font=("Arial", 16),
            width=24,
            height=2,
            command=self.load_existing_data
        )
        self.load_button.pack(pady=15)

    def create_new_calculation(self):
        self.root.destroy()
        new_root = tk.Tk()
        NewCalculationWindow(new_root)
        new_root.mainloop()

    def load_existing_data(self):
        file_path = filedialog.askopenfilename(
            title="Выбрать JSON-файл",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt")]
        )

        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            if "rooms" not in data or not isinstance(data["rooms"], list):
                raise ValueError("Неверная структура файла. Ожидался JSON с ключом 'rooms'.")

            self.root.destroy()
            new_root = tk.Tk()
            NewCalculationWindow(new_root, loaded_data=data)
            new_root.mainloop()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    StartWindow(root)
    root.mainloop()