import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from app.catalog import EquipmentCatalog
from app.schemas import ProjectInput
from app.storage import save_input_data, load_input_data
from app.excel_export import export_project_to_excel


DEFAULT_CONSTANTS = {
    "k1": 0.5,
    "k2": 0.7,
    "k_empirical": 180.0,
    "z_m": 1.1,
    "ko": 0.8,
    "a": 1.25,
}


class EquipmentBlock:
    def __init__(self, parent, room_block, number, catalog: EquipmentCatalog, data=None):
        self.parent = parent
        self.room_block = room_block
        self.catalog = catalog
        self.number = number
        self.entries = {}

        self.frame = tk.Frame(parent, bd=3, relief="solid", padx=10, pady=10)
        self.frame.pack(anchor="w", fill="x", padx=20, pady=8)

        default_name = f"Оборудование{number}"
        if data and data.get("name"):
            default_name = data["name"]

        tk.Label(self.frame, text="Наименование", font=("Arial", 11)).grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.name_var = tk.StringVar(value=default_name)
        name_entry = tk.Entry(self.frame, textvariable=self.name_var, width=32, font=("Arial", 11))
        name_entry.grid(row=0, column=1, sticky="w", pady=4)
        self.entries["name"] = name_entry

        tk.Label(self.frame, text="Тип оборудования", font=("Arial", 11)).grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.equipment_type_var = tk.StringVar()
        equipment_type_combo = ttk.Combobox(
            self.frame,
            textvariable=self.equipment_type_var,
            values=self.catalog.get_equipment_type_names(),
            state="readonly",
            width=29,
            font=("Arial", 11)
        )
        equipment_type_combo.grid(row=1, column=1, sticky="w", pady=4)
        equipment_type_combo.bind("<<ComboboxSelected>>", self.on_equipment_type_changed)
        self.entries["equipment_type_name"] = equipment_type_combo

        fields = [
            ("quantity", "Количество", "normal"),
            ("qy_kw", "Qy, кВт", "normal"),
            ("ka_w_per_kw", "Ka, Вт/кВт", "readonly"),
            ("kz", "Kз", "readonly"),
            ("width_mm", "Ширина, мм", "normal"),
            ("depth_mm", "Глубина, мм", "normal"),
        ]

        start_row = 2
        for i, (key, label, state) in enumerate(fields, start=start_row):
            tk.Label(self.frame, text=label, font=("Arial", 11)).grid(
                row=i, column=0, sticky="w", padx=(0, 10), pady=4
            )
            entry = tk.Entry(self.frame, width=20, font=("Arial", 11))
            entry.grid(row=i, column=1, sticky="w", pady=4)

            if state == "readonly":
                entry.config(state="readonly")

            self.entries[key] = entry

        position_row = start_row + len(fields)
        tk.Label(self.frame, text="Положение", font=("Arial", 11)).grid(
            row=position_row, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.position_var = tk.StringVar()
        position_combo = ttk.Combobox(
            self.frame,
            textvariable=self.position_var,
            values=self.catalog.get_position_names(),
            state="readonly",
            width=29,
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

        if data:
            self.load_data(data)
        else:
            self.apply_room_category_kz()

    def set_entry_value(self, key, value):
        entry = self.entries[key]
        current_state = entry.cget("state")

        if current_state == "readonly":
            entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, str(value))
            entry.config(state="readonly")
        else:
            entry.delete(0, tk.END)
            entry.insert(0, str(value))

    def load_data(self, data):
        self.entries["name"].delete(0, tk.END)
        self.entries["name"].insert(0, str(data.get("name", "")))

        type_id = str(data.get("equipment_type_id", ""))
        type_row = self.catalog.get_equipment_type_by_id(type_id)
        if type_row and type_row.get("type_name"):
            self.equipment_type_var.set(str(type_row["type_name"]))

        self.set_entry_value("quantity", data.get("quantity", ""))
        self.set_entry_value("qy_kw", data.get("qy_kw", ""))
        self.set_entry_value("ka_w_per_kw", data.get("ka_w_per_kw", ""))
        self.set_entry_value("kz", data.get("kz", ""))
        self.set_entry_value("width_mm", data.get("width_mm", ""))
        self.set_entry_value("depth_mm", data.get("depth_mm", ""))
        self.position_var.set(str(data.get("position", "")))

        if str(data.get("kz", "")).strip() == "":
            self.apply_room_category_kz()

    def on_equipment_type_changed(self, event=None):
        selected_name = self.equipment_type_var.get().strip()
        row = self.catalog.get_equipment_type_by_name(selected_name)
        if not row:
            return

        if row.get("default_qy_kw") not in (None, ""):
            self.set_entry_value("qy_kw", row["default_qy_kw"])

        if row.get("ka_w_per_kw") not in (None, ""):
            self.set_entry_value("ka_w_per_kw", row["ka_w_per_kw"])

    def apply_room_category_kz(self):
        selected_category = self.room_block.room_category_var.get().strip()
        if not selected_category:
            return

        defaults = self.catalog.get_room_category_defaults(selected_category)
        if not defaults:
            return

        kz_default = defaults.get("kz_default")
        if kz_default not in (None, ""):
            self.set_entry_value("kz", kz_default)

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
                if widget.cget("state") == "readonly":
                    current_state = widget.cget("state")
                    widget.config(state="normal", bg="white")
                    widget.config(state=current_state)
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
            if widget.cget("state") == "readonly":
                current_state = widget.cget("state")
                widget.config(state="normal", bg="#ffb3b3")
                widget.config(state=current_state)
            else:
                widget.config(bg="#ffb3b3")

    def validate(self):
        self.reset_highlight()
        ok = True

        required_keys = [
            "name",
            "equipment_type_name",
            "quantity",
            "qy_kw",
            "ka_w_per_kw",
            "kz",
            "position",
            "width_mm",
            "depth_mm",
        ]

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

        for key in ["qy_kw", "ka_w_per_kw", "kz", "width_mm", "depth_mm"]:
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
        equipment_type_name = self.get_widget_value("equipment_type_name")
        type_row = self.catalog.get_equipment_type_by_name(equipment_type_name)
        type_id = str(type_row["type_id"]) if type_row else ""

        return {
            "name": self.get_widget_value("name"),
            "equipment_type_id": type_id,
            "quantity": int(self.get_widget_value("quantity")),
            "qy_kw": float(self.get_widget_value("qy_kw")),
            "ka_w_per_kw": float(self.get_widget_value("ka_w_per_kw")),
            "kz": float(self.get_widget_value("kz")),
            "position": self.get_widget_value("position"),
            "width_mm": float(self.get_widget_value("width_mm")),
            "depth_mm": float(self.get_widget_value("depth_mm")),
            "room_name": self.room_block.name_entry.get().strip(),
        }


class RoomBlock:
    def __init__(self, parent, app, number, catalog: EquipmentCatalog, data=None):
        self.parent = parent
        self.app = app
        self.catalog = catalog
        self.number = number
        self.equipment_blocks = []

        self.frame = tk.Frame(parent, bd=3, relief="solid", padx=10, pady=10)
        self.frame.pack(anchor="w", fill="x", padx=20, pady=12)

        top_row = tk.Frame(self.frame)
        top_row.pack(anchor="w", fill="x", pady=(0, 10))

        default_name = f"Помещение{number}"
        if data and data.get("room_name"):
            default_name = data["room_name"]

        tk.Label(top_row, text="Имя помещения", font=("Arial", 11)).pack(side="left")
        self.name_var = tk.StringVar(value=default_name)
        self.name_entry = tk.Entry(top_row, textvariable=self.name_var, width=24, font=("Arial", 11))
        self.name_entry.pack(side="left", padx=(10, 15))

        tk.Label(top_row, text="Категория помещения", font=("Arial", 11)).pack(side="left")
        self.room_category_var = tk.StringVar()
        self.room_category_combo = ttk.Combobox(
            top_row,
            textvariable=self.room_category_var,
            values=self.catalog.get_room_categories(),
            state="readonly",
            width=22,
            font=("Arial", 11)
        )
        self.room_category_combo.pack(side="left", padx=(10, 15))
        self.room_category_combo.bind("<<ComboboxSelected>>", self.on_room_category_changed)

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

        if data:
            if data.get("room_category"):
                self.room_category_var.set(str(data["room_category"]))

            for eq_data in data.get("equipment", []):
                self.add_equipment(eq_data)

            self.apply_kz_to_all_equipment()

    def on_room_category_changed(self, event=None):
        self.apply_kz_to_all_equipment()

    def apply_kz_to_all_equipment(self):
        selected_category = self.room_category_var.get().strip()
        defaults = self.catalog.get_room_category_defaults(selected_category)
        if not defaults:
            return

        kz_default = defaults.get("kz_default")
        if kz_default in (None, ""):
            return

        for eq in self.equipment_blocks:
            eq.set_entry_value("kz", kz_default)

    def delete_self(self):
        if messagebox.askyesno("Подтверждение", "Удалить это помещение со всем оборудованием?"):
            self.frame.destroy()
            self.app.remove_room(self)

    def add_equipment(self, data=None):
        equipment_number = len(self.equipment_blocks) + 1
        block = EquipmentBlock(self.equipment_container, self, equipment_number, self.catalog, data=data)
        self.equipment_blocks.append(block)
        block.apply_room_category_kz()

    def remove_equipment(self, equipment_block):
        if equipment_block in self.equipment_blocks:
            self.equipment_blocks.remove(equipment_block)

    def reset_highlight(self):
        self.name_entry.config(bg="white")
        self.add_equipment_button.config(bg="SystemButtonFace")
        try:
            self.room_category_combo.configure(style="TCombobox")
        except Exception:
            pass

        for eq in self.equipment_blocks:
            eq.reset_highlight()

    def validate(self):
        self.reset_highlight()
        ok = True

        if not self.name_entry.get().strip():
            self.name_entry.config(bg="#ffb3b3")
            ok = False

        if not self.room_category_var.get().strip():
            try:
                self.room_category_combo.configure(style="Invalid.TCombobox")
            except Exception:
                pass
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
            "room_category": self.room_category_var.get().strip(),
            "equipment": [eq.get_data() for eq in self.equipment_blocks]
        }


class EditorWindow:
    def __init__(self, root, catalog: EquipmentCatalog, loaded_data=None):
        self.root = root
        self.catalog = catalog
        self.root.title("Редактор исходных данных")
        self.root.geometry("1220x820")
        self.root.minsize(980, 680)

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

        self.calculate_button = tk.Button(
            self.bottom_frame,
            text="Рассчитать",
            font=("Arial", 12),
            command=self.calculate_and_export
        )
        self.calculate_button.pack(side="right", padx=(10, 20))

        self.save_json_button = tk.Button(
            self.bottom_frame,
            text="Сохранить исходные данные",
            font=("Arial", 12),
            command=self.save_json
        )
        self.save_json_button.pack(side="right", padx=(10, 0))

        self.cancel_button = tk.Button(
            self.bottom_frame,
            text="Отмена",
            font=("Arial", 12),
            command=self.cancel_and_return
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
            ("k_empirical", "Эмпирический коэффициент k"),
            ("z_m", "Расстояние до зонта z, м"),
            ("ko", "Коэффициент эффективности Ko"),
            ("a", "Поправочный коэффициент a"),
        ]

        for i, (key, label) in enumerate(fields, start=1):
            tk.Label(block, text=label, font=("Arial", 11)).grid(
                row=i, column=0, sticky="w", padx=(0, 10), pady=4
            )
            entry = tk.Entry(block, width=22, font=("Arial", 11))
            entry.insert(0, str(DEFAULT_CONSTANTS[key]))
            entry.grid(row=i, column=1, sticky="w", pady=4)
            self.constants_entries[key] = entry

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_room(self, data=None):
        room_number = len(self.room_blocks) + 1
        block = RoomBlock(self.rooms_container, self, room_number, self.catalog, data=data)
        self.room_blocks.append(block)

    def remove_room(self, room_block):
        if room_block in self.room_blocks:
            self.room_blocks.remove(room_block)

    def load_into_form(self, loaded_data):
        constants = loaded_data.get("constants", {})
        for key, entry in self.constants_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, str(constants.get(key, DEFAULT_CONSTANTS[key])))

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
        data = {
            "constants": {
                key: float(entry.get().strip())
                for key, entry in self.constants_entries.items()
            },
            "rooms": [room.get_data() for room in self.room_blocks]
        }

        validated = ProjectInput(**data)
        return validated

    def save_json(self):
        if not self.validate_all():
            messagebox.showerror("Ошибка", "Не все параметры заполнены или заполнены неверно!")
            return False

        try:
            project = self.collect_data()
            data = project.model_dump()
        except Exception as e:
            messagebox.showerror("Ошибка в данных", str(e))
            return False

        file_path = filedialog.asksaveasfilename(
            title="Сохранить исходные данные",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )

        if not file_path:
            return False

        try:
            save_input_data(data, file_path)
            messagebox.showinfo("Успех", f"Исходные данные сохранены:\n{file_path}")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
            return False

    def calculate_and_export(self):
        if not self.validate_all():
            messagebox.showerror("Ошибка", "Не все параметры заполнены или заполнены неверно!")
            return

        try:
            project = self.collect_data()
        except Exception as e:
            messagebox.showerror("Ошибка в данных", str(e))
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить расчётный Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:
            return

        try:
            export_project_to_excel(project, self.catalog, file_path)
            messagebox.showinfo("Успех", f"Расчётный файл создан:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка расчёта", f"Не удалось сформировать расчётный Excel:\n{e}")

    def cancel_and_return(self):
        answer = messagebox.askyesnocancel("Выход", "Сохранить исходные данные?")

        if answer is None:
            return

        if answer is True:
            saved = self.save_json()
            if not saved:
                return

        self.open_start_window()

    def open_start_window(self):
        self.root.destroy()
        new_root = tk.Tk()
        StartWindow(new_root, self.catalog)
        new_root.mainloop()


class StartWindow:
    def __init__(self, root, catalog: EquipmentCatalog):
        self.root = root
        self.catalog = catalog

        self.root.title("Модуль исходных данных")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)

        container = tk.Frame(root)
        container.place(relx=0.5, rely=0.5, anchor="center")

        self.new_button = tk.Button(
            container,
            text="Создать исходные данные",
            font=("Arial", 16),
            width=26,
            height=2,
            command=self.create_new
        )
        self.new_button.pack(pady=15)

        self.load_button = tk.Button(
            container,
            text="Загрузить исходные данные",
            font=("Arial", 16),
            width=26,
            height=2,
            command=self.load_existing
        )
        self.load_button.pack(pady=15)

    def create_new(self):
        self.root.destroy()
        new_root = tk.Tk()
        EditorWindow(new_root, self.catalog)
        new_root.mainloop()

    def load_existing(self):
        file_path = filedialog.askopenfilename(
            title="Выбрать JSON-файл",
            filetypes=[("JSON files", "*.json")]
        )

        if not file_path:
            return

        try:
            data = load_input_data(file_path)
            ProjectInput(**data)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить исходные данные:\n{e}")
            return

        self.root.destroy()
        new_root = tk.Tk()
        EditorWindow(new_root, self.catalog, loaded_data=data)
        new_root.mainloop()