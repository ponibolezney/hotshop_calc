from __future__ import annotations

from typing import Dict, List, Tuple
import xlsxwriter

from app.catalog import EquipmentCatalog
from app.schemas import ProjectInput


def _build_equipment_rows(project: ProjectInput) -> List[dict]:
    rows = []
    for room in project.rooms:
        for equipment in room.equipment:
            rows.append({
                "room_name": room.room_name,
                "room_category": room.room_category,
                "equipment_name": equipment.name,
                "equipment_type_id": equipment.equipment_type_id,
                "quantity": equipment.quantity,
                "qy_kw": equipment.qy_kw,
                "ka_w_per_kw": equipment.ka_w_per_kw,
                "kz": equipment.kz,
                "position": equipment.position,
                "width_mm": equipment.width_mm,
                "depth_mm": equipment.depth_mm,
            })
    return rows


def _catalog_maps(catalog: EquipmentCatalog) -> Tuple[Dict[str, float], Dict[str, float], Dict[str, str]]:
    position_r_map: Dict[str, float] = {}
    for row in catalog.position_coefficients:
        name = row.get("position_name")
        r_value = row.get("r_value")
        if name not in (None, "") and r_value not in (None, ""):
            position_r_map[str(name)] = float(r_value)

    room_kz_map: Dict[str, float] = {}
    for row in catalog.room_categories:
        category = row.get("room_category")
        kz = row.get("kz_default")
        if category not in (None, "") and kz not in (None, ""):
            room_kz_map[str(category)] = float(kz)

    equipment_name_map: Dict[str, str] = {}
    for row in catalog.equipment_types:
        type_id = row.get("type_id")
        type_name = row.get("type_name")
        if type_id not in (None, "") and type_name not in (None, ""):
            equipment_name_map[str(type_id)] = str(type_name)

    return position_r_map, room_kz_map, equipment_name_map


def export_project_to_excel(
    project: ProjectInput,
    catalog: EquipmentCatalog,
    filepath: str,
):
    workbook = xlsxwriter.Workbook(filepath)

    ws_input = workbook.add_worksheet("Исходные данные")
    ws_ref = workbook.add_worksheet("Справочники")
    ws_calc = workbook.add_worksheet("Расчёт")

    # ---------- Форматы ----------
    fmt_header = workbook.add_format({
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "text_wrap": True,
    })

    fmt_cell = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    fmt_text = workbook.add_format({
        "border": 1,
        "align": "left",
        "valign": "vcenter",
    })

    fmt_room = workbook.add_format({
        "border": 1,
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })

    fmt_number = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "0.00",
    })

    fmt_number3 = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "num_format": "0.000",
    })

    fmt_total = workbook.add_format({
        "border": 1,
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "num_format": "0.00",
    })

    # ---------- Лист: Исходные данные ----------
    ws_input.set_column("A:A", 24)
    ws_input.set_column("B:B", 22)
    ws_input.set_column("C:C", 20)
    ws_input.set_column("D:D", 26)
    ws_input.set_column("E:K", 16)

    # Константы проекта
    ws_input.merge_range("A1:B1", "Константы проекта", fmt_room)
    constant_rows = [
        ("K1", project.constants.k1),
        ("K2", project.constants.k2),
        ("Эмпирический коэффициент k", project.constants.k_empirical),
        ("Расстояние до зонта z, м", project.constants.z_m),
        ("Коэффициент эффективности Ko", project.constants.ko),
        ("Поправочный коэффициент a", project.constants.a),
    ]

    row = 1
    for label, value in constant_rows:
        ws_input.write(row, 0, label, fmt_text)
        ws_input.write(row, 1, value, fmt_number3)
        row += 1

    row += 1

    input_headers = [
        "Помещение",
        "Категория помещения",
        "Наименование",
        "Тип оборудования",
        "Количество",
        "Qy, кВт",
        "Ka, Вт/кВт",
        "Kз",
        "Положение",
        "Ширина, мм",
        "Глубина, мм",
    ]

    for col, header in enumerate(input_headers):
        ws_input.write(row, col, header, fmt_header)

    equipment_name_map = {
        str(r["type_id"]): str(r["type_name"])
        for r in catalog.equipment_types
        if r.get("type_id") not in (None, "") and r.get("type_name") not in (None, "")
    }

    input_start_row = row + 1
    current_row = input_start_row

    for room in project.rooms:
        room_start = current_row
        for eq in room.equipment:
            ws_input.write(current_row, 0, room.room_name, fmt_text)
            ws_input.write(current_row, 1, room.room_category, fmt_text)
            ws_input.write(current_row, 2, eq.name, fmt_text)
            ws_input.write(current_row, 3, equipment_name_map.get(eq.equipment_type_id, eq.equipment_type_id), fmt_text)
            ws_input.write(current_row, 4, eq.quantity, fmt_cell)
            ws_input.write(current_row, 5, eq.qy_kw, fmt_number3)
            ws_input.write(current_row, 6, eq.ka_w_per_kw, fmt_number3)
            ws_input.write(current_row, 7, eq.kz, fmt_number3)
            ws_input.write(current_row, 8, eq.position, fmt_text)
            ws_input.write(current_row, 9, eq.width_mm, fmt_number3)
            ws_input.write(current_row, 10, eq.depth_mm, fmt_number3)
            current_row += 1

        # визуальная группировка помещений
        if current_row > room_start:
            for col in range(len(input_headers)):
                ws_input.write_blank(current_row, col, None, fmt_cell)
            current_row += 1

    # ---------- Лист: Справочники ----------
    ws_ref.set_column("A:H", 24)

    # equipment_types
    ref_row = 0
    ws_ref.write(ref_row, 0, "equipment_types", fmt_room)
    ref_row += 1

    eq_headers = ["type_id", "type_name", "energy_source", "default_qy_kw", "ka_w_per_kw", "notes"]
    for col, header in enumerate(eq_headers):
        ws_ref.write(ref_row, col, header, fmt_header)
    ref_row += 1

    for row_data in catalog.equipment_types:
        for col, key in enumerate(eq_headers):
            value = row_data.get(key, "")
            if isinstance(value, (int, float)):
                ws_ref.write(ref_row, col, value, fmt_number3)
            else:
                ws_ref.write(ref_row, col, value, fmt_text)
        ref_row += 1

    ref_row += 2

    # room_categories
    ws_ref.write(ref_row, 0, "room_categories", fmt_room)
    ref_row += 1

    room_headers = ["room_category", "kz_default", "notes"]
    for col, header in enumerate(room_headers):
        ws_ref.write(ref_row, col, header, fmt_header)
    ref_row += 1

    room_categories_start_excel_row = ref_row + 1  # 1-based for Excel formulas

    room_categories_count = 0
    for row_data in catalog.room_categories:
        room_categories_count += 1
        for col, key in enumerate(room_headers):
            value = row_data.get(key, "")
            if isinstance(value, (int, float)):
                ws_ref.write(ref_row, col, value, fmt_number3)
            else:
                ws_ref.write(ref_row, col, value, fmt_text)
        ref_row += 1

    room_categories_end_excel_row = room_categories_start_excel_row + room_categories_count - 1

    ref_row += 2

    # position_coefficients
    ws_ref.write(ref_row, 0, "position_coefficients", fmt_room)
    ref_row += 1

    pos_headers = ["position_name", "r_value", "notes"]
    for col, header in enumerate(pos_headers):
        ws_ref.write(ref_row, col, header, fmt_header)
    ref_row += 1

    position_start_excel_row = ref_row + 1  # 1-based for Excel formulas

    position_count = 0
    for row_data in catalog.position_coefficients:
        position_count += 1
        for col, key in enumerate(pos_headers):
            value = row_data.get(key, "")
            if isinstance(value, (int, float)):
                ws_ref.write(ref_row, col, value, fmt_number3)
            else:
                ws_ref.write(ref_row, col, value, fmt_text)
        ref_row += 1

    position_end_excel_row = position_start_excel_row + position_count - 1

    # ---------- Лист: Расчёт ----------
    ws_calc.set_column("A:A", 22)
    ws_calc.set_column("B:B", 20)
    ws_calc.set_column("C:C", 28)
    ws_calc.set_column("D:D", 20)
    ws_calc.set_column("E:I", 14)
    ws_calc.set_column("J:P", 16)

    calc_headers = [
        "Помещение",              # A
        "Категория помещения",    # B
        "Наименование",           # C
        "Положение",              # D
        "Количество",             # E
        "Qy, кВт",                # F
        "Ka, Вт/кВт",             # G
        "Kз",                     # H
        "Ширина, м",              # I
        "Глубина, м",             # J
        "D, м",                   # K
        "r",                      # L
        "Qк, кВт",                # M
        "Lк, м³/ч",               # N
        "a",                      # O
        "Ko",                     # P
        "Li, м³/ч",               # Q
        "Проверка Kз из справ.",  # R
    ]

    for col, header in enumerate(calc_headers):
        ws_calc.write(0, col, header, fmt_header)

    equipment_rows = _build_equipment_rows(project)

    # Константы с листа "Исходные данные"
    # B2..B7 в Excel-координатах:
    const_k1 = "'Исходные данные'!$B$2"
    const_k2 = "'Исходные данные'!$B$3"
    const_k = "'Исходные данные'!$B$4"
    const_z = "'Исходные данные'!$B$5"
    const_ko = "'Исходные данные'!$B$6"
    const_a = "'Исходные данные'!$B$7"

    for idx, row_data in enumerate(equipment_rows, start=2):
        excel_row = idx

        width_m = float(row_data["width_mm"]) / 1000.0
        depth_m = float(row_data["depth_mm"]) / 1000.0

        ws_calc.write(f"A{excel_row}", row_data["room_name"], fmt_text)
        ws_calc.write(f"B{excel_row}", row_data["room_category"], fmt_text)
        ws_calc.write(f"C{excel_row}", row_data["equipment_name"], fmt_text)
        ws_calc.write(f"D{excel_row}", row_data["position"], fmt_text)
        ws_calc.write_number(f"E{excel_row}", row_data["quantity"], fmt_cell)
        ws_calc.write_number(f"F{excel_row}", row_data["qy_kw"], fmt_number3)
        ws_calc.write_number(f"G{excel_row}", row_data["ka_w_per_kw"], fmt_number3)
        ws_calc.write_number(f"H{excel_row}", row_data["kz"], fmt_number3)
        ws_calc.write_number(f"I{excel_row}", width_m, fmt_number3)
        ws_calc.write_number(f"J{excel_row}", depth_m, fmt_number3)

        # K = D = 2*A*B/(A+B)
        ws_calc.write_formula(
            f"K{excel_row}",
            f"=2*I{excel_row}*J{excel_row}/(I{excel_row}+J{excel_row})",
            fmt_number3
        )

        # L = r по положению
        ws_calc.write_formula(
            f"L{excel_row}",
            (
                f'=XLOOKUP(D{excel_row},'
                f"'Справочники'!$A${position_start_excel_row}:$A${position_end_excel_row},"
                f"'Справочники'!$B${position_start_excel_row}:$B${position_end_excel_row},"
                f'"")'
            ),
            fmt_number3
        )

        # M = Qк = quantity * Qy * Ka * K1 * Kз / 1000
        ws_calc.write_formula(
            f"M{excel_row}",
            f"=E{excel_row}*F{excel_row}*G{excel_row}*{const_k1}*H{excel_row}/1000",
            fmt_number3
        )

        # N = Lк = k * Qк^(1/3) * (z + 1.7*D)^(5/3) * r
        ws_calc.write_formula(
            f"N{excel_row}",
            f"={const_k}*(M{excel_row}^(1/3))*(({const_z}+1.7*K{excel_row})^(5/3))*L{excel_row}",
            fmt_number
        )

        # O = a
        ws_calc.write_formula(f"O{excel_row}", f"={const_a}", fmt_number3)

        # P = Ko
        ws_calc.write_formula(f"P{excel_row}", f"={const_ko}", fmt_number3)

        # Q = Li = Lк * a / Ko
        ws_calc.write_formula(
            f"Q{excel_row}",
            f"=N{excel_row}*O{excel_row}/P{excel_row}",
            fmt_number
        )

        # R = проверка Kз по категории помещения из справочника
        ws_calc.write_formula(
            f"R{excel_row}",
            (
                f'=XLOOKUP(B{excel_row},'
                f"'Справочники'!$A${room_categories_start_excel_row}:$A${room_categories_end_excel_row},"
                f"'Справочники'!$B${room_categories_start_excel_row}:$B${room_categories_end_excel_row},"
                f'"")'
            ),
            fmt_number3
        )

    total_row = len(equipment_rows) + 2
    ws_calc.write(total_row - 1, 15, "Итого по Li, м³/ч", fmt_total)
    ws_calc.write_formula(total_row - 1, 16, f"=SUM(Q2:Q{len(equipment_rows)+1})", fmt_total)
    
    workbook.close()