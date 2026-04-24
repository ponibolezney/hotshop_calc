import json
from pathlib import Path

from app.catalog import EquipmentCatalog
from app.schemas import ProjectInput


DEFAULT_CONSTANTS = {
    "k1": 0.5,
    "k2": 0.7,
    "k_empirical": 180.0,
    "z_m": 1.1,
    "ko": 0.8,
    "a": 1.25,
}


def normalize_structured_to_project_input(structured_data: dict, catalog: EquipmentCatalog) -> dict:
    result = {
        "constants": DEFAULT_CONSTANTS,
        "rooms": []
    }

    for room in structured_data.get("rooms", []):
        room_name = str(room.get("room_name", "")).strip() or "Неопределённое помещение"
        room_category = str(room.get("room_category", "")).strip() or "Горячий цех"

        room_defaults = catalog.get_room_category_defaults(room_category)
        kz = 0.0
        if room_defaults and room_defaults.get("kz_default") not in (None, ""):
            kz = float(room_defaults.get("kz_default"))

        new_room = {
            "room_name": room_name,
            "room_category": room_category,
            "equipment": []
        }

        for eq in room.get("equipment", []):
            type_id = str(eq.get("probable_type_id", "")).strip()
            type_row = catalog.get_equipment_type_by_id(type_id)

            if not type_row:
                continue

            qy_kw = float(type_row.get("default_qy_kw") or 0)
            ka_w_per_kw = float(type_row.get("ka_w_per_kw") or 0)

            new_room["equipment"].append({
                "name": str(eq.get("clean_name", "")).strip() or str(eq.get("raw_name", "")).strip(),
                "equipment_type_id": type_id,
                "quantity": int(eq.get("quantity", 1) or 1),
                "qy_kw": qy_kw,
                "ka_w_per_kw": ka_w_per_kw,
                "kz": kz,
                "position": "",
                "width_mm": 0.0,
                "depth_mm": 0.0,
                "room_name": room_name,
            })

        if new_room["equipment"]:
            result["rooms"].append(new_room)

    Path("data/debug_stage3_project_input.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    validated = ProjectInput(**result)
    return validated.model_dump()