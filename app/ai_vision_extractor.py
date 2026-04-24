import base64
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

from app.ai_schema import AI_EXTRACTION_JSON_SCHEMA
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


def encode_image_base64(image_path: str) -> str:
    path = Path(image_path)
    data = path.read_bytes()
    encoded = base64.b64encode(data).decode("utf-8")
    return f"data:image/png;base64,{encoded}"


def build_catalog_text(catalog: EquipmentCatalog) -> str:
    equipment_lines = []
    for row in catalog.equipment_types:
        equipment_lines.append(
            f"- type_id={row.get('type_id')}; "
            f"type_name={row.get('type_name')}; "
            f"default_qy_kw={row.get('default_qy_kw')}; "
            f"ka_w_per_kw={row.get('ka_w_per_kw')}"
        )

    room_lines = []
    for row in catalog.room_categories:
        room_lines.append(
            f"- room_category={row.get('room_category')}; "
            f"kz_default={row.get('kz_default')}"
        )

    position_lines = []
    for row in catalog.position_coefficients:
        position_lines.append(
            f"- position_name={row.get('position_name')}; "
            f"r_value={row.get('r_value')}"
        )

    return (
        "СПРАВОЧНИК ТИПОВ ОБОРУДОВАНИЯ:\n"
        + "\n".join(equipment_lines)
        + "\n\nКАТЕГОРИИ ПОМЕЩЕНИЙ:\n"
        + "\n".join(room_lines)
        + "\n\nДОПУСТИМЫЕ ПОЛОЖЕНИЯ ОБОРУДОВАНИЯ:\n"
        + "\n".join(position_lines)
    )


def normalize_ai_data_to_project_input(ai_data: dict, catalog: EquipmentCatalog) -> dict:
    """
    Превращает расширенный AI JSON в обычный ProjectInput JSON.
    Поля source_text/confidence/needs_review выкидываются,
    потому что текущий UI их пока не знает.
    """
    result = {
        "constants": ai_data.get("constants", DEFAULT_CONSTANTS),
        "rooms": []
    }

    for room in ai_data.get("rooms", []):
        room_name = room.get("room_name", "").strip()
        room_category = room.get("room_category", "").strip()

        if not room_name:
            room_name = "Неопределённое помещение"

        if not room_category:
            room_category = "Горячий цех"

        new_room = {
            "room_name": room_name,
            "room_category": room_category,
            "equipment": []
        }

        for eq in room.get("equipment", []):
            source_text = str(eq.get("source_text", "")).strip()
            name = str(eq.get("name", "")).strip()
            type_id = str(eq.get("equipment_type_id", "")).strip()

            # Главный фильтр против галлюцинаций:
            # если нет источника с PDF или имени — не включаем оборудование
            if not source_text or not name:
                continue

            type_row = catalog.get_equipment_type_by_id(type_id)
            if not type_row:
                # если модель дала неизвестный type_id — пропускаем, чтобы не ломать расчёт
                continue

            qy_kw = eq.get("qy_kw", 0)
            ka_w_per_kw = type_row.get("ka_w_per_kw", 0)
            kz = eq.get("kz", 0)

            # Kz надёжнее брать из категории помещения
            room_defaults = catalog.get_room_category_defaults(room_category)
            if room_defaults and room_defaults.get("kz_default") not in (None, ""):
                kz = room_defaults.get("kz_default")

            # Qy: если модель не нашла мощность, берём дефолт из справочника
            # но это всё равно должно проверяться человеком в UI
            if qy_kw in (None, "", 0):
                qy_kw = type_row.get("default_qy_kw", 0)

            new_room["equipment"].append({
                "name": name,
                "equipment_type_id": type_id,
                "quantity": int(eq.get("quantity", 1) or 1),
                "qy_kw": float(qy_kw or 0),
                "ka_w_per_kw": float(ka_w_per_kw or 0),
                "kz": float(kz or 0),
                "position": str(eq.get("position", "") or ""),
                "width_mm": float(eq.get("width_mm", 0) or 0),
                "depth_mm": float(eq.get("depth_mm", 0) or 0),
                "room_name": room_name,
            })

        if new_room["equipment"]:
            result["rooms"].append(new_room)

    return result


def extract_project_json_from_images(image_paths: list[str], catalog: EquipmentCatalog) -> dict:
    load_dotenv()

    api_key = os.getenv("OPENROUTER_API_KEY")
    model = os.getenv("OPENROUTER_VISION_MODEL", "openai/gpt-4o-mini")

    if not api_key:
        raise RuntimeError("Не найден OPENROUTER_API_KEY в .env")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )

    catalog_text = build_catalog_text(catalog)

    system_prompt = """
Ты инженерный ассистент для подготовки ЧЕРНОВЫХ исходных данных расчёта вентиляции горячего цеха.

Важнейшее правило:
ТЫ НЕ ДОЛЖЕН ДОБАВЛЯТЬ ОБОРУДОВАНИЕ, ЕСЛИ ЕГО НАЗВАНИЕ ИЛИ МАРКИРОВКА НЕ ВИДНЫ НА ИЗОБРАЖЕНИИ PDF.

Твоя задача:
1. Проанализировать изображения страниц PDF.
2. Найти только явно подписанные помещения и явно подписанное кухонное оборудование.
3. Вернуть строго JSON по заданной схеме.
4. Не выполнять расчёт.
5. Не придумывать оборудование из справочника.
6. Справочник используется только для выбора equipment_type_id и табличных коэффициентов.
7. Запрещено использовать примеры из справочника как найденное оборудование.
8. Если на PDF написано "Пароконвектомат" — можно добавить "Пароконвектомат".
9. Если на PDF написано "Плита 70/40" — можно добавить "Плита 70/40".
10. Если на PDF НЕ написано "JOSPER" — запрещено писать "JOSPER".
11. Если мощность явно указана на PDF рядом с оборудованием или в его подписи, поставь qy_kw равным этой мощности.
12. Если мощность явно НЕ указана на PDF, поставь qy_kw = 0. Не бери qy_kw из справочника.
13. ka_w_per_kw = 0 и kz = 0. Эти значения НЕ нужно определять. Они будут заполнены программой после выбора equipment_type_id.
14. Если размеры явно не указаны рядом с оборудованием, поставь width_mm = 0 и depth_mm = 0.
15. Если положение оборудования относительно стены неочевидно, поставь position = "".
16. Для каждой позиции equipment обязательно заполни source_text — точный фрагмент текста с PDF, из которого ты сделал вывод.
17. Если source_text невозможно указать — не добавляй это оборудование.
18. confidence ставь от 0 до 1.
19. needs_review ставь true почти всегда, если не видны мощность, размеры или положение.
20. room_category выбирай из списка категорий помещений.
21. constants всегда используй стандартные.
22. Никакого текста вне JSON.
"""

    content = [
        {
            "type": "text",
            "text": (
                "СТАНДАРТНЫЕ КОНСТАНТЫ:\n"
                + json.dumps(DEFAULT_CONSTANTS, ensure_ascii=False, indent=2)
                + "\n\n"
                + catalog_text
                + "\n\n"
                + "Проанализируй изображения и верни JSON. "
                + "Не используй справочник как список найденного оборудования. "
                + "Каждое найденное оборудование должно иметь source_text с PDF. "
                + "Не бери qy_kw, ka_w_per_kw и kz из справочника: qy_kw только если мощность явно видна на PDF, иначе 0; ka_w_per_kw и kz всегда 0."
            )
        }
    ]

    for image_path in image_paths:
        content.append({
            "type": "image_url",
            "image_url": {
                "url": encode_image_base64(image_path)
            }
        })

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": content},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "ai_extraction_result",
                "strict": True,
                "schema": AI_EXTRACTION_JSON_SCHEMA,
            },
        },
        temperature=0,
    )

    raw_content = response.choices[0].message.content

    try:
        ai_data = json.loads(raw_content)
    except Exception:
        Path("data/debug_ai_raw_response.txt").write_text(raw_content, encoding="utf-8")
        raise

    # Сохраняем расширенный AI JSON для отладки
    Path("data/debug_ai_full_output.json").write_text(
        json.dumps(ai_data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    project_data = normalize_ai_data_to_project_input(ai_data, catalog)

    # Сохраняем финальный JSON, который идёт в редактор
    Path("data/debug_ai_project_input.json").write_text(
        json.dumps(project_data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    validated = ProjectInput(**project_data)
    return validated.model_dump()