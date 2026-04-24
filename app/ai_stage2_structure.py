import json
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

from app.catalog import EquipmentCatalog


STRUCTURED_EXTRACTION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "rooms": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "room_name": {"type": "string"},
                    "room_category": {"type": "string"},
                    "equipment": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "raw_name": {"type": "string"},
                                "clean_name": {"type": "string"},
                                "probable_type_id": {"type": "string"},
                                "quantity": {"type": "integer"},
                                "source_text": {"type": "string"},
                                "confidence": {"type": "number"},
                                "needs_review": {"type": "boolean"}
                            },
                            "required": [
                                "raw_name",
                                "clean_name",
                                "probable_type_id",
                                "quantity",
                                "source_text",
                                "confidence",
                                "needs_review"
                            ]
                        }
                    }
                },
                "required": ["room_name", "room_category", "equipment"]
            }
        }
    },
    "required": ["rooms"]
}


def build_catalog_text(catalog: EquipmentCatalog) -> str:
    lines = []
    for row in catalog.equipment_types:
        lines.append(
            f"- type_id={row.get('type_id')}; type_name={row.get('type_name')}"
        )

    categories = []
    for row in catalog.room_categories:
        categories.append(str(row.get("room_category")))

    return (
        "ДОСТУПНЫЕ ТИПЫ ОБОРУДОВАНИЯ:\n"
        + "\n".join(lines)
        + "\n\nДОСТУПНЫЕ КАТЕГОРИИ ПОМЕЩЕНИЙ:\n"
        + "\n".join(categories)
    )


def structure_raw_plan_data(raw_data: dict, catalog: EquipmentCatalog) -> dict:
    load_dotenv()

    api_key = os.getenv("OPENROUTER_API_KEY")
    model = os.getenv("OPENROUTER_TEXT_MODEL", "openai/gpt-5.4-nano")

    if not api_key:
        raise RuntimeError("Не найден OPENROUTER_API_KEY в .env")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )

    catalog_text = build_catalog_text(catalog)

    system_prompt = """
Ты инженерный ассистент для структурирования списка оборудования.

На входе у тебя грязный JSON после vision-OCR: помещения и сырые подписи.
Твоя задача — превратить его в аккуратный список оборудования.

Правила:
1. Не добавляй оборудование, которого нет в raw_labels.
2. Не используй справочник как список найденного оборудования.
3. Справочник нужен только для выбора probable_type_id.
4. Если не уверен в типе — выбери ближайший probable_type_id, но needs_review=true.
5. Если подпись не является оборудованием для расчёта горячего цеха, можешь пропустить её.
6. Мойки, весы, принтеры, шкафы уборочного инвентаря, стеллажи обычно не являются тепловым оборудованием.
7. Плиты, печи, пароконвектоматы, фритюрницы, тепловые столы, варочные ванны — важные кандидаты.
8. quantity ставь 1, если по подписи не видно количество.
9. source_text должен быть точным фрагментом из raw_labels.
10. Никакого текста вне JSON.
"""

    user_prompt = {
        "catalog": catalog_text,
        "raw_data": raw_data
    }

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": json.dumps(user_prompt, ensure_ascii=False)},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "structured_plan_extraction",
                "strict": True,
                "schema": STRUCTURED_EXTRACTION_SCHEMA,
            },
        },
        temperature=0,
    )

    content = response.choices[0].message.content
    data = json.loads(content)

    Path("data/debug_stage2_structured.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return data