import base64
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI


RAW_EXTRACTION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "pages": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "page": {"type": "integer"},
                    "rooms": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "room_name": {"type": "string"},
                                "raw_labels": {
                                    "type": "array",
                                    "items": {"type": "string"}
                                },
                                "notes": {"type": "string"}
                            },
                            "required": ["room_name", "raw_labels", "notes"]
                        }
                    }
                },
                "required": ["page", "rooms"]
            }
        }
    },
    "required": ["pages"]
}


def encode_image_base64(image_path: str) -> str:
    data = Path(image_path).read_bytes()
    encoded = base64.b64encode(data).decode("utf-8")
    return f"data:image/png;base64,{encoded}"


def extract_raw_plan_data(image_paths: list[str]) -> dict:
    load_dotenv()

    api_key = os.getenv("OPENROUTER_API_KEY")
    model = os.getenv("OPENROUTER_VISION_MODEL", "openai/gpt-5.4-mini")
    fallback_model = os.getenv("OPENROUTER_FALLBACK_VISION_MODEL", "openai/gpt-5.4")

    if not api_key:
        raise RuntimeError("Не найден OPENROUTER_API_KEY в .env")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
    )

    system_prompt = """
Ты инженерный vision-OCR ассистент.

Твоя задача — НЕ делать расчёт и НЕ классифицировать оборудование.
Твоя задача — посмотреть план расстановки оборудования и выписать:
1. помещения;
2. все читаемые подписи внутри или рядом с каждым помещением;
3. только то, что реально видно на изображении.

Правила:
- Не используй справочники.
- Не придумывай оборудование.
- Не заменяй подписи на типовые названия.
- Если написано "Плита 70/40", так и пиши "Плита 70/40".
- Если написано "FRE10", так и пиши "FRE10".
- Если подпись непонятна, всё равно выпиши как видишь.
- Никаких расчётных коэффициентов.
- Никакого текста вне JSON.
"""

    content = [
        {
            "type": "text",
            "text": "Проанализируй изображения страниц PDF и верни грязный список помещений и подписей."
        }
    ]

    for image_path in image_paths:
        content.append({
            "type": "image_url",
            "image_url": {"url": encode_image_base64(image_path)}
        })

    def call_model(selected_model: str):
        return client.chat.completions.create(
            model=selected_model,
            messages=[
                {"role": "system", "content": system_prompt.strip()},
                {"role": "user", "content": content},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": {
                    "name": "raw_plan_extraction",
                    "strict": True,
                    "schema": RAW_EXTRACTION_SCHEMA,
                },
            },
            temperature=0,
        )

    try:
        response = call_model(model)
    except Exception:
        response = call_model(fallback_model)

    raw_content = response.choices[0].message.content
    data = json.loads(raw_content)

    Path("data/debug_stage1_raw.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return data