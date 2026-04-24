import json
import os
import sys

from app.catalog import EquipmentCatalog
from app.pdf_extract import extract_pdf_text
from app.ai_extractor import extract_project_json_from_text
from app.storage import save_input_data


CATALOG_PATH = os.path.join("data", "equipment_catalog.xlsx")
OUTPUT_PATH = os.path.join("data", "ai_extracted_input.json")


def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print("python test_ai_extract.py путь_к_pdf")
        return

    pdf_path = sys.argv[1]

    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    print("Читаю PDF...")
    text = extract_pdf_text(pdf_path)

    print("Отправляю текст в ИИ...")
    data = extract_project_json_from_text(text, catalog)

    save_input_data(data, OUTPUT_PATH)

    print(f"Готово. JSON сохранён:")
    print(OUTPUT_PATH)

    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()