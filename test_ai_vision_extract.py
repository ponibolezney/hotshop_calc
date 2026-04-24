import json
import os
import sys

from app.catalog import EquipmentCatalog
from app.pdf_to_images import render_pdf_pages_to_images
from app.ai_vision_extractor import extract_project_json_from_images
from app.storage import save_input_data


CATALOG_PATH = os.path.join("data", "equipment_catalog.xlsx")
OUTPUT_PATH = os.path.join("data", "ai_extracted_input.json")


def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print("python test_ai_vision_extract.py путь_к_pdf")
        return

    pdf_path = sys.argv[1]

    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    print("Рендерю PDF в картинки...")
    image_paths = render_pdf_pages_to_images(pdf_path, dpi=180)

    # Для первого теста берём только первую страницу.
    # Потом увеличим до нескольких страниц.
    image_paths = image_paths[:1]

    print("Отправляю изображения в ИИ...")
    data = extract_project_json_from_images(image_paths, catalog)

    save_input_data(data, OUTPUT_PATH)

    print(f"Готово. JSON сохранён:")
    print(OUTPUT_PATH)

    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()