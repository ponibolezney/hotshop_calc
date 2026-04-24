import json
import os
import sys

from app.catalog import EquipmentCatalog
from app.pdf_to_images import render_pdf_pages_to_images
from app.ai_stage1_vision import extract_raw_plan_data
from app.ai_stage2_structure import structure_raw_plan_data
from app.ai_normalizer import normalize_structured_to_project_input
from app.storage import save_input_data


CATALOG_PATH = os.path.join("data", "equipment_catalog.xlsx")
OUTPUT_PATH = os.path.join("data", "ai_extracted_input.json")


def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print("python test_ai_pipeline.py путь_к_pdf")
        return

    pdf_path = sys.argv[1]

    max_pages = int(os.getenv("AI_MAX_PAGES", "1"))
    dpi = int(os.getenv("AI_PDF_DPI", "180"))

    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    print("1/4 Рендерю PDF в картинки...")
    image_paths = render_pdf_pages_to_images(pdf_path, dpi=dpi)
    image_paths = image_paths[:max_pages]

    print("2/4 Vision-модель вытаскивает грязный список...")
    raw_data = extract_raw_plan_data(image_paths)

    print("3/4 Text-модель структурирует оборудование...")
    structured_data = structure_raw_plan_data(raw_data, catalog)

    print("4/4 Python нормализует в ProjectInput...")
    project_input = normalize_structured_to_project_input(structured_data, catalog)

    save_input_data(project_input, OUTPUT_PATH)

    print(f"Готово. JSON сохранён:")
    print(OUTPUT_PATH)

    print(json.dumps(project_input, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()