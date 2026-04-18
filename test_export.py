import os

from app.catalog import EquipmentCatalog
from app.schemas import ProjectInput
from app.storage import load_input_data
from app.excel_export import export_project_to_excel


CATALOG_PATH = os.path.join("data", "equipment_catalog.xlsx")
INPUT_PATH = os.path.join("data", "input.json")
OUTPUT_PATH = os.path.join("data", "hotshop_result.xlsx")


def main():
    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    raw_data = load_input_data(INPUT_PATH)
    project = ProjectInput(**raw_data)

    export_project_to_excel(project, catalog, OUTPUT_PATH)
    print(f"Excel-файл создан: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()