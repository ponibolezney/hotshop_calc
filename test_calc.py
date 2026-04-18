import os

from app.catalog import EquipmentCatalog
from app.schemas import ProjectInput
from app.storage import load_input_data
from app.calculator_hotshop import calculate_project, result_to_dict


CATALOG_PATH = os.path.join("data", "equipment_catalog.xlsx")
INPUT_PATH = os.path.join("data", "input.json")


def main():
    catalog = EquipmentCatalog(CATALOG_PATH)
    catalog.load()

    raw_data = load_input_data(INPUT_PATH)
    project = ProjectInput(**raw_data)

    result = calculate_project(project, catalog)
    result_dict = result_to_dict(result)

    for room in result_dict["room_results"]:
        print("=" * 80)
        print("Помещение:", room["room_name"])
        print("Категория:", room["room_category"])
        print("Итого по помещению Li, м3/ч:", round(room["room_total_li_m3h"], 2))
        print()

        for eq in room["equipment_results"]:
            print(f"Оборудование: {eq['equipment_name']}")
            print(f"  Qк, кВт:  {eq['qk_kw']:.4f}")
            print(f"  D, м:     {eq['d_m']:.4f}")
            print(f"  r:        {eq['r']:.4f}")
            print(f"  Lк, м3/ч: {eq['lk_m3h']:.2f}")
            print(f"  Li, м3/ч: {eq['li_m3h']:.2f}")
            print()

    print("=" * 80)
    print("ИТОГО ПО ПРОЕКТУ Li, м3/ч:", round(result_dict["project_total_li_m3h"], 2))


if __name__ == "__main__":
    main()