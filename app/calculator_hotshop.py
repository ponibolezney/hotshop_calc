from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any

from app.schemas import ProjectInput
from app.catalog import EquipmentCatalog


@dataclass
class EquipmentCalculationResult:
    room_name: str
    room_category: str

    equipment_name: str
    equipment_type_id: str
    quantity: int
    position: str

    width_mm: float
    depth_mm: float
    width_m: float
    depth_m: float

    qy_kw: float
    ka_w_per_kw: float
    kz: float

    k1: float
    k_empirical: float
    z_m: float
    ko: float
    a: float
    r: float

    qk_kw: float
    d_m: float
    lk_m3h: float
    li_m3h: float


@dataclass
class RoomCalculationResult:
    room_name: str
    room_category: str
    equipment_results: list[EquipmentCalculationResult]
    room_total_li_m3h: float


@dataclass
class ProjectCalculationResult:
    room_results: list[RoomCalculationResult]
    project_total_li_m3h: float


class HotshopCalculationError(Exception):
    pass


def mm_to_m(value_mm: float) -> float:
    return float(value_mm) / 1000.0


def calc_qk_kw(
    quantity: int,
    qy_kw: float,
    ka_w_per_kw: float,
    k1: float,
    kz: float,
) -> float:
    """
    Конвективные теплопоступления от оборудования, кВт.

    Формула подогнана под твой рабочий образец:
    Qк = quantity * Qy * Ka * K1 * Kз / 1000
    """
    return quantity * qy_kw * ka_w_per_kw * k1 * kz / 1000.0


def calc_equivalent_diameter_m(width_m: float, depth_m: float) -> float:
    """
    Эквивалентный диаметр горячей поверхности, м.

    D = 2 * A * B / (A + B)
    """
    if width_m <= 0 or depth_m <= 0:
        raise HotshopCalculationError("Ширина и глубина должны быть больше 0.")
    return 2.0 * width_m * depth_m / (width_m + depth_m)


def calc_lk_m3h(
    qk_kw: float,
    k_empirical: float,
    z_m: float,
    d_m: float,
    r: float,
) -> float:
    """
    Базовый расход местного отсоса, м3/ч.

    Lк = k * Qк^(1/3) * (z + 1.7*D)^(5/3) * r
    """
    if qk_kw <= 0:
        raise HotshopCalculationError("Qк должен быть больше 0.")
    if z_m <= 0:
        raise HotshopCalculationError("z должно быть больше 0.")
    if d_m <= 0:
        raise HotshopCalculationError("D должно быть больше 0.")
    if r <= 0:
        raise HotshopCalculationError("r должно быть больше 0.")

    return k_empirical * (qk_kw ** (1.0 / 3.0)) * ((z_m + 1.7 * d_m) ** (5.0 / 3.0)) * r


def calc_li_m3h(
    lk_m3h: float,
    a: float,
    ko: float,
) -> float:
    """
    Итоговый расход удаляемого воздуха от оборудования, м3/ч.

    Li = Lк * a / Ko
    """
    if lk_m3h <= 0:
        raise HotshopCalculationError("Lк должен быть больше 0.")
    if a <= 0:
        raise HotshopCalculationError("a должно быть больше 0.")
    if ko <= 0:
        raise HotshopCalculationError("Ko должно быть больше 0.")

    return lk_m3h * a / ko


def get_position_r_value(catalog: EquipmentCatalog, position_name: str) -> float:
    r_value = catalog.get_position_r(position_name)
    if r_value in (None, ""):
        raise HotshopCalculationError(
            f"Для положения '{position_name}' не найден коэффициент r в справочнике."
        )

    try:
        return float(r_value)
    except Exception as exc:
        raise HotshopCalculationError(
            f"Некорректное значение r='{r_value}' для положения '{position_name}'."
        ) from exc


def calculate_equipment(
    room_name: str,
    room_category: str,
    equipment: Any,
    constants: Any,
    catalog: EquipmentCatalog,
) -> EquipmentCalculationResult:
    width_m = mm_to_m(equipment.width_mm)
    depth_m = mm_to_m(equipment.depth_mm)

    r_value = get_position_r_value(catalog, equipment.position)

    qk_kw = calc_qk_kw(
        quantity=equipment.quantity,
        qy_kw=equipment.qy_kw,
        ka_w_per_kw=equipment.ka_w_per_kw,
        k1=constants.k1,
        kz=equipment.kz,
    )

    d_m = calc_equivalent_diameter_m(width_m, depth_m)

    lk_m3h = calc_lk_m3h(
        qk_kw=qk_kw,
        k_empirical=constants.k_empirical,
        z_m=constants.z_m,
        d_m=d_m,
        r=r_value,
    )

    li_m3h = calc_li_m3h(
        lk_m3h=lk_m3h,
        a=constants.a,
        ko=constants.ko,
    )

    return EquipmentCalculationResult(
        room_name=room_name,
        room_category=room_category,

        equipment_name=equipment.name,
        equipment_type_id=equipment.equipment_type_id,
        quantity=equipment.quantity,
        position=equipment.position,

        width_mm=equipment.width_mm,
        depth_mm=equipment.depth_mm,
        width_m=width_m,
        depth_m=depth_m,

        qy_kw=equipment.qy_kw,
        ka_w_per_kw=equipment.ka_w_per_kw,
        kz=equipment.kz,

        k1=constants.k1,
        k_empirical=constants.k_empirical,
        z_m=constants.z_m,
        ko=constants.ko,
        a=constants.a,
        r=r_value,

        qk_kw=qk_kw,
        d_m=d_m,
        lk_m3h=lk_m3h,
        li_m3h=li_m3h,
    )


def calculate_project(
    project: ProjectInput,
    catalog: EquipmentCatalog,
) -> ProjectCalculationResult:
    room_results: list[RoomCalculationResult] = []

    for room in project.rooms:
        equipment_results: list[EquipmentCalculationResult] = []

        for equipment in room.equipment:
            result = calculate_equipment(
                room_name=room.room_name,
                room_category=room.room_category,
                equipment=equipment,
                constants=project.constants,
                catalog=catalog,
            )
            equipment_results.append(result)

        room_total = sum(item.li_m3h for item in equipment_results)

        room_results.append(
            RoomCalculationResult(
                room_name=room.room_name,
                room_category=room.room_category,
                equipment_results=equipment_results,
                room_total_li_m3h=room_total,
            )
        )

    project_total = sum(room.room_total_li_m3h for room in room_results)

    return ProjectCalculationResult(
        room_results=room_results,
        project_total_li_m3h=project_total,
    )


def result_to_dict(result: ProjectCalculationResult) -> dict:
    return {
        "room_results": [
            {
                "room_name": room.room_name,
                "room_category": room.room_category,
                "room_total_li_m3h": room.room_total_li_m3h,
                "equipment_results": [asdict(eq) for eq in room.equipment_results],
            }
            for room in result.room_results
        ],
        "project_total_li_m3h": result.project_total_li_m3h,
    }