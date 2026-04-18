from pydantic import BaseModel, Field
from typing import List, Optional


class ProjectConstants(BaseModel):
    k1: float = Field(gt=0, description="Доля конвективных тепловыделений")
    k2: float = Field(gt=0, description="Коэффициент одновременности")
    k_empirical: float = Field(gt=0, description="Эмпирический коэффициент k")
    z_m: float = Field(gt=0, description="Расстояние от оборудования до зонта, м")
    ko: float = Field(gt=0, description="Коэффициент эффективности местного отсоса")
    a: float = Field(gt=0, description="Поправочный коэффициент на способ воздухораспределения")


class EquipmentItem(BaseModel):
    name: str
    equipment_type_id: str
    quantity: int = Field(ge=1)

    qy_kw: float = Field(gt=0, description="Установленная мощность оборудования, кВт")
    ka_w_per_kw: float = Field(gt=0, description="Удельные конвективные тепловыделения, Вт/кВт")
    kz: float = Field(gt=0, description="Коэффициент загрузки оборудования")

    position: str
    width_mm: float = Field(gt=0)
    depth_mm: float = Field(gt=0)

    room_name: Optional[str] = None


class RoomItem(BaseModel):
    room_name: str
    room_category: str
    equipment: List[EquipmentItem]


class ProjectInput(BaseModel):
    constants: ProjectConstants
    rooms: List[RoomItem]