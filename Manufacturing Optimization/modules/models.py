from dataclasses import dataclass, field
from datetime import datetime, date
from typing import List


@dataclass
class Phase:
    name: str
    planned_hours: float
    completed_hours: float = 0.0
    parallel_group: int = 0
    equipment_id: str = ""
    assigned_employee: str = ""


@dataclass
class Product:
    product_id: str
    part_number: str = ""
    quantity: int = 1
    produced_qty: int = 0
    unit_weight_g: float = 0.0
    phases: List[Phase] = field(default_factory=list)


@dataclass
class Equipment:
    equipment_id: str
    category: str = ""
    total_count: int = 1
    available_count: int = 1
    shift_template_name: str = ""


@dataclass
class ShiftDayPlan:
    shift_count: int = 1
    hours_per_shift: float = 8.0

    def total_hours(self) -> float:
        return max(0, self.shift_count) * max(0.0, self.hours_per_shift)


@dataclass
class ShiftTemplate:
    name: str
    week_plan: List[ShiftDayPlan] = field(default_factory=list)

    def hours_for_weekday(self, weekday: int) -> float:
        if weekday < 0:
            return 0.0
        if len(self.week_plan) < 7:
            return 0.0
        if weekday >= len(self.week_plan):
            return 0.0
        return self.week_plan[weekday].total_hours()


@dataclass
class Event:
    day: date
    hours_lost: float
    reason: str
    remark: str = ""


@dataclass
class CapacityAdjustment:
    day: date
    extra_hours: float
    reason: str
    equipment_ids: List[str] = field(default_factory=list)


@dataclass
class DefectRecord:
    product_id: str
    count: int
    category: str
    detail: str = ""
    timestamp: datetime = field(default_factory=datetime.now)


@dataclass
class LogEntry:
    timestamp: datetime
    user: str
    content: str
    order_id: str = ""


@dataclass
class MemoEntry:
    day: date
    user: str
    content: str


@dataclass
class Order:
    order_id: str
    start_dt: datetime
    order_date: date = field(default_factory=date.today)
    customer_code: str = ""
    shipping_method: str = ""
    due_date: date = field(default_factory=date.today)
    products: List[Product] = field(default_factory=list)
    events: List[Event] = field(default_factory=list)
    adjustments: List[CapacityAdjustment] = field(default_factory=list)
    defects: List[DefectRecord] = field(default_factory=list)
    equipment: List[Equipment] = field(default_factory=list)
    employees: List[str] = field(default_factory=list)
    logs: List[LogEntry] = field(default_factory=list)


@dataclass
class UserAccount:
    username: str
    password: str
