from datetime import datetime, timedelta, date
from typing import Dict, List

from .models import Order, Product, Phase, Equipment, CapacityAdjustment


class WorkCalendar:
    def __init__(self, shift_template=None):
        self.shift_template = shift_template

    def capacity_for_day(self, d: date) -> float:
        if not self.shift_template:
            return 0.0
        return self.shift_template.hours_for_weekday(d.weekday())


def _normalize_equipment_ids(items: List[str]) -> List[str]:
    seen = set()
    result: List[str] = []
    for item in items:
        text = (item or "").strip()
        if not text or text in ("-", "无需设备"):
            continue
        if text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def _split_equipment_ids(text: str) -> List[str]:
    if not text:
        return []
    if text.strip() in ("-", "无需设备"):
        return []
    return _normalize_equipment_ids(text.split(","))


def _format_equipment_ids(ids: List[str]) -> str:
    return ",".join(_normalize_equipment_ids(ids))


def _equipment_available_map_from_list(equipment: List[Equipment]) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for eq in equipment:
        if not eq.equipment_id:
            continue
        result[eq.equipment_id] = max(0, int(eq.available_count))
    return result


def _equipment_available_map(order: Order) -> Dict[str, int]:
    return _equipment_available_map_from_list(order.equipment)


def _phase_effective_hours(phase: Phase, quantity: int, equipment_map: Dict[str, int]) -> float:
    base = max(0.0, float(phase.planned_hours))
    hours = base
    equipment_ids = _split_equipment_ids(phase.equipment_id)
    if equipment_ids:
        available = sum(max(0, equipment_map.get(eq_id, 0)) for eq_id in equipment_ids)
        if available > 0:
            hours = hours / available
    return hours


def _phase_total_hours(phase: Phase, quantity: int) -> float:
    return max(0.0, float(phase.planned_hours))


def _phase_completion_ratio(phase: Phase, quantity: int) -> float:
    total = _phase_total_hours(phase, quantity)
    if total <= 0:
        return 0.0
    completed = max(0.0, float(phase.completed_hours))
    completed = min(completed, total)
    return completed / total


def _product_remaining_hours(product: Product, equipment_map: Dict[str, int]) -> float:
    total = 0.0
    parallel_groups: Dict[int, List[float]] = {}
    for phase in product.phases:
        hours = _phase_effective_hours(phase, product.quantity, equipment_map)
        ratio = _phase_completion_ratio(phase, product.quantity)
        remaining = hours * (1.0 - ratio)
        if phase.parallel_group > 0:
            parallel_groups.setdefault(phase.parallel_group, []).append(remaining)
        else:
            total += remaining
    for group_hours in parallel_groups.values():
        total += max(group_hours)
    return total


def _product_progress(product: Product, equipment_map: Dict[str, int]) -> float:
    total = 0.0
    done = 0.0
    for phase in product.phases:
        hours = _phase_effective_hours(phase, product.quantity, equipment_map)
        total += hours
        ratio = _phase_completion_ratio(phase, product.quantity)
        done += hours * ratio
    if total <= 0:
        return 0.0
    return min(done / total, 1.0)


def _product_quantity_progress(product: Product) -> float:
    qty = max(int(product.quantity), 0)
    if qty <= 0:
        return 0.0
    produced = max(0, int(product.produced_qty))
    if produced > qty:
        produced = qty
    return produced / qty


def compute_eta(order: Order, cal: WorkCalendar) -> Dict[str, object]:
    equipment_map = _equipment_available_map(order)
    remaining_hours = sum(_product_remaining_hours(p, equipment_map) for p in order.products)

    lost_map: Dict[date, float] = {}
    reason_map: Dict[date, List[str]] = {}
    for ev in order.events:
        lost_map[ev.day] = lost_map.get(ev.day, 0.0) + ev.hours_lost
        reason_map.setdefault(ev.day, []).append(f"{ev.reason}(-{ev.hours_lost:g}h)")

    extra_map: Dict[date, float] = {}
    extra_reason_map: Dict[date, List[str]] = {}
    for adj in order.adjustments:
        equipment_count = len(adj.equipment_ids) if adj.equipment_ids else 1
        extra_total = adj.extra_hours * max(equipment_count, 1)
        extra_map[adj.day] = extra_map.get(adj.day, 0.0) + extra_total
        eq_text = ""
        if adj.equipment_ids:
            eq_text = f"设备:{','.join(adj.equipment_ids)} "
        label = (
            f"{eq_text}{adj.reason}(+{adj.extra_hours:g}h)"
            if adj.reason
            else f"{eq_text}+{adj.extra_hours:g}h"
        )
        extra_reason_map.setdefault(adj.day, []).append(label.strip())

    explanation: List[str] = []
    if order.products:
        explanation.append("Product workload summary:")
        for p in order.products:
            hours = _product_remaining_hours(p, equipment_map)
            explanation.append(
                f"- {p.product_id} (PN={p.part_number or '-'} qty={p.quantity}): {hours:g}h"
            )
        explanation.append("")

    if remaining_hours <= 0:
        return {
            "eta_dt": order.start_dt,
            "remaining_hours": 0.0,
            "daily_capacity_map": {},
            "explanation": explanation + ["All phases completed. ETA equals start time."],
        }

    current_day = order.start_dt.date()
    hours_left = remaining_hours
    daily_capacity_map: Dict[date, float] = {}

    for _ in range(3650):
        base_cap = cal.capacity_for_day(current_day)
        lost = lost_map.get(current_day, 0.0)
        extra = extra_map.get(current_day, 0.0)
        cap = max(base_cap - lost + extra, 0.0)

        if base_cap > 0 or extra > 0 or lost > 0:
            daily_capacity_map[current_day] = cap

            if lost > 0 or extra > 0:
                parts: List[str] = []
                if lost > 0:
                    parts.append(f"- {lost:g}h")
                if extra > 0:
                    parts.append(f"+ {extra:g}h")
                change_text = " ".join(parts) if parts else ""
                reasons = reason_map.get(current_day, []) + extra_reason_map.get(current_day, [])
                explanation.append(
                    f"{current_day.isoformat()}: capacity {base_cap:g}h "
                    f"{change_text} "
                    f"=> {cap:g}h "
                    f"({', '.join(reasons)})"
                )

            if cap > 0:
                if hours_left <= cap:
                    finish_time = datetime.combine(current_day, datetime.min.time()).replace(hour=9, minute=0)
                    finish_time += timedelta(hours=hours_left)
                    return {
                        "eta_dt": finish_time,
                        "remaining_hours": remaining_hours,
                        "daily_capacity_map": daily_capacity_map,
                        "explanation": explanation or ["No blocking events."],
                    }
                hours_left -= cap
        current_day = current_day + timedelta(days=1)

    raise RuntimeError("ETA computation exceeded safe bounds.")
