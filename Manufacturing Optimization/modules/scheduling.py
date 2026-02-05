# -*- coding: utf-8 -*-
"""
生产交期优化系统 - 排程与交期计算
Manufacturing Optimization System - Scheduling and ETA Computation
"""

from datetime import datetime, timedelta, date
from typing import Dict, List, Optional

from .models import ShiftTemplate, Order
from .utils import equipment_available_map, product_remaining_hours


class WorkCalendar:
    """工作日历，用于计算每日产能"""
    
    def __init__(self, shift_template: Optional[ShiftTemplate] = None):
        self.shift_template = shift_template

    def capacity_for_day(self, d: date) -> float:
        """获取指定日期的产能（工时）"""
        if not self.shift_template:
            return 0.0
        return self.shift_template.hours_for_weekday(d.weekday())


def compute_eta(order: Order, cal: WorkCalendar) -> Dict[str, object]:
    """
    计算订单的预计完成时间（ETA）
    
    Args:
        order: 订单对象
        cal: 工作日历
        
    Returns:
        包含以下键的字典:
        - eta_dt: 预计完成时间
        - remaining_hours: 剩余工时
        - daily_capacity_map: 每日产能映射
        - explanation: 解释说明列表
    """
    equipment_map = equipment_available_map(order)
    remaining_hours = sum(product_remaining_hours(p, equipment_map) for p in order.products)

    # 构建损失工时映射
    lost_map: Dict[date, float] = {}
    reason_map: Dict[date, List[str]] = {}
    for ev in order.events:
        lost_map[ev.day] = lost_map.get(ev.day, 0.0) + ev.hours_lost
        reason_map.setdefault(ev.day, []).append(f"{ev.reason}(-{ev.hours_lost:g}h)")

    # 构建加班工时映射
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
            hours = product_remaining_hours(p, equipment_map)
            explanation.append(
                f"- {p.product_id} (PN={p.part_number or '-'} qty={p.quantity}): {hours:g}h"
            )
        explanation.append("")

    # 如果没有剩余工时，直接返回
    if remaining_hours <= 0:
        return {
            "eta_dt": order.start_dt,
            "remaining_hours": 0.0,
            "daily_capacity_map": {},
            "explanation": explanation + ["All phases completed. ETA equals start time."]
        }

    current_day = order.start_dt.date()
    hours_left = remaining_hours
    daily_capacity_map: Dict[date, float] = {}

    # 逐日计算，直到工时消耗完毕
    for _ in range(3650):  # 最多10年
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
                        "explanation": explanation or ["No blocking events."]
                    }
                hours_left -= cap
        current_day = current_day + timedelta(days=1)

    raise RuntimeError("ETA computation exceeded safe bounds.")
