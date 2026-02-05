# -*- coding: utf-8 -*-
"""
生产交期优化系统 - 工具函数
Manufacturing Optimization System - Utility Functions
"""

from datetime import date
from typing import List, Dict, Optional
from PySide6 import QtCore

from .models import Phase, Product, Equipment, Order


def qdate_to_date(qdate: QtCore.QDate) -> date:
    """将 QDate 转换为 Python date"""
    return date(qdate.year(), qdate.month(), qdate.day())


def date_to_qdate(d: date) -> QtCore.QDate:
    """将 Python date 转换为 QDate"""
    return QtCore.QDate(d.year, d.month, d.day)


def chinese_locale() -> QtCore.QLocale:
    """获取中文语言环境"""
    try:
        return QtCore.QLocale(QtCore.QLocale.Language.Chinese, QtCore.QLocale.Country.China)
    except AttributeError:
        return QtCore.QLocale(QtCore.QLocale.Chinese, QtCore.QLocale.China)


def equipment_available_map_from_list(equipment: List[Equipment]) -> Dict[str, int]:
    """从设备列表创建设备可用数量映射"""
    result: Dict[str, int] = {}
    for eq in equipment:
        if not eq.equipment_id:
            continue
        result[eq.equipment_id] = max(0, int(eq.available_count))
    return result


def equipment_available_map(order: Order) -> Dict[str, int]:
    """从订单获取设备可用数量映射"""
    return equipment_available_map_from_list(order.equipment)


def normalize_equipment_ids(items: List[str]) -> List[str]:
    """规范化设备ID列表，去重并过滤无效值"""
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


def split_equipment_ids(text: str) -> List[str]:
    """将逗号分隔的设备ID字符串分割为列表"""
    if not text:
        return []
    if text.strip() in ("-", "无需设备"):
        return []
    return normalize_equipment_ids(text.split(","))


def format_equipment_ids(ids: List[str]) -> str:
    """将设备ID列表格式化为逗号分隔的字符串"""
    return ",".join(normalize_equipment_ids(ids))


def phase_effective_hours(phase: Phase, quantity: int, equipment_map: Dict[str, int]) -> float:
    """计算工序的有效工时（考虑设备数量）"""
    base = max(0.0, float(phase.planned_hours))
    hours = base
    equipment_ids = split_equipment_ids(phase.equipment_id)
    if equipment_ids:
        available = sum(max(0, equipment_map.get(eq_id, 0)) for eq_id in equipment_ids)
        if available > 0:
            hours = hours / available
    return hours


def phase_total_hours(phase: Phase, quantity: int) -> float:
    """获取工序的总计划工时"""
    return max(0.0, float(phase.planned_hours))


def phase_completion_ratio(phase: Phase, quantity: int) -> float:
    """计算工序的完成比例"""
    total = phase_total_hours(phase, quantity)
    if total <= 0:
        return 0.0
    completed = max(0.0, float(phase.completed_hours))
    completed = min(completed, total)
    return completed / total


def product_remaining_hours(product: Product, equipment_map: Dict[str, int]) -> float:
    """计算产品的剩余工时"""
    total = 0.0
    parallel_groups: Dict[int, List[float]] = {}
    for phase in product.phases:
        hours = phase_effective_hours(phase, product.quantity, equipment_map)
        ratio = phase_completion_ratio(phase, product.quantity)
        remaining = hours * (1.0 - ratio)
        if phase.parallel_group > 0:
            parallel_groups.setdefault(phase.parallel_group, []).append(remaining)
        else:
            total += remaining
    for group_hours in parallel_groups.values():
        total += max(group_hours)
    return total


def product_progress(product: Product, equipment_map: Dict[str, int]) -> float:
    """计算产品的工时进度"""
    total = 0.0
    done = 0.0
    for phase in product.phases:
        hours = phase_effective_hours(phase, product.quantity, equipment_map)
        total += hours
        ratio = phase_completion_ratio(phase, product.quantity)
        done += hours * ratio
    if total <= 0:
        return 0.0
    return min(done / total, 1.0)


def product_quantity_progress(product: Product) -> float:
    """计算产品的数量进度"""
    qty = max(int(product.quantity), 0)
    if qty <= 0:
        return 0.0
    produced = max(0, int(product.produced_qty))
    if produced > qty:
        produced = qty
    return produced / qty
