# -*- coding: utf-8 -*-
"""
生产交期优化系统 - 模块化设计
Manufacturing Optimization System - Modular Design
"""

from .models import (
    Phase, Product, Equipment, ShiftDayPlan, ShiftTemplate,
    Event, CapacityAdjustment, DefectRecord, LogEntry, MemoEntry,
    Order, UserAccount
)
from .scheduling import WorkCalendar, compute_eta
from .utils import (
    qdate_to_date, date_to_qdate, chinese_locale,
    equipment_available_map_from_list, equipment_available_map,
    normalize_equipment_ids, split_equipment_ids, format_equipment_ids,
    phase_effective_hours, phase_total_hours, phase_completion_ratio,
    product_remaining_hours, product_progress, product_quantity_progress
)
from .delegates import ComboBoxDelegate, SpinBoxDelegate
from .data_io import order_to_dict, order_from_dict, factory_to_dict

__all__ = [
    # Models
    'Phase', 'Product', 'Equipment', 'ShiftDayPlan', 'ShiftTemplate',
    'Event', 'CapacityAdjustment', 'DefectRecord', 'LogEntry', 'MemoEntry',
    'Order', 'UserAccount',
    # Scheduling
    'WorkCalendar', 'compute_eta',
    # Utils
    'qdate_to_date', 'date_to_qdate', 'chinese_locale',
    'equipment_available_map_from_list', 'equipment_available_map',
    'normalize_equipment_ids', 'split_equipment_ids', 'format_equipment_ids',
    'phase_effective_hours', 'phase_total_hours', 'phase_completion_ratio',
    'product_remaining_hours', 'product_progress', 'product_quantity_progress',
    # Delegates
    'ComboBoxDelegate', 'SpinBoxDelegate',
    # Data IO
    'order_to_dict', 'order_from_dict', 'factory_to_dict',
]
