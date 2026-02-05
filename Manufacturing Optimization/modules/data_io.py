from datetime import datetime, date
from typing import Dict, List

from .models import (
    Order,
    Product,
    Phase,
    Equipment,
    Event,
    CapacityAdjustment,
    DefectRecord,
    LogEntry,
)


def order_to_dict(order: Order, include_equipment: bool = True) -> Dict[str, object]:
    data: Dict[str, object] = {
        "version": 12,
        "order_id": order.order_id,
        "start_dt": order.start_dt.isoformat(),
        "order_date": order.order_date.isoformat(),
        "customer_code": order.customer_code,
        "shipping_method": order.shipping_method,
        "due_date": order.due_date.isoformat(),
        "employees": order.employees,
        "products": [
            {
                "product_id": p.product_id,
                "part_number": p.part_number,
                "quantity": p.quantity,
                "produced_qty": p.produced_qty,
                "unit_weight_g": p.unit_weight_g,
                "phases": [
                    {
                        "name": ph.name,
                        "planned_hours": ph.planned_hours,
                        "completed_hours": ph.completed_hours,
                        "parallel_group": ph.parallel_group,
                        "equipment_id": ph.equipment_id,
                        "assigned_employee": ph.assigned_employee,
                    }
                    for ph in p.phases
                ],
            }
            for p in order.products
        ],
        "events": [
            {
                "day": e.day.isoformat(),
                "hours_lost": e.hours_lost,
                "reason": e.reason,
                "remark": e.remark,
            }
            for e in order.events
        ],
        "capacity_adjustments": [
            {
                "day": a.day.isoformat(),
                "extra_hours": a.extra_hours,
                "reason": a.reason,
                "equipment_ids": a.equipment_ids,
            }
            for a in order.adjustments
        ],
        "defects": [
            {
                "product_id": d.product_id,
                "count": d.count,
                "category": d.category,
                "detail": d.detail,
                "timestamp": d.timestamp.isoformat(),
            }
            for d in order.defects
        ],
    }
    if include_equipment:
        data["equipment"] = [
            {
                "equipment_id": e.equipment_id,
                "category": e.category,
                "total_count": e.total_count,
                "available_count": e.available_count,
                "shift_template_name": e.shift_template_name,
            }
            for e in order.equipment
        ]
    return data


def order_from_dict(data: Dict[str, object]) -> Order:
    version = int(data.get("version", 0) or 0)
    per_unit_hours = version < 11
    order_id = data.get("order_id", "O-UNKNOWN")
    start_dt = datetime.fromisoformat(data.get("start_dt", datetime.now().isoformat()))
    order_date_raw = data.get("order_date", "")
    if order_date_raw:
        try:
            order_date_val = date.fromisoformat(order_date_raw)
        except Exception:
            order_date_val = start_dt.date()
    else:
        order_date_val = start_dt.date()
    due_date_raw = data.get("due_date", "")
    if due_date_raw:
        try:
            due_date_val = date.fromisoformat(due_date_raw)
        except Exception:
            due_date_val = order_date_val
    else:
        due_date_val = order_date_val
    customer_code = data.get("customer_code", "")
    shipping_method = data.get("shipping_method", "")

    equipment = [
        Equipment(
            equipment_id=e.get("equipment_id", ""),
            category=e.get("category", ""),
            total_count=int(e.get("total_count", 1)),
            available_count=int(e.get("available_count", 1)),
            shift_template_name=e.get("shift_template_name", ""),
        )
        for e in data.get("equipment", [])
    ]

    employees = list(data.get("employees", []))

    products: List[Product] = []
    if "products" in data:
        for p in data.get("products", []):
            quantity = int(p.get("quantity", 1))
            produced_qty = int(p.get("produced_qty", 0))
            if quantity < 0:
                quantity = 0
            if produced_qty < 0:
                produced_qty = 0
            if produced_qty > quantity:
                produced_qty = quantity
            unit_weight_g = float(p.get("unit_weight_g", 0.0))
            phases = []
            for ph in p.get("phases", []):
                planned_hours = float(ph.get("planned_hours", 0))
                if per_unit_hours:
                    planned_hours = planned_hours * max(quantity, 1)
                completed = ph.get("completed_hours", None)
                if completed is None:
                    completed = planned_hours if ph.get("done", False) else 0.0
                phases.append(
                    Phase(
                        name=ph.get("name", ""),
                        planned_hours=planned_hours,
                        completed_hours=float(completed),
                        parallel_group=int(ph.get("parallel_group", 0)),
                        equipment_id=ph.get("equipment_id", ""),
                        assigned_employee=ph.get("assigned_employee", ""),
                    )
                )
            products.append(
                Product(
                    product_id=p.get("product_id", "Product"),
                    part_number=p.get("part_number", ""),
                    quantity=quantity,
                    produced_qty=produced_qty,
                    unit_weight_g=unit_weight_g,
                    phases=phases,
                )
            )
    elif "phases" in data:
        # Backward compatibility: old single-product format
        quantity = int(data.get("quantity", 1))
        if quantity < 0:
            quantity = 0
        phases = []
        for ph in data.get("phases", []):
            planned_hours = float(ph.get("planned_hours", 0))
            if per_unit_hours:
                planned_hours = planned_hours * max(quantity, 1)
            completed = ph.get("completed_hours", None)
            if completed is None:
                completed = planned_hours if ph.get("done", False) else 0.0
            phases.append(
                Phase(
                    name=ph.get("name", ""),
                    planned_hours=planned_hours,
                    completed_hours=float(completed),
                    parallel_group=int(ph.get("parallel_group", 0)),
                    equipment_id=ph.get("equipment_id", ""),
                )
            )
        products.append(
            Product(
                product_id="产品1",
                part_number=data.get("part_number", ""),
                quantity=quantity,
                produced_qty=0,
                unit_weight_g=float(data.get("unit_weight_g", 0.0)),
                phases=phases,
            )
        )

    events = [
        Event(
            day=date.fromisoformat(e.get("day")),
            hours_lost=float(e.get("hours_lost", 0)),
            reason=e.get("reason", ""),
            remark=e.get("remark", ""),
        )
        for e in data.get("events", [])
        if e.get("day")
    ]

    adjustments = [
        CapacityAdjustment(
            day=date.fromisoformat(a.get("day")),
            extra_hours=float(a.get("extra_hours", 0)),
            reason=a.get("reason", ""),
            equipment_ids=list(a.get("equipment_ids", [])),
        )
        for a in data.get("capacity_adjustments", [])
        if a.get("day")
    ]

    defects = [
        DefectRecord(
            product_id=d.get("product_id", ""),
            count=int(d.get("count", 0)),
            category=d.get("category", ""),
            detail=d.get("detail", ""),
            timestamp=datetime.fromisoformat(d.get("timestamp"))
            if d.get("timestamp")
            else datetime.now(),
        )
        for d in data.get("defects", [])
    ]

    logs = [
        LogEntry(
            timestamp=datetime.fromisoformat(l.get("timestamp"))
            if l.get("timestamp")
            else datetime.now(),
            user=l.get("user", ""),
            content=l.get("content", ""),
            order_id=l.get("order_id", ""),
        )
        for l in data.get("logs", [])
    ]

    return Order(
        order_id=order_id,
        start_dt=start_dt,
        order_date=order_date_val,
        customer_code=customer_code,
        shipping_method=shipping_method,
        due_date=due_date_val,
        products=products,
        events=events,
        adjustments=adjustments,
        defects=defects,
        equipment=equipment,
        employees=employees,
        logs=logs,
    )
