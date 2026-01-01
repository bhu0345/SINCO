import tkinter as tk
from tkinter import ttk, messagebox
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta, date
from typing import List, Dict, Optional
import json
import os

# ----------------------------
# Data models
# ----------------------------

@dataclass
class Phase:
    name: str
    planned_hours: float
    done: bool = False
    parallel_group: int = 0  # 0=顺序执行, >0=并行组编号

@dataclass
class Event:
    day: date
    hours_lost: float
    reason: str

@dataclass
class Order:
    order_id: str
    start_dt: datetime
    phases: List[Phase] = field(default_factory=list)
    events: List[Event] = field(default_factory=list)
    lathe_ops: int = 2
    blank_lead_days: int = 3
    quantity: int = 1  # 订单数量

# ----------------------------
# Scheduling / ETA computation
# ----------------------------

class WorkCalendar:
    def __init__(self, working_hours_per_day: float = 8.0):
        self.working_hours_per_day = working_hours_per_day

    @staticmethod
    def is_workday(d: date) -> bool:
        return d.weekday() < 5  # Mon-Fri

def compute_eta(order: Order, cal: WorkCalendar) -> Dict[str, object]:
    # 计算实际总工时（考虑并行工序）
    def calculate_total_hours(phases: List[Phase]) -> float:
        """计算考虑并行的总工时"""
        total = 0.0
        parallel_groups = {}
        
        for p in phases:
            if p.done:
                continue
            
            if p.parallel_group == 0:
                # 顺序执行的工序，直接累加
                total += p.planned_hours
            else:
                # 并行工序，记录到对应的组
                if p.parallel_group not in parallel_groups:
                    parallel_groups[p.parallel_group] = []
                parallel_groups[p.parallel_group].append(p.planned_hours)
        
        # 对于每个并行组，只计入最长的工时
        for group_hours in parallel_groups.values():
            total += max(group_hours)
        
        return total
    
    remaining_hours = calculate_total_hours(order.phases)

    lost_map: Dict[date, float] = {}
    reason_map: Dict[date, List[str]] = {}
    for ev in order.events:
        lost_map[ev.day] = lost_map.get(ev.day, 0.0) + ev.hours_lost
        reason_map.setdefault(ev.day, []).append(f"{ev.reason}(-{ev.hours_lost:g}h)")

    explanation = []
    
    # 统计并行组信息
    parallel_info = {}
    parallel_hours = {}  # 记录实际工时
    for p in order.phases:
        if not p.done and p.parallel_group > 0:
            if p.parallel_group not in parallel_info:
                parallel_info[p.parallel_group] = []
                parallel_hours[p.parallel_group] = []
            parallel_info[p.parallel_group].append(f"{p.name}({p.planned_hours:g}h)")
            parallel_hours[p.parallel_group].append(p.planned_hours)
    
    if parallel_info:
        explanation.append("=== 并行工序组 ===")
        for group in sorted(parallel_info.keys()):
            phases = parallel_info[group]
            max_hours = max(parallel_hours[group])
            explanation.append(f"并行组{group}: {', '.join(phases)} -> 取最长{max_hours:g}h")
        explanation.append("")
    
    if remaining_hours <= 0:
        return {
            "eta_dt": order.start_dt,
            "remaining_hours": 0.0,
            "daily_capacity_map": {},
            "explanation": ["所有工序已完成。预计交期 = 开始时间。"]
        }

    current_day = order.start_dt.date()
    hours_left = remaining_hours
    daily_capacity_map: Dict[date, float] = {}

    for _ in range(3650):
        if cal.is_workday(current_day):
            lost = lost_map.get(current_day, 0.0)
            cap = max(cal.working_hours_per_day - lost, 0.0)
            daily_capacity_map[current_day] = cap

            if lost > 0:
                explanation.append(
                    f"{current_day.isoformat()}: capacity {cal.working_hours_per_day:g}h - lost {lost:g}h => {cap:g}h "
                    f"({', '.join(reason_map.get(current_day, []))})"
                )

            if cap > 0:
                if hours_left <= cap:
                    finish_time = datetime.combine(current_day, datetime.min.time()).replace(hour=9, minute=0)
                    finish_time += timedelta(hours=hours_left)
                    return {
                        "eta_dt": finish_time,
                        "remaining_hours": remaining_hours,
                        "daily_capacity_map": daily_capacity_map,
                        "explanation": explanation or ["没有影响进度的事件。"]
                    }
                else:
                    hours_left -= cap
        current_day = current_day + timedelta(days=1)

    raise RuntimeError("交期计算超出安全限制。")

# ----------------------------
# Phase generation helpers
# ----------------------------

def build_lathe_chain(n_ops: int, hours_lathe: float = 12.0, hours_insp: float = 4.0) -> List[Phase]:
    phases: List[Phase] = []
    for i in range(1, n_ops + 1):
        phases.append(Phase(f"车床工序{i}", hours_lathe))
        phases.append(Phase(f"检验{i}", hours_insp))
    return phases

def template_with_mold(lathe_ops: int) -> List[Phase]:
    phases = [
        Phase("模具开发(外协)", 80),
        Phase("工装夹具制作", 24),
        Phase("制定加工工艺", 16),
        Phase("量具/刀具准备", 8),
        Phase("采购(物料/毛坯)", 16),
    ]
    phases += build_lathe_chain(lathe_ops)
    phases += [
        Phase("表面处理(阳极/试漏等)", 16),
        Phase("检验入库", 8),
        Phase("包装", 8),
        Phase("等待发货", 0),
    ]
    return phases

def template_no_mold(lathe_ops: int) -> List[Phase]:
    phases = [
        Phase("制定工艺路线", 12),
        Phase("采购刀具量具", 8),
        Phase("工装夹具制作", 24),
        Phase("采购(物料/毛坯)", 16),
    ]
    phases += build_lathe_chain(lathe_ops)
    phases += [
        Phase("表面处理(阳极/试漏等)", 16),
        Phase("检验入库", 8),
        Phase("包装", 8),
        Phase("等待发货", 0),
    ]
    return phases

# ----------------------------
# GUI
# ----------------------------

class ETAGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("生产交期预测系统 v2")
        self.geometry("1200x750")
        
        print("正在初始化GUI...")  # 调试输出

        self.cal = WorkCalendar(working_hours_per_day=8.0)
        self.order: Optional[Order] = None
        self.route_mode = "with_mold"
        self.save_file = "order_data.json"

        print("开始构建UI...")  # 调试输出
        self._build_ui()
        print("UI构建完成")  # 调试输出
        
        self._load_order()  # 启动时自动加载
        
        # 如果没有加载到订单，显示欢迎信息
        if not self.order:
            self._explain("欢迎使用生产交期预测系统！")
            self._explain("请点击'创建/重置订单'按钮开始。")
        
        print("GUI初始化完成，窗口应该已显示")  # 调试输出
        
        # 强制更新窗口显示
        self.update_idletasks()
        self.update()
        
        # 确保窗口显示在最前面（macOS可能需要）
        self.lift()
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
        
        print(f"窗口大小: {self.winfo_width()}x{self.winfo_height()}")
        print(f"窗口位置: ({self.winfo_x()}, {self.winfo_y()})")
        print(f"窗口是否可见: {self.winfo_viewable()}")

    def _build_ui(self):
        # 添加背景色，使窗口内容更明显
        self.configure(bg='#f0f0f0')
        
        # 添加醒目的标题
        # title_frame = ttk.Frame(self)
        # title_frame.pack(fill="x", padx=10, pady=(10, 0))
        # title_label = tk.Label(title_frame, text="生产交期预测系统 v2", 
        #                       font=("Helvetica", 16, "bold"), 
        #                       bg='#2196F3', fg='white', pady=10)
        # title_label.pack(fill="x")
        
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        # 第一行
        ttk.Label(top, text="订单编号").grid(row=0, column=0, sticky="w")
        self.order_id_var = tk.StringVar(value="O-001")
        ttk.Entry(top, textvariable=self.order_id_var, width=16).grid(row=0, column=1, padx=6)

        ttk.Label(top, text="订单数量（件）").grid(row=0, column=2, sticky="w", padx=(10,0))
        self.quantity_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self.quantity_var, width=10).grid(row=0, column=3, padx=6)

        self.route_var = tk.StringVar(value="with_mold")
        ttk.Radiobutton(top, text="需要模具开发", variable=self.route_var, value="with_mold").grid(row=0, column=4, padx=10)
        ttk.Radiobutton(top, text="不需要模具开发", variable=self.route_var, value="no_mold").grid(row=0, column=5, padx=10)

        # 第二行
        ttk.Label(top, text="车床工序数N").grid(row=1, column=0, sticky="w")
        self.lathe_ops_var = tk.StringVar(value="2")
        ttk.Entry(top, textvariable=self.lathe_ops_var, width=16).grid(row=1, column=1, padx=6)

        ttk.Label(top, text="重采毛坯周期(天)").grid(row=1, column=2, sticky="w", padx=(10,0))
        self.blank_days_var = tk.StringVar(value="3")
        ttk.Entry(top, textvariable=self.blank_days_var, width=10).grid(row=1, column=3, padx=6)

        # 按钮组
        btn_frame = ttk.Frame(top)
        btn_frame.grid(row=0, column=6, rowspan=2, padx=10)
        ttk.Button(btn_frame, text="创建/重置订单", command=self.create_order).pack(pady=2)
        ttk.Button(btn_frame, text="保存订单", command=self.save_order).pack(pady=2)
        ttk.Button(btn_frame, text="加载订单", command=self.load_order_button).pack(pady=2)

        ttk.Separator(self).pack(fill="x", padx=10, pady=6)

        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(main)
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(main)
        right.pack(side="right", fill="y")

        ttk.Label(left, text="工序阶段（可按住Ctrl多选工序，然后设置并行）").pack(anchor="w")
        self.phase_tree = ttk.Treeview(left, columns=("name", "hours", "parallel", "done"), show="headings", height=18, selectmode="extended")
        self.phase_tree.heading("name", text="工序名称")
        self.phase_tree.heading("hours", text="计划工时")
        self.phase_tree.heading("parallel", text="执行方式")
        self.phase_tree.heading("done", text="已完成？")
        self.phase_tree.column("name", width=280, anchor="w")
        self.phase_tree.column("hours", width=90, anchor="center")
        self.phase_tree.column("parallel", width=90, anchor="center")
        self.phase_tree.column("done", width=80, anchor="center")
        self.phase_tree.pack(fill="both", expand=True, pady=6)
        self.phase_tree.bind("<<TreeviewSelect>>", self._on_phase_select)

        edit = ttk.Frame(left)
        edit.pack(fill="x", pady=6)

        # 工序编辑部分
        edit_row1 = ttk.Frame(edit)
        edit_row1.pack(fill="x", pady=2)
        ttk.Label(edit_row1, text="工序名称:").pack(side="left")
        self.phase_name_var = tk.StringVar(value="")
        ttk.Entry(edit_row1, textvariable=self.phase_name_var, width=18).pack(side="left", padx=4)
        ttk.Label(edit_row1, text="工时:").pack(side="left")
        self.phase_hours_var = tk.StringVar(value="")
        ttk.Entry(edit_row1, textvariable=self.phase_hours_var, width=8).pack(side="left", padx=4)
        ttk.Button(edit_row1, text="添加新工序", command=self.add_phase).pack(side="left", padx=6)
        
        edit_row2 = ttk.Frame(edit)
        edit_row2.pack(fill="x", pady=2)
        ttk.Button(edit_row2, text="更新工时", command=self.update_phase_hours).pack(side="left", padx=6)
        ttk.Button(edit_row2, text="更新名称", command=self.update_phase_name).pack(side="left", padx=6)
        ttk.Button(edit_row2, text="切换完成状态", command=self.toggle_phase_done).pack(side="left", padx=6)
        ttk.Button(edit_row2, text="删除工序", command=self.delete_phase).pack(side="left", padx=6)
        
        edit_row3 = ttk.Frame(edit)
        edit_row3.pack(fill="x", pady=2)
        ttk.Label(edit_row3, text="⏸️ 并行设置:").pack(side="left")
        ttk.Button(edit_row3, text="设为并行工序", command=self.set_parallel_group).pack(side="left", padx=6)
        ttk.Button(edit_row3, text="取消并行(改为顺序)", command=self.clear_parallel).pack(side="left", padx=6)
        ttk.Button(edit_row3, text="报废重做", command=self.report_scrap).pack(side="left", padx=12)
        ttk.Button(edit_row3, text="重新计算交期", command=self.refresh_eta).pack(side="right")

        # 并行组说明
        hint = ttk.Label(edit, text="💡 并行操作: 先选中多个工序(按住Ctrl多选), 然后点击'设为并行工序'即可让它们同时进行", 
                        foreground="blue", font=("", 9))
        hint.pack(anchor="w", pady=2)

        # Events (lost hours)
        ttk.Label(right, text="事件（损失工时）").pack(anchor="w")
        evf = ttk.Frame(right)
        evf.pack(fill="x", pady=6)

        ttk.Label(evf, text="日期 (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        self.ev_date_var = tk.StringVar(value=datetime.now().date().isoformat())
        ttk.Entry(evf, textvariable=self.ev_date_var, width=16).grid(row=0, column=1, padx=6)

        ttk.Label(evf, text="损失工时").grid(row=1, column=0, sticky="w")
        self.ev_hours_var = tk.StringVar(value="8")
        ttk.Entry(evf, textvariable=self.ev_hours_var, width=16).grid(row=1, column=1, padx=6)

        ttk.Label(evf, text="原因").grid(row=2, column=0, sticky="w")
        self.ev_reason_var = tk.StringVar(value="员工请假")
        ttk.Entry(evf, textvariable=self.ev_reason_var, width=16).grid(row=2, column=1, padx=6)

        ttk.Button(evf, text="添加事件", command=self.add_event).grid(row=3, column=0, columnspan=2, pady=6, sticky="we")

        self.event_list = tk.Listbox(right, height=10, width=34)
        self.event_list.pack(fill="x", pady=6)
        ttk.Button(right, text="删除选中的事件", command=self.remove_event).pack(fill="x")

        ttk.Separator(right).pack(fill="x", pady=10)

        ttk.Label(right, text="预计交期").pack(anchor="w")
        self.eta_var = tk.StringVar(value="(请先创建订单)")
        ttk.Label(right, textvariable=self.eta_var, font=("Helvetica", 12, "bold")).pack(anchor="w", pady=6)

        self.remaining_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.remaining_var).pack(anchor="w")

        ttk.Label(right, text="说明").pack(anchor="w", pady=(10, 0))
        self.explain_text = tk.Text(right, height=12, width=38)
        self.explain_text.pack(fill="both", expand=True)

    def create_order(self):
        oid = self.order_id_var.get().strip() or "O-UNKNOWN"
        start = datetime.now()

        try:
            n_ops = int(self.lathe_ops_var.get().strip())
            if n_ops < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的数值", "车床工序数 N 必须是大于0的整数")
            return

        try:
            blank_days = int(self.blank_days_var.get().strip())
            if blank_days < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的数值", "重采毛坯周期(天) 必须是非负整数")
            return

        try:
            quantity = int(self.quantity_var.get().strip())
            if quantity < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的数值", "订单数量必须是大于0的整数")
            return

        self.route_mode = self.route_var.get()
        phases = template_with_mold(n_ops) if self.route_mode == "with_mold" else template_no_mold(n_ops)

        self.order = Order(
            order_id=oid,
            start_dt=start,
            phases=phases,
            events=[],
            lathe_ops=n_ops,
            blank_lead_days=blank_days,
            quantity=quantity
        )

        self._reload_phase_tree()
        self._reload_event_list()
        self.refresh_eta()
        self._explain(f"已创建订单 {oid}。数量={quantity}, 工艺路线={self.route_mode}, 车床工序数={n_ops}, 重采毛坯周期={blank_days}天。")

    def _reload_phase_tree(self):
        for row in self.phase_tree.get_children():
            self.phase_tree.delete(row)
        if not self.order:
            return
        for idx, p in enumerate(self.order.phases):
            # 并行显示：使用符号 ⏸️ 表示并行组
            if p.parallel_group > 0:
                parallel_display = f"⏸️组{p.parallel_group}"
            else:
                parallel_display = "→顺序"
            self.phase_tree.insert(
                "", "end", iid=str(idx),
                values=(p.name, f"{p.planned_hours:g}", parallel_display, "是" if p.done else "否")
            )

    def _reload_event_list(self):
        self.event_list.delete(0, tk.END)
        if not self.order:
            return
        for i, ev in enumerate(self.order.events):
            self.event_list.insert(tk.END, f"{i}. {ev.day.isoformat()}  -{ev.hours_lost:g}h  {ev.reason}")

    def _get_selected_phase_index(self) -> Optional[int]:
        sel = self.phase_tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _on_phase_select(self, event):
        """当选中工序时，自动填充到编辑框（仅单选时）"""
        sel = self.phase_tree.selection()
        if len(sel) == 1 and self.order:
            idx = int(sel[0])
            if idx < len(self.order.phases):
                phase = self.order.phases[idx]
                self.phase_name_var.set(phase.name)
                self.phase_hours_var.set(str(phase.planned_hours))

    def update_phase_hours(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        idx = self._get_selected_phase_index()
        if idx is None:
            messagebox.showwarning("未选择", "请先选择一个工序。")
            return
        try:
            h = float(self.phase_hours_var.get())
            if h < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的工时", "计划工时必须是非负数。")
            return
        self.order.phases[idx].planned_hours = h
        self._reload_phase_tree()
        self.refresh_eta()

    def update_phase_name(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        idx = self._get_selected_phase_index()
        if idx is None:
            messagebox.showwarning("未选择", "请先选择一个工序。")
            return
        name = self.phase_name_var.get().strip()
        if not name:
            messagebox.showerror("无效的名称", "工序名称不能为空。")
            return
        self.order.phases[idx].name = name
        self._reload_phase_tree()
        self.refresh_eta()

    def set_parallel_group(self):
        """将选中的多个工序设置为同一并行组"""
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        
        sel = self.phase_tree.selection()
        if len(sel) < 2:
            messagebox.showinfo("提示", "请先按住Ctrl键选中至少2个工序，然后点击此按钮将它们设为并行。")
            return
        
        # 找到当前最大的并行组编号
        max_group = max((p.parallel_group for p in self.order.phases), default=0)
        new_group = max_group + 1
        
        # 将选中的工序设为新的并行组
        phase_names = []
        for iid in sel:
            idx = int(iid)
            if idx < len(self.order.phases):
                self.order.phases[idx].parallel_group = new_group
                phase_names.append(self.order.phases[idx].name)
        
        self._reload_phase_tree()
        self.refresh_eta()
        self._explain(f"已将 {len(sel)} 个工序设为并行组{new_group}: {', '.join(phase_names)}")
        messagebox.showinfo("成功", f"已将以下工序设为并行组{new_group}（可同时进行）:\n\n" + "\n".join(phase_names))
    
    def clear_parallel(self):
        """将选中的工序改为顺序执行"""
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        
        sel = self.phase_tree.selection()
        if not sel:
            messagebox.showwarning("未选择", "请先选择要改为顺序执行的工序。")
            return
        
        phase_names = []
        for iid in sel:
            idx = int(iid)
            if idx < len(self.order.phases):
                self.order.phases[idx].parallel_group = 0
                phase_names.append(self.order.phases[idx].name)
        
        self._reload_phase_tree()
        self.refresh_eta()
        self._explain(f"已将 {len(sel)} 个工序改为顺序执行: {', '.join(phase_names)}")

    def add_phase(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        name = self.phase_name_var.get().strip()
        if not name:
            messagebox.showerror("无效的名称", "工序名称不能为空。")
            return
        try:
            hours = float(self.phase_hours_var.get())
            if hours < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的工时", "计划工时必须是非负数。")
            return
        
        # 在选中的工序后面插入，如果没有选中则添加到最后
        sel = self.phase_tree.selection()
        idx = int(sel[0]) if sel and len(sel) == 1 else None
        insert_pos = (idx + 1) if idx is not None else len(self.order.phases)
        self.order.phases.insert(insert_pos, Phase(name=name, planned_hours=hours, parallel_group=0))
        self._reload_phase_tree()
        self.refresh_eta()
        self._explain(f"已添加新工序: {name} ({hours}小时)")

    def delete_phase(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        idx = self._get_selected_phase_index()
        if idx is None:
            messagebox.showwarning("未选择", "请先选择一个工序。")
            return
        phase_name = self.order.phases[idx].name
        if messagebox.askyesno("确认删除", f"确定要删除工序 '{phase_name}' 吗？"):
            self.order.phases.pop(idx)
            self._reload_phase_tree()
            self.refresh_eta()
            self._explain(f"已删除工序: {phase_name}")

    def toggle_phase_done(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        idx = self._get_selected_phase_index()
        if idx is None:
            messagebox.showwarning("未选择", "请先选择一个工序。")
            return
        self.order.phases[idx].done = not self.order.phases[idx].done
        self._reload_phase_tree()
        self.refresh_eta()

    def add_event(self):
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        try:
            d = datetime.strptime(self.ev_date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("无效的日期", "日期格式必须是 YYYY-MM-DD。")
            return
        try:
            hours = float(self.ev_hours_var.get())
            if hours < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("无效的工时", "损失工时必须是非负数。")
            return
        reason = self.ev_reason_var.get().strip() or "事件"
        self.order.events.append(Event(day=d, hours_lost=hours, reason=reason))
        self._reload_event_list()
        self.refresh_eta()

    def remove_event(self):
        if not self.order:
            return
        sel = self.event_list.curselection()
        if not sel:
            return
        i = sel[0]
        if 0 <= i < len(self.order.events):
            self.order.events.pop(i)
        self._reload_event_list()
        self.refresh_eta()

    def report_scrap(self):
        """
        报废重做功能：
        - 支持部分报废（报废比例 0~1）
        - 根据报废比例计算需要补做的数量
        - 按比例缩放重做工序的工时
        """
        if not self.order:
            messagebox.showwarning("无订单", "请先创建订单。")
            return
        idx = self._get_selected_phase_index()
        if idx is None:
            messagebox.showwarning("未选择", "请先选择一个检验工序（检验X）。")
            return

        phase = self.order.phases[idx]
        if not phase.name.startswith("检验"):
            messagebox.showerror("非检验工序", "请选择一个'检验X'阶段，然后再点报废重做。")
            return

        # Guard: if any later phases already marked done, this MVP can't safely remodel that history
        if any(p.done for p in self.order.phases[idx+1:]):
            messagebox.showerror(
                "暂不允许报废",
                "你选择的检验后面已经有阶段标记为完成。\n"
                "这个 MVP 版本为了避免逻辑混乱，暂不支持在后续已完成时再触发整批重做。\n"
                "建议：把后续完成状态先取消，再触发报废重做。"
            )
            return

        # 弹出对话框输入报废比例
        scrap_dialog = tk.Toplevel(self)
        scrap_dialog.title("报废重做")
        scrap_dialog.geometry("400x200")
        scrap_dialog.transient(self)
        scrap_dialog.grab_set()

        ttk.Label(scrap_dialog, text=f"当前检验工序: {phase.name}", font=("", 10, "bold")).pack(pady=10)
        ttk.Label(scrap_dialog, text=f"订单总数量: {self.order.quantity} 件").pack(pady=5)

        frame = ttk.Frame(scrap_dialog)
        frame.pack(pady=10)
        
        ttk.Label(frame, text="报废比例 (0.0-1.0):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        scrap_ratio_var = tk.StringVar(value="1.0")
        ttk.Entry(frame, textvariable=scrap_ratio_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        result_label = ttk.Label(frame, text="", foreground="blue")
        result_label.grid(row=1, column=0, columnspan=2, pady=5)

        def update_preview(*args):
            try:
                ratio = float(scrap_ratio_var.get())
                if 0 <= ratio <= 1:
                    scrap_qty = int(self.order.quantity * ratio)
                    result_label.config(text=f"报废数量: {scrap_qty} 件\n需要补做: {scrap_qty} 件")
                else:
                    result_label.config(text="比例必须在 0.0 到 1.0 之间", foreground="red")
            except:
                result_label.config(text="请输入有效数字", foreground="red")

        scrap_ratio_var.trace('w', update_preview)
        update_preview()

        def confirm_scrap():
            try:
                ratio = float(scrap_ratio_var.get())
                if ratio < 0 or ratio > 1:
                    raise ValueError("比例必须在 0.0 到 1.0 之间")
                if ratio == 0:
                    messagebox.showinfo("提示", "报废比例为0，无需重做。")
                    scrap_dialog.destroy()
                    return
                
                scrap_qty = int(self.order.quantity * ratio)
                
                # Insert rebuild chain with scaled hours
                insert_pos = idx + 1
                extra: List[Phase] = []

                # 重采毛坯工时按比例缩放，添加缩进使其更醒目
                lead_hours = self.order.blank_lead_days * self.cal.working_hours_per_day * ratio
                if lead_hours > 0:
                    extra.append(Phase(f"    ↻ 重采毛坯(报废{scrap_qty}件) - {self.order.blank_lead_days}天×{ratio:.1%}", lead_hours))

                # 车床工序链按比例缩放，添加缩进和标记
                base_chain = build_lathe_chain(self.order.lathe_ops)
                for p in base_chain:
                    scaled_hours = p.planned_hours * ratio
                    extra.append(Phase(f"    ↻ {p.name}(补{scrap_qty}件)", scaled_hours))

                self.order.phases[insert_pos:insert_pos] = extra

                self._reload_phase_tree()
                self.refresh_eta()
                self._explain(
                    f"在'{phase.name}'处报废 {ratio:.1%} ({scrap_qty}件)。"
                    f"已插入重做工序，工时按比例缩放。"
                )
                
                scrap_dialog.destroy()
                messagebox.showinfo("完成", f"已添加报废重做工序\n报废数量: {scrap_qty} 件\n总新增工时已按 {ratio:.1%} 比例缩放")
                
            except ValueError as e:
                messagebox.showerror("输入错误", str(e))

        btn_frame = ttk.Frame(scrap_dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确认", command=confirm_scrap).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=scrap_dialog.destroy).pack(side="left", padx=5)

    def refresh_eta(self):
        if not self.order:
            return
        try:
            result = compute_eta(self.order, self.cal)
        except Exception as e:
            messagebox.showerror("交期计算错误", str(e))
            return

        eta_dt: datetime = result["eta_dt"]
        remaining_hours: float = result["remaining_hours"]
        explanation: List[str] = result["explanation"]

        self.eta_var.set(eta_dt.strftime("%Y-%m-%d %H:%M"))
        self.remaining_var.set(f"剩余计划工时: {remaining_hours:g}小时")

        self.explain_text.delete("1.0", tk.END)
        for line in explanation[:300]:
            self.explain_text.insert(tk.END, line + "\n")

    def _explain(self, line: str):
        self.explain_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {line}\n")
        self.explain_text.see(tk.END)

    def save_order(self):
        """保存订单数据到JSON文件，可自定义文件名"""
        if not self.order:
            messagebox.showwarning("无订单", "没有可保存的订单。")
            return
        
        # 弹出对话框让用户输入文件名
        save_dialog = tk.Toplevel(self)
        save_dialog.title("保存订单")
        save_dialog.geometry("400x150")
        save_dialog.transient(self)
        save_dialog.grab_set()

        ttk.Label(save_dialog, text="保存为:").pack(pady=10)
        
        filename_var = tk.StringVar(value=f"{self.order.order_id}.json")
        ttk.Entry(save_dialog, textvariable=filename_var, width=40).pack(pady=5)
        
        ttk.Label(save_dialog, text="(文件将保存在桌面)", font=("", 9), foreground="gray").pack()

        def do_save():
            filename = filename_var.get().strip()
            if not filename:
                messagebox.showerror("错误", "文件名不能为空")
                return
            if not filename.endswith('.json'):
                filename += '.json'
            
            try:
                data = {
                    "order_id": self.order.order_id,
                    "start_dt": self.order.start_dt.isoformat(),
                    "lathe_ops": self.order.lathe_ops,
                    "blank_lead_days": self.order.blank_lead_days,
                    "quantity": self.order.quantity,
                    "route_mode": self.route_mode,
                    "phases": [
                        {
                            "name": p.name,
                            "planned_hours": p.planned_hours,
                            "done": p.done,
                            "parallel_group": p.parallel_group
                        } for p in self.order.phases
                    ],
                    "events": [
                        {
                            "day": e.day.isoformat(),
                            "hours_lost": e.hours_lost,
                            "reason": e.reason
                        } for e in self.order.events
                    ]
                }
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                save_dialog.destroy()
                messagebox.showinfo("保存成功", f"订单已保存到 {filename}")
                self._explain(f"订单已保存到: {filename}")
            except Exception as e:
                messagebox.showerror("保存失败", f"保存时出错: {str(e)}")

        btn_frame = ttk.Frame(save_dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="保存", command=do_save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=save_dialog.destroy).pack(side="left", padx=5)

    def _load_order(self, show_message=False, filename=None):
        """从JSON文件加载订单数据"""
        if filename is None:
            filename = self.save_file
            
        if not os.path.exists(filename):
            if show_message:
                messagebox.showinfo("提示", f"没有找到文件: {filename}")
            return  # 文件不存在
        
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 重建订单对象
            phases = [Phase(
                name=p["name"],
                planned_hours=p["planned_hours"],
                done=p["done"],
                parallel_group=p.get("parallel_group", 0)  # 兼容旧版本
            ) for p in data["phases"]]
            
            events = [Event(
                day=date.fromisoformat(e["day"]),
                hours_lost=e["hours_lost"],
                reason=e["reason"]
            ) for e in data["events"]]
            
            self.order = Order(
                order_id=data["order_id"],
                start_dt=datetime.fromisoformat(data["start_dt"]),
                phases=phases,
                events=events,
                lathe_ops=data["lathe_ops"],
                blank_lead_days=data["blank_lead_days"],
                quantity=data.get("quantity", 1)  # 兼容旧版本
            )
            
            # 更新UI
            self.route_mode = data.get("route_mode", "with_mold")
            self.order_id_var.set(data["order_id"])
            self.lathe_ops_var.set(str(data["lathe_ops"]))
            self.blank_days_var.set(str(data["blank_lead_days"]))
            self.quantity_var.set(str(data.get("quantity", 1)))
            self.route_var.set(self.route_mode)
            
            self._reload_phase_tree()
            self._reload_event_list()
            self.refresh_eta()
            
            if show_message:
                messagebox.showinfo("加载成功", f"已成功加载订单: {self.order.order_id}\n数量: {self.order.quantity} 件\n工序数: {len(self.order.phases)}\n事件数: {len(self.order.events)}")
            self._explain(f"已加载订单: {self.order.order_id} ({self.order.quantity}件)")
        except Exception as e:
            error_msg = f"加载订单时出错: {str(e)}"
            if show_message:
                messagebox.showerror("加载失败", error_msg)
            else:
                print(error_msg)  # 启动时的错误输出到控制台

    def load_order_button(self):
        """点击加载按钮时调用，显示文件选择对话框"""
        # 列出所有JSON文件
        json_files = [f for f in os.listdir('.') if f.endswith('.json')]
        
        if not json_files:
            messagebox.showinfo("提示", "当前目录没有找到JSON订单文件。")
            return
        
        # 创建选择对话框
        load_dialog = tk.Toplevel(self)
        load_dialog.title("选择订单文件")
        load_dialog.geometry("450x400")
        load_dialog.transient(self)
        load_dialog.grab_set()

        ttk.Label(load_dialog, text="请选择要加载的订单:", font=("", 10, "bold")).pack(pady=10)
        
        # 文件列表
        listbox_frame = ttk.Frame(load_dialog)
        listbox_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")
        
        file_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, font=("", 10))
        file_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=file_listbox.yview)
        
        for f in sorted(json_files):
            file_listbox.insert(tk.END, f)
        
        # 显示文件预览
        preview_label = ttk.Label(load_dialog, text="", foreground="blue", wraplength=400)
        preview_label.pack(pady=5)

        def on_select(event):
            if file_listbox.curselection():
                filename = file_listbox.get(file_listbox.curselection()[0])
                try:
                    with open(filename, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    preview_label.config(text=f"订单: {data.get('order_id', '?')} | 数量: {data.get('quantity', '?')} 件 | 工序: {len(data.get('phases', []))} 个")
                except:
                    preview_label.config(text="无法读取文件信息")

        file_listbox.bind('<<ListboxSelect>>', on_select)

        def do_load():
            if not file_listbox.curselection():
                messagebox.showwarning("未选择", "请先选择一个文件")
                return
            filename = file_listbox.get(file_listbox.curselection()[0])
            load_dialog.destroy()
            self._load_order(show_message=True, filename=filename)

        btn_frame = ttk.Frame(load_dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="加载", command=do_load).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=load_dialog.destroy).pack(side="left", padx=5)

if __name__ == "__main__":
    app = ETAGUI()
    app.mainloop()
