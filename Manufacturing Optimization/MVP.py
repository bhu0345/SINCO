import json
import os
from dataclasses import dataclass, field
from datetime import datetime, timedelta, date
from typing import List, Dict, Optional
from PySide6 import QtCore, QtGui, QtWidgets


QT_BINDING = "PySide6"
ADMIN_PASSWORD = "admin123"
ADMIN_SECRET_KEY = "change-me"


# ----------------------------
# Data models
# ----------------------------


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
    phases: List[Phase] = field(default_factory=list)


@dataclass
class Equipment:
    equipment_id: str
    category: str = ""
    total_count: int = 1
    available_count: int = 1


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


@dataclass
class Order:
    order_id: str
    start_dt: datetime
    products: List[Product] = field(default_factory=list)
    events: List[Event] = field(default_factory=list)
    defects: List[DefectRecord] = field(default_factory=list)
    equipment: List[Equipment] = field(default_factory=list)
    employees: List[str] = field(default_factory=list)
    logs: List[LogEntry] = field(default_factory=list)


@dataclass
class UserAccount:
    username: str
    password: str


# ----------------------------
# Scheduling / ETA computation
# ----------------------------

class WorkCalendar:
    def __init__(self, shift_template: Optional[ShiftTemplate] = None):
        self.shift_template = shift_template

    def capacity_for_day(self, d: date) -> float:
        if not self.shift_template:
            return 0.0
        return self.shift_template.hours_for_weekday(d.weekday())


def _equipment_available_map(order: Order) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for eq in order.equipment:
        if not eq.equipment_id:
            continue
        result[eq.equipment_id] = max(0, int(eq.available_count))
    return result


def _phase_effective_hours(phase: Phase, quantity: int, equipment_map: Dict[str, int]) -> float:
    qty = max(int(quantity), 1)
    base = phase.planned_hours * qty
    hours = base
    if phase.equipment_id:
        available = equipment_map.get(phase.equipment_id, 1)
        if available > 0:
            hours = hours / available
    return hours


def _phase_total_hours(phase: Phase, quantity: int) -> float:
    return max(0.0, phase.planned_hours) * max(int(quantity), 1)


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


def compute_eta(order: Order, cal: WorkCalendar) -> Dict[str, object]:
    equipment_map = _equipment_available_map(order)
    remaining_hours = sum(_product_remaining_hours(p, equipment_map) for p in order.products)

    lost_map: Dict[date, float] = {}
    reason_map: Dict[date, List[str]] = {}
    for ev in order.events:
        lost_map[ev.day] = lost_map.get(ev.day, 0.0) + ev.hours_lost
        reason_map.setdefault(ev.day, []).append(f"{ev.reason}(-{ev.hours_lost:g}h)")

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
            "explanation": explanation + ["All phases completed. ETA equals start time."]
        }

    current_day = order.start_dt.date()
    hours_left = remaining_hours
    daily_capacity_map: Dict[date, float] = {}

    for _ in range(3650):
        base_cap = cal.capacity_for_day(current_day)
        if base_cap > 0:
            lost = lost_map.get(current_day, 0.0)
            cap = max(base_cap - lost, 0.0)
            daily_capacity_map[current_day] = cap

            if lost > 0:
                explanation.append(
                    f"{current_day.isoformat()}: capacity {base_cap:g}h - lost {lost:g}h => {cap:g}h "
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
                        "explanation": explanation or ["No blocking events."]
                    }
                hours_left -= cap
        current_day = current_day + timedelta(days=1)

    raise RuntimeError("ETA computation exceeded safe bounds.")


# ----------------------------
# Qt UI
# ----------------------------


def _qdate_to_date(qdate: QtCore.QDate) -> date:
    return date(qdate.year(), qdate.month(), qdate.day())


def _date_to_qdate(d: date) -> QtCore.QDate:
    return QtCore.QDate(d.year, d.month, d.day)


def _chinese_locale() -> QtCore.QLocale:
    try:
        return QtCore.QLocale(QtCore.QLocale.Language.Chinese, QtCore.QLocale.Country.China)
    except AttributeError:
        return QtCore.QLocale(QtCore.QLocale.Chinese, QtCore.QLocale.China)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("生产交期优化系统 V3.0.1")
        self.resize(1280, 820)

        self.locale_cn = _chinese_locale()
        QtCore.QLocale.setDefault(self.locale_cn)

        self.order: Optional[Order] = None
        self.equipment: List[Equipment] = []
        self.employees: List[str] = []
        self.event_reasons = ["员工请假", "设备故障", "停电", "材料短缺", "质量问题", "其他"]
        self.user_accounts: List[UserAccount] = []
        self.current_user = ""
        self.autosave_path: Optional[str] = None
        self.app_template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "app_templates.json"
        )
        self.equipment_categories = ["车床", "加工中心", "检验设备", "辅助设备"]
        self.defect_categories = ["设备", "原材料", "员工"]
        self.logo_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "GUI", "sinco.JPG"
        )
        self.equipment_templates: List[Equipment] = []
        self.phase_templates: List[Phase] = []
        self.employee_templates: List[str] = []
        self.shift_templates: List[ShiftTemplate] = []
        self.active_shift_template_name = ""
        self.last_eta_dt: Optional[datetime] = None
        self.last_remaining_hours: float = 0.0
        self.active_product_id: str = ""
        self._load_app_templates()
        self._ensure_default_templates()
        self.cal = WorkCalendar(self._current_shift_template())

        self._apply_app_font()
        self._apply_modern_style()

        self.stack = QtWidgets.QStackedWidget()
        self.setCentralWidget(self.stack)

        self.login_page = self._build_login_page()
        self.dashboard = self._build_dashboard()
        self.detail_page = self._build_detail_page()
        self.admin_page = self._build_admin_page()
        self.visual_page = self._build_visual_page()

        self.stack.addWidget(self.login_page)
        self.stack.addWidget(self.dashboard)
        self.stack.addWidget(self.detail_page)
        self.stack.addWidget(self.admin_page)
        self.stack.addWidget(self.visual_page)
        self.stack.setCurrentWidget(self.login_page)

        self.statusBar().showMessage(f"Using {QT_BINDING}")

    # ------------------------
    # Dashboard UI
    # ------------------------

    def _setup_date_edit(self, date_edit: QtWidgets.QDateEdit) -> None:
        date_edit.setCalendarPopup(True)
        date_edit.setLocale(self.locale_cn)
        date_edit.setDisplayFormat("yyyy年MM月dd日")
        calendar = date_edit.calendarWidget()
        if calendar:
            calendar.setLocale(self.locale_cn)

    def _configure_combo_popup(self, combo: QtWidgets.QComboBox, min_width: int = 180) -> None:
        combo.setMinimumWidth(min_width)
        view = combo.view()
        view.setMinimumWidth(min_width)
        view.setTextElideMode(QtCore.Qt.TextElideMode.ElideNone)

    def _update_phase_product_label(self, product: Optional[Product]) -> None:
        if not hasattr(self, "phase_product_label"):
            return
        if not product:
            self.phase_product_label.setText("当前产品: -")
            return
        part = product.part_number or "-"
        self.phase_product_label.setText(
            f"当前产品: {product.product_id} | 零件号: {part} | 数量: {product.quantity}"
        )

    def _select_product_by_id(self, product_id: str) -> None:
        if not self.order or not product_id:
            return
        for idx, product in enumerate(self.order.products):
            if product.product_id == product_id:
                self.products_table.setCurrentCell(idx, 0)
                self.products_table.selectRow(idx)
                return

    def _apply_app_font(self) -> None:
        preferred = [
            "PingFang SC",
            "Source Han Sans SC",
            "Noto Sans CJK SC",
            "Microsoft YaHei",
        ]
        db = QtGui.QFontDatabase()
        for family in preferred:
            if family in db.families():
                app = QtWidgets.QApplication.instance()
                if app:
                    app.setFont(QtGui.QFont(family, 10))
                break

    def _apply_modern_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #f2f6fb; }
            QFrame#heroCard, QFrame#metricCard { background: #ffffff; border: 1px solid #d8e1ec; border-radius: 12px; }
            QGroupBox { background: #ffffff; border: 1px solid #d8e1ec; border-radius: 12px; margin-top: 14px; }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #2a3b4d; font-weight: 600; }
            QLabel#heroTitle { font-size: 20px; font-weight: 700; color: #1f2d3d; }
            QLabel#heroSubtitle { color: #5c6f84; }
            QLabel#metricValue { font-size: 18px; font-weight: 700; color: #0f5e9c; }
            QLabel#metricLabel { color: #6b7c93; }
            QPushButton { background: #e8eef5; border: 1px solid #c8d4e2; padding: 6px 12px; border-radius: 6px; }
            QPushButton#primaryAction { background: #0f5e9c; color: #ffffff; border: none; }
            QPushButton#dangerAction { background: #c64545; color: #ffffff; border: none; }
            QLineEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox {
                background: #f7f9fb; border: 1px solid #cfd9e6; padding: 4px 6px; border-radius: 6px;
            }
            QComboBox QAbstractItemView {
                background: #ffffff; border: 1px solid #cfd9e6; selection-background-color: #d6e6f7;
                selection-color: #1f2d3d; outline: 0;
            }
            QComboBox QAbstractItemView::item { padding: 4px 8px; color: #1f2d3d; }
            QComboBox QAbstractItemView::item:hover { background: #e3edf8; color: #1f2d3d; }
            QComboBox QAbstractItemView::item:selected { background: #cfe0f4; color: #1f2d3d; }
            QSplitter::handle:vertical { background: #d8e1ec; height: 8px; }
            QSplitter::handle:horizontal { background: #d8e1ec; width: 8px; }
            QTableWidget {
                background: #ffffff; border: 1px solid #d8e1ec; border-radius: 10px; gridline-color: #e1e7ef;
            }
            QHeaderView::section {
                background: #edf2f7; padding: 6px 8px; border: none; color: #2a3b4d; font-weight: 600;
            }
            QProgressBar { border: 1px solid #cfd9e6; border-radius: 6px; text-align: center; height: 16px; background: #eef2f7; }
            QProgressBar::chunk { background: #2f7dd1; border-radius: 6px; }
            QProgressBar[complete="true"]::chunk { background: #2f855a; }
            QTabWidget::pane { border: 1px solid #d8e1ec; border-radius: 10px; }
            QTabBar::tab {
                background: #eef2f7; padding: 6px 12px; border: 1px solid #d8e1ec; border-bottom: none;
                border-top-left-radius: 8px; border-top-right-radius: 8px; margin-right: 4px;
            }
            QTabBar::tab:selected { background: #ffffff; }
            """
        )

    def _build_login_page(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.addStretch(1)

        card = QtWidgets.QFrame()
        card.setObjectName("heroCard")
        card_layout = QtWidgets.QGridLayout(card)
        card_layout.setContentsMargins(24, 20, 24, 20)
        card_layout.setSpacing(12)

        title = QtWidgets.QLabel("用户登录")
        title.setObjectName("heroTitle")
        card_layout.addWidget(title, 0, 0, 1, 2)

        self.login_user_edit = QtWidgets.QLineEdit()
        self.login_user_edit.setPlaceholderText("用户名")
        self.login_pass_edit = QtWidgets.QLineEdit()
        self.login_pass_edit.setPlaceholderText("密码")
        self.login_pass_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.login_btn = QtWidgets.QPushButton("登录")
        self.login_btn.setObjectName("primaryAction")
        self.login_btn.clicked.connect(self.attempt_login)
        self.login_user_edit.returnPressed.connect(self.attempt_login)
        self.login_pass_edit.returnPressed.connect(self.attempt_login)
        self.login_status_label = QtWidgets.QLabel("")
        self.login_status_label.setObjectName("metricLabel")

        card_layout.addWidget(QtWidgets.QLabel("用户名"), 1, 0)
        card_layout.addWidget(self.login_user_edit, 1, 1)
        card_layout.addWidget(QtWidgets.QLabel("密码"), 2, 0)
        card_layout.addWidget(self.login_pass_edit, 2, 1)
        card_layout.addWidget(self.login_status_label, 3, 0, 1, 2)
        card_layout.addWidget(self.login_btn, 4, 0, 1, 2)

        self.login_user_edit.setFocus()
        layout.addWidget(card, 0, QtCore.Qt.AlignmentFlag.AlignHCenter)
        layout.addStretch(2)
        return page

    def attempt_login(self) -> None:
        username = self.login_user_edit.text().strip()
        password = self.login_pass_edit.text()
        if not username or not password:
            self.login_status_label.setText("请输入用户名和密码。")
            return
        if not self._validate_user(username, password):
            self.login_status_label.setText("用户名或密码不正确。")
            return
        self.current_user = username
        self.login_status_label.setText("")
        self.login_pass_edit.clear()
        self.statusBar().showMessage(f"当前用户: {username}", 3000)
        self.stack.setCurrentWidget(self.dashboard)

    def _validate_user(self, username: str, password: str) -> bool:
        return any(
            account.username == username and account.password == password
            for account in self.user_accounts
        )

    def _toggle_admin_password_visibility(self, checked: bool) -> None:
        if not hasattr(self, "admin_user_pass_edit"):
            return
        username = ""
        if hasattr(self, "admin_user_name_edit"):
            username = self.admin_user_name_edit.text().strip()
        if username == "admin":
            self.admin_user_pass_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            return
        mode = (
            QtWidgets.QLineEdit.EchoMode.Normal
            if checked
            else QtWidgets.QLineEdit.EchoMode.Password
        )
        self.admin_user_pass_edit.setEchoMode(mode)

    def _toggle_form_section(
        self,
        button: QtWidgets.QPushButton,
        form_widget: QtWidgets.QWidget,
        label: str,
    ) -> None:
        visible = not form_widget.isVisible()
        form_widget.setVisible(visible)
        button.setText(f"{'收起' if visible else '编辑'}{label}")

    def _load_logo_pixmap(self, label: QtWidgets.QLabel) -> None:
        if not os.path.exists(self.logo_path):
            label.setText("SINCO")
            label.setObjectName("heroTitle")
            return
        pixmap = QtGui.QPixmap(self.logo_path)
        if pixmap.isNull():
            label.setText("SINCO")
            label.setObjectName("heroTitle")
            return
        scaled = pixmap.scaled(
            140,
            60,
            QtCore.Qt.AspectRatioMode.KeepAspectRatio,
            QtCore.Qt.TransformationMode.SmoothTransformation,
        )
        label.setPixmap(scaled)

    def _build_dashboard(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        hero = QtWidgets.QFrame()
        hero.setObjectName("heroCard")
        hero_layout = QtWidgets.QHBoxLayout(hero)
        hero_layout.setContentsMargins(16, 12, 16, 12)
        hero_layout.setSpacing(16)

        self.logo_label = QtWidgets.QLabel()
        self.logo_label.setMinimumWidth(140)
        self.logo_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter)
        self._load_logo_pixmap(self.logo_label)
        hero_layout.addWidget(self.logo_label)

        title_layout = QtWidgets.QVBoxLayout()
        self.hero_title_label = QtWidgets.QLabel("SINCO 生产交期优化系统")
        self.hero_title_label.setObjectName("heroTitle")
        self.hero_subtitle_label = QtWidgets.QLabel("订单、设备、班次与进度一屏掌控")
        self.hero_subtitle_label.setObjectName("heroSubtitle")
        title_layout.addWidget(self.hero_title_label)
        title_layout.addWidget(self.hero_subtitle_label)
        hero_layout.addLayout(title_layout, 1)

        stats_layout = QtWidgets.QVBoxLayout()
        self.shift_summary_label = QtWidgets.QLabel("当前班次: -")
        self.shift_summary_label.setObjectName("metricLabel")
        self.dashboard_summary_label = QtWidgets.QLabel("产品数: 0 | 设备数: 0")
        self.dashboard_summary_label.setObjectName("metricLabel")
        stats_layout.addWidget(self.shift_summary_label)
        stats_layout.addWidget(self.dashboard_summary_label)
        hero_layout.addLayout(stats_layout)

        layout.addWidget(hero)

        header = QtWidgets.QGroupBox("订单管理")
        header_layout = QtWidgets.QGridLayout(header)

        self.order_id_edit = QtWidgets.QLineEdit()
        self.order_id_edit.setPlaceholderText("O-001")
        self.start_date_edit = QtWidgets.QDateEdit()
        self.start_date_edit.setDate(QtCore.QDate.currentDate())
        self._setup_date_edit(self.start_date_edit)

        header_layout.addWidget(QtWidgets.QLabel("订单编号"), 0, 0)
        header_layout.addWidget(self.order_id_edit, 0, 1)
        header_layout.addWidget(QtWidgets.QLabel("开始日期"), 0, 2)
        header_layout.addWidget(self.start_date_edit, 0, 3)

        self.create_order_btn = QtWidgets.QPushButton("创建订单")
        self.create_order_btn.setObjectName("primaryAction")
        self.create_order_btn.clicked.connect(self.create_order)
        self.save_order_btn = QtWidgets.QPushButton("保存订单")
        self.save_order_btn.clicked.connect(self.save_order)
        self.load_order_btn = QtWidgets.QPushButton("加载订单")
        self.load_order_btn.clicked.connect(self.load_order)
        self.goto_detail_btn = QtWidgets.QPushButton("进入订单详情")
        self.goto_detail_btn.clicked.connect(self.go_to_detail)
        self.visual_btn = QtWidgets.QPushButton("数据看板")
        self.visual_btn.clicked.connect(self.go_to_visuals)
        self.admin_btn = QtWidgets.QPushButton("管理员")
        self.admin_btn.clicked.connect(self.open_admin_login)
        self.switch_user_btn = QtWidgets.QPushButton("切换用户")
        self.switch_user_btn.clicked.connect(self.switch_user)
        self.logout_btn = QtWidgets.QPushButton("退出登录")
        self.logout_btn.clicked.connect(self.logout_user)

        header_layout.addWidget(self.create_order_btn, 0, 4)
        header_layout.addWidget(self.save_order_btn, 0, 5)
        header_layout.addWidget(self.load_order_btn, 0, 6)
        header_layout.addWidget(self.goto_detail_btn, 0, 7)
        header_layout.addWidget(self.visual_btn, 0, 8)
        header_layout.addWidget(self.admin_btn, 0, 9)
        header_layout.addWidget(self.switch_user_btn, 0, 10)
        header_layout.addWidget(self.logout_btn, 0, 11)

        layout.addWidget(header)

        mid = QtWidgets.QHBoxLayout()
        layout.addLayout(mid)

        equipment_group = QtWidgets.QGroupBox("设备可用性")
        equipment_layout = QtWidgets.QVBoxLayout(equipment_group)

        self.equipment_table = QtWidgets.QTableWidget(0, 4)
        self.equipment_table.setHorizontalHeaderLabels(["设备编号", "类别", "总数量", "可用数量"])
        self.equipment_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.equipment_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.equipment_table.verticalHeader().setVisible(False)
        self.equipment_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.equipment_table.itemSelectionChanged.connect(self.on_equipment_select)

        equipment_layout.addWidget(self.equipment_table)

        eq_form = QtWidgets.QGridLayout()
        eq_form.setColumnStretch(3, 2)
        self.equipment_id_edit = QtWidgets.QLineEdit()
        self.equipment_category_combo = QtWidgets.QComboBox()
        self.equipment_category_combo.setEditable(False)
        self.equipment_category_combo.addItems(self.equipment_categories)
        self._configure_combo_popup(self.equipment_category_combo, 200)
        self.equipment_total_spin = QtWidgets.QSpinBox()
        self.equipment_total_spin.setRange(1, 9999)
        self.equipment_available_spin = QtWidgets.QSpinBox()
        self.equipment_available_spin.setRange(0, 9999)

        eq_form.addWidget(QtWidgets.QLabel("设备编号"), 0, 0)
        eq_form.addWidget(self.equipment_id_edit, 0, 1)
        eq_form.addWidget(QtWidgets.QLabel("类别"), 0, 2)
        eq_form.addWidget(self.equipment_category_combo, 0, 3)
        eq_form.addWidget(QtWidgets.QLabel("总数量"), 0, 4)
        eq_form.addWidget(self.equipment_total_spin, 0, 5)
        eq_form.addWidget(QtWidgets.QLabel("可用数量"), 0, 6)
        eq_form.addWidget(self.equipment_available_spin, 0, 7)

        self.eq_add_btn = QtWidgets.QPushButton("添加/更新设备")
        self.eq_add_btn.clicked.connect(self.add_or_update_equipment)
        self.eq_remove_btn = QtWidgets.QPushButton("删除设备")
        self.eq_remove_btn.clicked.connect(self.remove_equipment)

        eq_form.addWidget(self.eq_add_btn, 1, 6)
        eq_form.addWidget(self.eq_remove_btn, 1, 7)

        equipment_layout.addLayout(eq_form)

        mid.addWidget(equipment_group, 2)

        progress_group = QtWidgets.QGroupBox("进度与交期")
        progress_layout = QtWidgets.QVBoxLayout(progress_group)

        self.overall_progress = QtWidgets.QProgressBar()
        self.overall_progress.setValue(0)
        progress_layout.addWidget(QtWidgets.QLabel("订单总体进度"))
        progress_layout.addWidget(self.overall_progress)

        self.product_progress_table = QtWidgets.QTableWidget(0, 3)
        self.product_progress_table.setHorizontalHeaderLabels(["产品", "零件号", "进度"])
        self.product_progress_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.product_progress_table.verticalHeader().setVisible(False)
        self.product_progress_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        progress_layout.addWidget(self.product_progress_table)

        metrics_card = QtWidgets.QFrame()
        metrics_card.setObjectName("metricCard")
        metrics_layout = QtWidgets.QGridLayout(metrics_card)
        metrics_layout.setContentsMargins(12, 8, 12, 8)

        eta_label = QtWidgets.QLabel("预计交期")
        eta_label.setObjectName("metricLabel")
        self.eta_value = QtWidgets.QLabel("-")
        self.eta_value.setObjectName("metricValue")

        remaining_label = QtWidgets.QLabel("剩余工时")
        remaining_label.setObjectName("metricLabel")
        self.remaining_value = QtWidgets.QLabel("-")
        self.remaining_value.setObjectName("metricValue")

        metrics_layout.addWidget(eta_label, 0, 0)
        metrics_layout.addWidget(self.eta_value, 1, 0)
        metrics_layout.addWidget(remaining_label, 0, 1)
        metrics_layout.addWidget(self.remaining_value, 1, 1)
        progress_layout.addWidget(metrics_card)

        self.refresh_eta_btn = QtWidgets.QPushButton("刷新交期")
        self.refresh_eta_btn.setObjectName("primaryAction")
        self.refresh_eta_btn.clicked.connect(self.refresh_eta)
        progress_layout.addWidget(self.refresh_eta_btn)

        mid.addWidget(progress_group, 3)

        return page

    # ------------------------
    # Detail Page UI
    # ------------------------

    def _build_detail_page(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)

        top_bar = QtWidgets.QHBoxLayout()
        self.back_btn = QtWidgets.QPushButton("返回主界面")
        self.back_btn.clicked.connect(self.go_to_dashboard)
        self.detail_visual_btn = QtWidgets.QPushButton("数据看板")
        self.detail_visual_btn.clicked.connect(self.go_to_visuals)
        self.detail_switch_user_btn = QtWidgets.QPushButton("切换用户")
        self.detail_switch_user_btn.clicked.connect(self.switch_user)
        self.detail_logout_btn = QtWidgets.QPushButton("退出登录")
        self.detail_logout_btn.clicked.connect(self.logout_user)
        self.order_summary_label = QtWidgets.QLabel("订单: -")
        self.detail_eta_label = QtWidgets.QLabel("预计交期: -")
        self.detail_progress_label = QtWidgets.QLabel("总体进度: 0%")
        self.detail_eta_label.setObjectName("metricValue")
        self.detail_progress_label.setObjectName("metricLabel")

        top_bar.addWidget(self.back_btn)
        top_bar.addWidget(self.detail_visual_btn)
        top_bar.addWidget(self.detail_switch_user_btn)
        top_bar.addWidget(self.detail_logout_btn)
        top_bar.addWidget(self.order_summary_label)
        top_bar.addStretch(1)
        top_bar.addWidget(self.detail_progress_label)
        top_bar.addWidget(self.detail_eta_label)

        layout.addLayout(top_bar)

        splitter = QtWidgets.QSplitter()
        layout.addWidget(splitter, 1)

        # Left panel: Products + Employees
        left_panel = QtWidgets.QWidget()
        left_panel.setMinimumWidth(280)
        left_layout = QtWidgets.QVBoxLayout(left_panel)

        products_group = QtWidgets.QGroupBox("产品列表")
        products_layout = QtWidgets.QVBoxLayout(products_group)

        self.products_table = QtWidgets.QTableWidget(0, 4)
        self.products_table.setHorizontalHeaderLabels(["产品", "零件号", "数量", "进度"])
        self.products_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.products_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.products_table.verticalHeader().setVisible(False)
        self.products_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.products_table.itemSelectionChanged.connect(self.on_product_select)

        products_layout.addWidget(self.products_table)

        prod_form = QtWidgets.QGridLayout()
        self.product_id_edit = QtWidgets.QLineEdit()
        self.product_id_edit.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        self.product_part_edit = QtWidgets.QLineEdit()
        self.product_part_edit.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        self.product_qty_spin = QtWidgets.QSpinBox()
        self.product_qty_spin.setRange(1, 999999)
        self.product_qty_spin.setMaximumWidth(110)

        prod_form.setColumnStretch(1, 3)
        prod_form.setColumnStretch(3, 2)
        prod_form.addWidget(QtWidgets.QLabel("产品名"), 0, 0)
        prod_form.addWidget(self.product_id_edit, 0, 1)
        prod_form.addWidget(QtWidgets.QLabel("零件号"), 0, 2)
        prod_form.addWidget(self.product_part_edit, 0, 3)
        prod_form.addWidget(QtWidgets.QLabel("数量"), 0, 4)
        prod_form.addWidget(self.product_qty_spin, 0, 5)

        self.product_add_btn = QtWidgets.QPushButton("添加/更新产品")
        self.product_add_btn.clicked.connect(self.add_or_update_product)
        self.product_remove_btn = QtWidgets.QPushButton("删除产品")
        self.product_remove_btn.clicked.connect(self.remove_product)

        prod_form.addWidget(self.product_add_btn, 1, 4)
        prod_form.addWidget(self.product_remove_btn, 1, 5)

        products_layout.addLayout(prod_form)
        left_layout.addWidget(products_group)

        employees_group = QtWidgets.QGroupBox("员工列表")
        employees_layout = QtWidgets.QVBoxLayout(employees_group)

        self.employee_list = QtWidgets.QListWidget()
        employees_layout.addWidget(self.employee_list)

        emp_form = QtWidgets.QHBoxLayout()
        self.employee_name_edit = QtWidgets.QLineEdit()
        self.employee_name_edit.setPlaceholderText("员工姓名")
        self.employee_add_btn = QtWidgets.QPushButton("添加")
        self.employee_add_btn.clicked.connect(self.add_employee)
        self.employee_remove_btn = QtWidgets.QPushButton("删除")
        self.employee_remove_btn.clicked.connect(self.remove_employee)

        emp_form.addWidget(self.employee_name_edit)
        emp_form.addWidget(self.employee_add_btn)
        emp_form.addWidget(self.employee_remove_btn)

        employees_layout.addLayout(emp_form)
        left_layout.addWidget(employees_group)

        splitter.addWidget(left_panel)

        # Right panel: Phases + Events
        right_panel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_panel)
        right_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)
        right_splitter.setChildrenCollapsible(False)
        right_splitter.setHandleWidth(8)

        phases_group = QtWidgets.QGroupBox("工序设置")
        phases_layout = QtWidgets.QVBoxLayout(phases_group)
        phases_toolbar = QtWidgets.QHBoxLayout()
        self.phase_product_label = QtWidgets.QLabel("当前产品: -")
        self.phase_product_label.setObjectName("metricLabel")
        phases_toolbar.addWidget(self.phase_product_label)
        phases_toolbar.addStretch(1)
        self.phase_form_toggle_btn = QtWidgets.QPushButton("编辑工序")
        self.phase_form_toggle_btn.clicked.connect(
            lambda: self._toggle_form_section(
                self.phase_form_toggle_btn,
                self.phase_form_widget,
                "工序",
            )
        )
        phases_toolbar.addWidget(self.phase_form_toggle_btn)
        phases_layout.addLayout(phases_toolbar)

        self.phases_table = QtWidgets.QTableWidget(0, 6)
        self.phases_table.setHorizontalHeaderLabels(
            ["工序名称", "工时", "设备", "员工", "并行组", "进度"]
        )
        self.phases_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.phases_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.phases_table.verticalHeader().setVisible(False)
        self.phases_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.phases_table.itemSelectionChanged.connect(self.on_phase_select)

        phases_layout.addWidget(self.phases_table)

        self.phase_form_widget = QtWidgets.QWidget()
        phase_form = QtWidgets.QGridLayout(self.phase_form_widget)
        self.phase_name_edit = QtWidgets.QLineEdit()
        self.phase_hours_spin = QtWidgets.QDoubleSpinBox()
        self.phase_hours_spin.setRange(0, 99999)
        self.phase_hours_spin.setDecimals(2)
        self.phase_equipment_combo = QtWidgets.QComboBox()
        self.phase_equipment_combo.setEditable(True)
        self.phase_employee_combo = QtWidgets.QComboBox()
        self.phase_employee_combo.setEditable(True)
        self.phase_parallel_spin = QtWidgets.QSpinBox()
        self.phase_parallel_spin.setRange(0, 9999)
        self.phase_completed_spin = QtWidgets.QDoubleSpinBox()
        self.phase_completed_spin.setRange(0, 999999)
        self.phase_completed_spin.setDecimals(2)

        phase_form.addWidget(QtWidgets.QLabel("名称"), 0, 0)
        phase_form.addWidget(self.phase_name_edit, 0, 1)
        phase_form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        phase_form.addWidget(self.phase_hours_spin, 0, 3)
        phase_form.addWidget(QtWidgets.QLabel("设备"), 0, 4)
        phase_form.addWidget(self.phase_equipment_combo, 0, 5)

        phase_form.addWidget(QtWidgets.QLabel("员工"), 1, 0)
        phase_form.addWidget(self.phase_employee_combo, 1, 1)
        phase_form.addWidget(QtWidgets.QLabel("并行组"), 1, 2)
        phase_form.addWidget(self.phase_parallel_spin, 1, 3)
        phase_form.addWidget(QtWidgets.QLabel("完成工时"), 1, 4)
        phase_form.addWidget(self.phase_completed_spin, 1, 5)

        self.phase_add_btn = QtWidgets.QPushButton("添加/更新工序")
        self.phase_add_btn.clicked.connect(self.add_or_update_phase)
        self.phase_remove_btn = QtWidgets.QPushButton("删除工序")
        self.phase_remove_btn.clicked.connect(self.remove_phase)
        self.phase_parallel_btn = QtWidgets.QPushButton("设为并行组")
        self.phase_parallel_btn.clicked.connect(self.set_parallel_group)
        self.phase_parallel_clear_btn = QtWidgets.QPushButton("取消并行")
        self.phase_parallel_clear_btn.clicked.connect(self.clear_parallel_group)

        phase_form.addWidget(self.phase_add_btn, 2, 4)
        phase_form.addWidget(self.phase_remove_btn, 2, 5)
        phase_form.addWidget(self.phase_parallel_btn, 3, 4)
        phase_form.addWidget(self.phase_parallel_clear_btn, 3, 5)

        phases_layout.addWidget(self.phase_form_widget)
        self.phase_form_widget.setVisible(False)
        right_splitter.addWidget(phases_group)

        events_group = QtWidgets.QGroupBox("事件(损失工时)")
        events_layout = QtWidgets.QVBoxLayout(events_group)
        events_toolbar = QtWidgets.QHBoxLayout()
        events_toolbar.addStretch(1)
        self.event_form_toggle_btn = QtWidgets.QPushButton("编辑事件")
        self.event_form_toggle_btn.clicked.connect(
            lambda: self._toggle_form_section(
                self.event_form_toggle_btn,
                self.event_form_widget,
                "事件",
            )
        )
        events_toolbar.addWidget(self.event_form_toggle_btn)
        events_layout.addLayout(events_toolbar)

        self.events_table = QtWidgets.QTableWidget(0, 3)
        self.events_table.setHorizontalHeaderLabels(["日期", "损失工时", "原因"])
        self.events_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.events_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.events_table.verticalHeader().setVisible(False)
        self.events_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )

        events_layout.addWidget(self.events_table)

        self.event_form_widget = QtWidgets.QWidget()
        ev_form = QtWidgets.QGridLayout(self.event_form_widget)
        self.event_date_edit = QtWidgets.QDateEdit()
        self.event_date_edit.setDate(QtCore.QDate.currentDate())
        self._setup_date_edit(self.event_date_edit)
        self.event_hours_spin = QtWidgets.QDoubleSpinBox()
        self.event_hours_spin.setRange(0, 24)
        self.event_hours_spin.setDecimals(2)
        self.event_hours_spin.setValue(8.0)
        self.event_reason_combo = QtWidgets.QComboBox()
        self.event_reason_combo.setEditable(False)
        self.event_reason_combo.addItems(self.event_reasons)

        ev_form.addWidget(QtWidgets.QLabel("日期"), 0, 0)
        ev_form.addWidget(self.event_date_edit, 0, 1)
        ev_form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        ev_form.addWidget(self.event_hours_spin, 0, 3)
        ev_form.addWidget(QtWidgets.QLabel("原因"), 0, 4)
        ev_form.addWidget(self.event_reason_combo, 0, 5)

        self.event_add_btn = QtWidgets.QPushButton("添加事件")
        self.event_add_btn.clicked.connect(self.add_event)
        self.event_remove_btn = QtWidgets.QPushButton("删除事件")
        self.event_remove_btn.clicked.connect(self.remove_event)

        ev_form.addWidget(self.event_add_btn, 1, 4)
        ev_form.addWidget(self.event_remove_btn, 1, 5)

        events_layout.addWidget(self.event_form_widget)
        self.event_form_widget.setVisible(False)
        right_splitter.addWidget(events_group)

        defects_group = QtWidgets.QGroupBox("不合格品")
        defects_layout = QtWidgets.QVBoxLayout(defects_group)
        defects_toolbar = QtWidgets.QHBoxLayout()
        defects_toolbar.addStretch(1)
        self.defect_form_toggle_btn = QtWidgets.QPushButton("编辑不合格")
        self.defect_form_toggle_btn.clicked.connect(
            lambda: self._toggle_form_section(
                self.defect_form_toggle_btn,
                self.defect_form_widget,
                "不合格",
            )
        )
        defects_toolbar.addWidget(self.defect_form_toggle_btn)
        defects_layout.addLayout(defects_toolbar)

        self.defects_table = QtWidgets.QTableWidget(0, 5)
        self.defects_table.setHorizontalHeaderLabels(["产品", "数量", "原因类别", "说明", "时间"])
        self.defects_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.defects_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.defects_table.verticalHeader().setVisible(False)
        self.defects_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.defects_table.itemSelectionChanged.connect(self.on_defect_select)
        defects_layout.addWidget(self.defects_table)

        self.defect_form_widget = QtWidgets.QWidget()
        defect_form = QtWidgets.QGridLayout(self.defect_form_widget)
        self.defect_product_combo = QtWidgets.QComboBox()
        self.defect_category_combo = QtWidgets.QComboBox()
        self.defect_category_combo.addItems(self.defect_categories)
        self.defect_category_combo.currentTextChanged.connect(self.on_defect_category_change)
        self.defect_count_spin = QtWidgets.QSpinBox()
        self.defect_count_spin.setRange(1, 999999)
        self.defect_detail_stack = QtWidgets.QStackedWidget()
        self.defect_detail_edit = QtWidgets.QLineEdit()
        self.defect_detail_edit.setPlaceholderText("原材料/原因说明")
        self.defect_detail_equipment_combo = QtWidgets.QComboBox()
        self.defect_detail_equipment_combo.setEditable(True)
        self.defect_detail_employee_combo = QtWidgets.QComboBox()
        self.defect_detail_employee_combo.setEditable(True)
        self.defect_detail_stack.addWidget(self.defect_detail_edit)
        self.defect_detail_stack.addWidget(self.defect_detail_equipment_combo)
        self.defect_detail_stack.addWidget(self.defect_detail_employee_combo)
        self.on_defect_category_change(self.defect_category_combo.currentText())

        defect_form.addWidget(QtWidgets.QLabel("产品"), 0, 0)
        defect_form.addWidget(self.defect_product_combo, 0, 1)
        defect_form.addWidget(QtWidgets.QLabel("原因"), 0, 2)
        defect_form.addWidget(self.defect_category_combo, 0, 3)
        defect_form.addWidget(QtWidgets.QLabel("数量"), 0, 4)
        defect_form.addWidget(self.defect_count_spin, 0, 5)
        defect_form.addWidget(QtWidgets.QLabel("说明"), 1, 0)
        defect_form.addWidget(self.defect_detail_stack, 1, 1, 1, 5)

        self.defect_add_btn = QtWidgets.QPushButton("添加不合格")
        self.defect_add_btn.clicked.connect(self.add_defect)
        self.defect_update_btn = QtWidgets.QPushButton("更新记录")
        self.defect_update_btn.clicked.connect(self.update_defect)
        self.defect_remove_btn = QtWidgets.QPushButton("删除记录")
        self.defect_remove_btn.clicked.connect(self.remove_defect)

        defect_form.addWidget(self.defect_add_btn, 2, 3)
        defect_form.addWidget(self.defect_update_btn, 2, 4)
        defect_form.addWidget(self.defect_remove_btn, 2, 5)

        defects_layout.addWidget(self.defect_form_widget)
        self.defect_form_widget.setVisible(False)
        right_splitter.addWidget(defects_group)

        right_layout.addWidget(right_splitter, 1)
        right_splitter.setSizes([560, 220, 240])

        splitter.addWidget(right_panel)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([640, 640])

        return page

    def _build_admin_page(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)

        top_bar = QtWidgets.QHBoxLayout()
        self.admin_back_btn = QtWidgets.QPushButton("返回主界面")
        self.admin_back_btn.clicked.connect(self.go_to_dashboard)
        self.admin_switch_user_btn = QtWidgets.QPushButton("切换用户")
        self.admin_switch_user_btn.clicked.connect(self.switch_user)
        self.admin_logout_btn = QtWidgets.QPushButton("退出登录")
        self.admin_logout_btn.clicked.connect(self.logout_user)
        self.admin_title_label = QtWidgets.QLabel("管理员界面")

        top_bar.addWidget(self.admin_back_btn)
        top_bar.addWidget(self.admin_switch_user_btn)
        top_bar.addWidget(self.admin_logout_btn)
        top_bar.addWidget(self.admin_title_label)
        top_bar.addStretch(1)
        layout.addLayout(top_bar)

        self.admin_tabs = QtWidgets.QTabWidget()
        self.admin_tabs.addTab(self._build_admin_events_tab(), "事件管理")
        self.admin_tabs.addTab(self._build_admin_log_tab(), "订单日志")
        self.admin_tabs.addTab(self._build_admin_users_tab(), "用户管理")
        self.admin_tabs.addTab(self._build_admin_reasons_tab(), "事件原因")
        self.admin_tabs.addTab(self._build_admin_defect_categories_tab(), "不合格原因")
        self.admin_tabs.addTab(self._build_admin_equipment_categories_tab(), "设备分类")
        self.admin_tabs.addTab(self._build_admin_equipment_templates_tab(), "设备模板")
        self.admin_tabs.addTab(self._build_admin_phase_templates_tab(), "工序模板")
        self.admin_tabs.addTab(self._build_admin_employee_templates_tab(), "员工模板")
        self.admin_tabs.addTab(self._build_admin_shift_templates_tab(), "班次模板")

        layout.addWidget(self.admin_tabs)
        return page

    def _build_visual_page(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        top_bar = QtWidgets.QHBoxLayout()
        self.visual_back_btn = QtWidgets.QPushButton("返回主界面")
        self.visual_back_btn.clicked.connect(self.go_to_dashboard)
        self.visual_detail_btn = QtWidgets.QPushButton("订单详情")
        self.visual_detail_btn.clicked.connect(self.go_to_detail)
        self.visual_switch_user_btn = QtWidgets.QPushButton("切换用户")
        self.visual_switch_user_btn.clicked.connect(self.switch_user)
        self.visual_logout_btn = QtWidgets.QPushButton("退出登录")
        self.visual_logout_btn.clicked.connect(self.logout_user)
        self.visual_title = QtWidgets.QLabel("数据看板")
        self.visual_title.setObjectName("heroTitle")

        top_bar.addWidget(self.visual_back_btn)
        top_bar.addWidget(self.visual_detail_btn)
        top_bar.addWidget(self.visual_switch_user_btn)
        top_bar.addWidget(self.visual_logout_btn)
        top_bar.addWidget(self.visual_title)
        top_bar.addStretch(1)
        layout.addLayout(top_bar)

        summary_group = QtWidgets.QGroupBox("关键指标")
        summary_layout = QtWidgets.QGridLayout(summary_group)
        self.visual_order_label = QtWidgets.QLabel("订单: -")
        self.visual_shift_label = QtWidgets.QLabel("班次: -")
        self.visual_counts_label = QtWidgets.QLabel("产品/设备/员工: -/-/-")
        self.visual_eta_label = QtWidgets.QLabel("预计交期: -")
        self.visual_total_hours_label = QtWidgets.QLabel("总工时: -")
        self.visual_done_hours_label = QtWidgets.QLabel("已完成工时: -")
        self.visual_remaining_hours_label = QtWidgets.QLabel("剩余工时: -")
        self.visual_lost_hours_label = QtWidgets.QLabel("累计损失: -")
        self.visual_defect_total_label = QtWidgets.QLabel("不合格数量: -")
        self.visual_equipment_avg_label = QtWidgets.QLabel("平均设备使用率: -")
        self.visual_equipment_peak_label = QtWidgets.QLabel("最高设备使用率: -")
        summary_labels = [
            self.visual_order_label,
            self.visual_shift_label,
            self.visual_counts_label,
            self.visual_eta_label,
            self.visual_total_hours_label,
            self.visual_done_hours_label,
            self.visual_remaining_hours_label,
            self.visual_lost_hours_label,
            self.visual_defect_total_label,
            self.visual_equipment_avg_label,
            self.visual_equipment_peak_label,
        ]
        for label in summary_labels:
            label.setObjectName("metricLabel")
        summary_layout.addWidget(self.visual_order_label, 0, 0)
        summary_layout.addWidget(self.visual_shift_label, 0, 1)
        summary_layout.addWidget(self.visual_counts_label, 0, 2)
        summary_layout.addWidget(self.visual_eta_label, 0, 3)
        summary_layout.addWidget(self.visual_total_hours_label, 1, 0)
        summary_layout.addWidget(self.visual_done_hours_label, 1, 1)
        summary_layout.addWidget(self.visual_remaining_hours_label, 1, 2)
        summary_layout.addWidget(self.visual_lost_hours_label, 1, 3)
        summary_layout.addWidget(self.visual_defect_total_label, 2, 0)
        summary_layout.addWidget(self.visual_equipment_avg_label, 2, 1)
        summary_layout.addWidget(self.visual_equipment_peak_label, 2, 2)
        for col in range(4):
            summary_layout.setColumnStretch(col, 1)
        layout.addWidget(summary_group)

        upper = QtWidgets.QHBoxLayout()
        layout.addLayout(upper)

        progress_group = QtWidgets.QGroupBox("订单进度")
        progress_layout = QtWidgets.QVBoxLayout(progress_group)
        self.visual_progress_bar = QtWidgets.QProgressBar()
        progress_layout.addWidget(self.visual_progress_bar)
        self.visual_progress_table = QtWidgets.QTableWidget(0, 3)
        self.visual_progress_table.setHorizontalHeaderLabels(["产品", "数量", "进度"])
        self.visual_progress_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_progress_table.verticalHeader().setVisible(False)
        self.visual_progress_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        progress_layout.addWidget(self.visual_progress_table)
        upper.addWidget(progress_group, 2)

        defect_group = QtWidgets.QGroupBox("零件不合格统计")
        defect_layout = QtWidgets.QVBoxLayout(defect_group)
        self.visual_defect_table = QtWidgets.QTableWidget(0, 2)
        self.visual_defect_table.setHorizontalHeaderLabels(["原因类别", "数量"])
        self.visual_defect_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_defect_table.verticalHeader().setVisible(False)
        self.visual_defect_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        defect_layout.addWidget(self.visual_defect_table)
        upper.addWidget(defect_group, 1)

        event_group = QtWidgets.QGroupBox("事件损失统计")
        event_layout = QtWidgets.QVBoxLayout(event_group)
        self.visual_event_table = QtWidgets.QTableWidget(0, 2)
        self.visual_event_table.setHorizontalHeaderLabels(["原因", "损失工时"])
        self.visual_event_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_event_table.verticalHeader().setVisible(False)
        self.visual_event_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        event_layout.addWidget(self.visual_event_table)
        upper.addWidget(event_group, 1)

        equipment_group = QtWidgets.QGroupBox("设备使用率")
        equipment_layout = QtWidgets.QVBoxLayout(equipment_group)
        self.visual_equipment_table = QtWidgets.QTableWidget(0, 5)
        self.visual_equipment_table.setHorizontalHeaderLabels(
            ["设备", "类别", "负载工时", "可用工时", "使用率"]
        )
        self.visual_equipment_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_equipment_table.verticalHeader().setVisible(False)
        self.visual_equipment_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        equipment_layout.addWidget(self.visual_equipment_table)
        layout.addWidget(equipment_group)

        return page

    def _build_admin_events_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_events_table = QtWidgets.QTableWidget(0, 3)
        self.admin_events_table.setHorizontalHeaderLabels(["日期", "损失工时", "原因"])
        self.admin_events_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.admin_events_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.admin_events_table.verticalHeader().setVisible(False)
        self.admin_events_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.admin_events_table.itemSelectionChanged.connect(self.on_admin_event_select)
        layout.addWidget(self.admin_events_table)

        form = QtWidgets.QGridLayout()
        self.admin_event_date_edit = QtWidgets.QDateEdit()
        self.admin_event_date_edit.setDate(QtCore.QDate.currentDate())
        self._setup_date_edit(self.admin_event_date_edit)
        self.admin_event_hours_spin = QtWidgets.QDoubleSpinBox()
        self.admin_event_hours_spin.setRange(0, 24)
        self.admin_event_hours_spin.setDecimals(2)
        self.admin_event_hours_spin.setValue(8.0)
        self.admin_event_reason_combo = QtWidgets.QComboBox()
        self.admin_event_reason_combo.setEditable(False)
        self.admin_event_reason_combo.addItems(self.event_reasons)

        form.addWidget(QtWidgets.QLabel("日期"), 0, 0)
        form.addWidget(self.admin_event_date_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        form.addWidget(self.admin_event_hours_spin, 0, 3)
        form.addWidget(QtWidgets.QLabel("原因"), 0, 4)
        form.addWidget(self.admin_event_reason_combo, 0, 5)

        self.admin_event_add_btn = QtWidgets.QPushButton("新增事件")
        self.admin_event_add_btn.clicked.connect(self.admin_add_event)
        self.admin_event_update_btn = QtWidgets.QPushButton("更新事件")
        self.admin_event_update_btn.clicked.connect(self.admin_update_event)
        self.admin_event_remove_btn = QtWidgets.QPushButton("删除事件")
        self.admin_event_remove_btn.clicked.connect(self.admin_remove_event)

        form.addWidget(self.admin_event_add_btn, 1, 3)
        form.addWidget(self.admin_event_update_btn, 1, 4)
        form.addWidget(self.admin_event_remove_btn, 1, 5)

        layout.addLayout(form)
        return page

    def _build_admin_log_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        tip = QtWidgets.QLabel("订单日志仅管理员可查看")
        tip.setObjectName("metricLabel")
        layout.addWidget(tip)

        self.admin_log_table = QtWidgets.QTableWidget(0, 3)
        self.admin_log_table.setHorizontalHeaderLabels(["时间", "用户", "内容"])
        self.admin_log_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.admin_log_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.admin_log_table.verticalHeader().setVisible(False)
        self.admin_log_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.admin_log_table)
        return page

    def _build_admin_users_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_user_list = QtWidgets.QListWidget()
        self.admin_user_list.itemSelectionChanged.connect(self.on_admin_user_select)
        layout.addWidget(self.admin_user_list)

        form = QtWidgets.QGridLayout()
        self.admin_user_name_edit = QtWidgets.QLineEdit()
        self.admin_user_name_edit.setPlaceholderText("用户名")
        self.admin_user_pass_edit = QtWidgets.QLineEdit()
        self.admin_user_pass_edit.setPlaceholderText("密码")
        self.admin_user_pass_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.admin_show_password_check = QtWidgets.QCheckBox("显示密码")
        self.admin_show_password_check.toggled.connect(self._toggle_admin_password_visibility)

        self.admin_user_add_btn = QtWidgets.QPushButton("添加用户")
        self.admin_user_add_btn.clicked.connect(self.admin_add_user)
        self.admin_user_update_btn = QtWidgets.QPushButton("更新用户")
        self.admin_user_update_btn.clicked.connect(self.admin_update_user)
        self.admin_user_remove_btn = QtWidgets.QPushButton("删除用户")
        self.admin_user_remove_btn.clicked.connect(self.admin_remove_user)

        form.addWidget(QtWidgets.QLabel("用户名"), 0, 0)
        form.addWidget(self.admin_user_name_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("密码"), 0, 2)
        form.addWidget(self.admin_user_pass_edit, 0, 3)
        form.addWidget(self.admin_show_password_check, 1, 0)
        form.addWidget(self.admin_user_add_btn, 1, 1)
        form.addWidget(self.admin_user_update_btn, 1, 2)
        form.addWidget(self.admin_user_remove_btn, 1, 3)

        layout.addLayout(form)
        return page

    def _build_admin_reasons_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_reason_list = QtWidgets.QListWidget()
        self.admin_reason_list.itemSelectionChanged.connect(self.on_admin_reason_select)
        layout.addWidget(self.admin_reason_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_reason_edit = QtWidgets.QLineEdit()
        self.admin_reason_edit.setPlaceholderText("事件原因")
        self.admin_reason_add_btn = QtWidgets.QPushButton("添加")
        self.admin_reason_add_btn.clicked.connect(self.admin_add_reason)
        self.admin_reason_update_btn = QtWidgets.QPushButton("更新选中")
        self.admin_reason_update_btn.clicked.connect(self.admin_update_reason)
        self.admin_reason_remove_btn = QtWidgets.QPushButton("删除选中")
        self.admin_reason_remove_btn.clicked.connect(self.admin_remove_reason)

        form.addWidget(self.admin_reason_edit)
        form.addWidget(self.admin_reason_add_btn)
        form.addWidget(self.admin_reason_update_btn)
        form.addWidget(self.admin_reason_remove_btn)

        layout.addLayout(form)
        return page

    def _build_admin_defect_categories_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_defect_category_list = QtWidgets.QListWidget()
        self.admin_defect_category_list.itemSelectionChanged.connect(
            self.on_admin_defect_category_select
        )
        layout.addWidget(self.admin_defect_category_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_defect_category_edit = QtWidgets.QLineEdit()
        self.admin_defect_category_edit.setPlaceholderText("不合格原因类别")
        self.admin_defect_category_add_btn = QtWidgets.QPushButton("添加")
        self.admin_defect_category_add_btn.clicked.connect(self.admin_add_defect_category)
        self.admin_defect_category_update_btn = QtWidgets.QPushButton("更新选中")
        self.admin_defect_category_update_btn.clicked.connect(self.admin_update_defect_category)
        self.admin_defect_category_remove_btn = QtWidgets.QPushButton("删除选中")
        self.admin_defect_category_remove_btn.clicked.connect(self.admin_remove_defect_category)

        form.addWidget(self.admin_defect_category_edit)
        form.addWidget(self.admin_defect_category_add_btn)
        form.addWidget(self.admin_defect_category_update_btn)
        form.addWidget(self.admin_defect_category_remove_btn)

        layout.addLayout(form)
        return page

    def _build_admin_equipment_categories_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_equipment_category_list = QtWidgets.QListWidget()
        self.admin_equipment_category_list.itemSelectionChanged.connect(
            self.on_admin_equipment_category_select
        )
        layout.addWidget(self.admin_equipment_category_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_equipment_category_edit = QtWidgets.QLineEdit()
        self.admin_equipment_category_edit.setPlaceholderText("设备分类名称")
        self.admin_equipment_category_add_btn = QtWidgets.QPushButton("添加")
        self.admin_equipment_category_add_btn.clicked.connect(self.admin_add_equipment_category)
        self.admin_equipment_category_update_btn = QtWidgets.QPushButton("更新选中")
        self.admin_equipment_category_update_btn.clicked.connect(self.admin_update_equipment_category)
        self.admin_equipment_category_remove_btn = QtWidgets.QPushButton("删除选中")
        self.admin_equipment_category_remove_btn.clicked.connect(self.admin_remove_equipment_category)

        form.addWidget(self.admin_equipment_category_edit)
        form.addWidget(self.admin_equipment_category_add_btn)
        form.addWidget(self.admin_equipment_category_update_btn)
        form.addWidget(self.admin_equipment_category_remove_btn)

        layout.addLayout(form)
        return page

    def _build_admin_equipment_templates_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_equipment_template_table = QtWidgets.QTableWidget(0, 4)
        self.admin_equipment_template_table.setHorizontalHeaderLabels(["设备编号", "类别", "总数量", "可用数量"])
        self.admin_equipment_template_table.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.admin_equipment_template_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.admin_equipment_template_table.verticalHeader().setVisible(False)
        self.admin_equipment_template_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.admin_equipment_template_table.itemSelectionChanged.connect(
            self.on_admin_equipment_template_select
        )
        layout.addWidget(self.admin_equipment_template_table)

        form = QtWidgets.QGridLayout()
        form.setColumnStretch(3, 2)
        self.admin_equipment_id_edit = QtWidgets.QLineEdit()
        self.admin_equipment_category_combo = QtWidgets.QComboBox()
        self.admin_equipment_category_combo.setEditable(False)
        self.admin_equipment_category_combo.addItems(self.equipment_categories)
        self._configure_combo_popup(self.admin_equipment_category_combo, 200)
        self.admin_equipment_total_spin = QtWidgets.QSpinBox()
        self.admin_equipment_total_spin.setRange(1, 9999)
        self.admin_equipment_available_spin = QtWidgets.QSpinBox()
        self.admin_equipment_available_spin.setRange(0, 9999)

        form.addWidget(QtWidgets.QLabel("设备编号"), 0, 0)
        form.addWidget(self.admin_equipment_id_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("类别"), 0, 2)
        form.addWidget(self.admin_equipment_category_combo, 0, 3)
        form.addWidget(QtWidgets.QLabel("总数量"), 0, 4)
        form.addWidget(self.admin_equipment_total_spin, 0, 5)
        form.addWidget(QtWidgets.QLabel("可用数量"), 0, 6)
        form.addWidget(self.admin_equipment_available_spin, 0, 7)

        self.admin_equipment_add_btn = QtWidgets.QPushButton("添加/更新模板")
        self.admin_equipment_add_btn.clicked.connect(self.admin_add_or_update_equipment_template)
        self.admin_equipment_remove_btn = QtWidgets.QPushButton("删除模板")
        self.admin_equipment_remove_btn.clicked.connect(self.admin_remove_equipment_template)
        self.admin_equipment_apply_btn = QtWidgets.QPushButton("应用到当前订单")
        self.admin_equipment_apply_btn.clicked.connect(self.admin_apply_equipment_template)

        form.addWidget(self.admin_equipment_add_btn, 1, 5)
        form.addWidget(self.admin_equipment_remove_btn, 1, 6)
        form.addWidget(self.admin_equipment_apply_btn, 1, 7)

        layout.addLayout(form)
        return page

    def _build_admin_phase_templates_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        apply_row = QtWidgets.QHBoxLayout()
        self.admin_phase_product_combo = QtWidgets.QComboBox()
        self.admin_phase_apply_btn = QtWidgets.QPushButton("应用到产品")
        self.admin_phase_apply_btn.clicked.connect(self.admin_apply_phase_template)
        apply_row.addWidget(QtWidgets.QLabel("应用到产品"))
        apply_row.addWidget(self.admin_phase_product_combo)
        apply_row.addStretch(1)
        apply_row.addWidget(self.admin_phase_apply_btn)
        layout.addLayout(apply_row)

        self.admin_phase_template_table = QtWidgets.QTableWidget(0, 5)
        self.admin_phase_template_table.setHorizontalHeaderLabels(
            ["工序名称", "工时", "设备", "员工", "并行组"]
        )
        self.admin_phase_template_table.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.admin_phase_template_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.admin_phase_template_table.verticalHeader().setVisible(False)
        self.admin_phase_template_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.admin_phase_template_table.itemSelectionChanged.connect(
            self.on_admin_phase_template_select
        )
        layout.addWidget(self.admin_phase_template_table)

        form = QtWidgets.QGridLayout()
        self.admin_phase_name_edit = QtWidgets.QLineEdit()
        self.admin_phase_hours_spin = QtWidgets.QDoubleSpinBox()
        self.admin_phase_hours_spin.setRange(0, 99999)
        self.admin_phase_hours_spin.setDecimals(2)
        self.admin_phase_equipment_combo = QtWidgets.QComboBox()
        self.admin_phase_equipment_combo.setEditable(True)
        self.admin_phase_employee_combo = QtWidgets.QComboBox()
        self.admin_phase_employee_combo.setEditable(True)
        self.admin_phase_parallel_spin = QtWidgets.QSpinBox()
        self.admin_phase_parallel_spin.setRange(0, 9999)

        form.addWidget(QtWidgets.QLabel("名称"), 0, 0)
        form.addWidget(self.admin_phase_name_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        form.addWidget(self.admin_phase_hours_spin, 0, 3)
        form.addWidget(QtWidgets.QLabel("设备"), 0, 4)
        form.addWidget(self.admin_phase_equipment_combo, 0, 5)

        form.addWidget(QtWidgets.QLabel("员工"), 1, 0)
        form.addWidget(self.admin_phase_employee_combo, 1, 1)
        form.addWidget(QtWidgets.QLabel("并行组"), 1, 2)
        form.addWidget(self.admin_phase_parallel_spin, 1, 3)

        self.admin_phase_add_btn = QtWidgets.QPushButton("添加/更新模板")
        self.admin_phase_add_btn.clicked.connect(self.admin_add_or_update_phase_template)
        self.admin_phase_remove_btn = QtWidgets.QPushButton("删除模板")
        self.admin_phase_remove_btn.clicked.connect(self.admin_remove_phase_template)

        form.addWidget(self.admin_phase_add_btn, 2, 3)
        form.addWidget(self.admin_phase_remove_btn, 2, 4)

        layout.addLayout(form)
        return page

    def _build_admin_employee_templates_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_employee_template_list = QtWidgets.QListWidget()
        self.admin_employee_template_list.itemSelectionChanged.connect(
            self.on_admin_employee_template_select
        )
        layout.addWidget(self.admin_employee_template_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_employee_name_edit = QtWidgets.QLineEdit()
        self.admin_employee_name_edit.setPlaceholderText("员工姓名")
        self.admin_employee_add_btn = QtWidgets.QPushButton("添加/更新模板")
        self.admin_employee_add_btn.clicked.connect(self.admin_add_or_update_employee_template)
        self.admin_employee_remove_btn = QtWidgets.QPushButton("删除模板")
        self.admin_employee_remove_btn.clicked.connect(self.admin_remove_employee_template)
        self.admin_employee_apply_btn = QtWidgets.QPushButton("应用到当前订单")
        self.admin_employee_apply_btn.clicked.connect(self.admin_apply_employee_template)

        form.addWidget(self.admin_employee_name_edit)
        form.addWidget(self.admin_employee_add_btn)
        form.addWidget(self.admin_employee_remove_btn)
        form.addWidget(self.admin_employee_apply_btn)

        layout.addLayout(form)
        return page

    def _build_admin_shift_templates_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(page)

        left_panel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_panel)

        self.admin_shift_template_list = QtWidgets.QListWidget()
        self.admin_shift_template_list.itemSelectionChanged.connect(self.on_admin_shift_template_select)
        left_layout.addWidget(self.admin_shift_template_list)

        name_row = QtWidgets.QHBoxLayout()
        self.admin_shift_name_edit = QtWidgets.QLineEdit()
        self.admin_shift_name_edit.setPlaceholderText("班次模板名称")
        name_row.addWidget(self.admin_shift_name_edit)
        left_layout.addLayout(name_row)

        button_row = QtWidgets.QHBoxLayout()
        self.admin_shift_save_btn = QtWidgets.QPushButton("添加/更新模板")
        self.admin_shift_save_btn.clicked.connect(self.admin_add_or_update_shift_template)
        self.admin_shift_delete_btn = QtWidgets.QPushButton("删除模板")
        self.admin_shift_delete_btn.clicked.connect(self.admin_remove_shift_template)
        self.admin_shift_activate_btn = QtWidgets.QPushButton("设为当前班次")
        self.admin_shift_activate_btn.clicked.connect(self.admin_set_active_shift_template)
        button_row.addWidget(self.admin_shift_save_btn)
        button_row.addWidget(self.admin_shift_delete_btn)
        button_row.addWidget(self.admin_shift_activate_btn)
        left_layout.addLayout(button_row)

        layout.addWidget(left_panel, 1)

        right_panel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_panel)

        self.admin_shift_active_label = QtWidgets.QLabel("当前班次: -")
        right_layout.addWidget(self.admin_shift_active_label)

        self.admin_shift_table = QtWidgets.QTableWidget(7, 4)
        self.admin_shift_table.setHorizontalHeaderLabels(["星期", "班次数", "每班小时", "当日总工时"])
        self.admin_shift_table.verticalHeader().setVisible(False)
        self.admin_shift_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.admin_shift_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        right_layout.addWidget(self.admin_shift_table)

        self.shift_count_spins: List[QtWidgets.QSpinBox] = []
        self.shift_hours_spins: List[QtWidgets.QDoubleSpinBox] = []
        days = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for row, day in enumerate(days):
            day_item = QtWidgets.QTableWidgetItem(day)
            day_item.setFlags(QtCore.Qt.ItemFlag.ItemIsEnabled)
            self.admin_shift_table.setItem(row, 0, day_item)

            count_spin = QtWidgets.QSpinBox()
            count_spin.setRange(0, 10)
            hours_spin = QtWidgets.QDoubleSpinBox()
            hours_spin.setRange(0, 24)
            hours_spin.setDecimals(1)
            hours_spin.setSingleStep(0.5)

            count_spin.valueChanged.connect(lambda _, r=row: self._update_shift_row_total(r))
            hours_spin.valueChanged.connect(lambda _, r=row: self._update_shift_row_total(r))

            self.admin_shift_table.setCellWidget(row, 1, count_spin)
            self.admin_shift_table.setCellWidget(row, 2, hours_spin)
            total_item = QtWidgets.QTableWidgetItem("0")
            total_item.setFlags(QtCore.Qt.ItemFlag.ItemIsEnabled)
            self.admin_shift_table.setItem(row, 3, total_item)

            self.shift_count_spins.append(count_spin)
            self.shift_hours_spins.append(hours_spin)

        layout.addWidget(right_panel, 2)
        return page

    # ------------------------
    # Order lifecycle
    # ------------------------

    def create_order(self) -> None:
        order_id = self.order_id_edit.text().strip() or "O-UNKNOWN"
        start_dt = datetime.combine(_qdate_to_date(self.start_date_edit.date()), datetime.min.time())
        self.autosave_path = None
        if not self._ensure_autosave_path(f"{order_id}.json"):
            QtWidgets.QMessageBox.information(self, "提示", "需要选择保存位置以启用自动保存。")
            return
        self.order = Order(
            order_id=order_id,
            start_dt=start_dt,
            products=[],
            events=[],
            defects=[],
            equipment=list(self.equipment),
            employees=list(self.employees),
        )
        self._log_change(f"创建订单 {order_id} (开始日期: {start_dt.date().isoformat()})")
        self._refresh_all()
        self._auto_save()
        self.go_to_detail()

    def save_order(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先创建或加载订单。")
            return
        default_name = f"{self.order.order_id}.json"
        filename, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "保存订单", default_name, "JSON Files (*.json)"
        )
        if not filename:
            return
        if not filename.endswith(".json"):
            filename += ".json"
        self._set_autosave_path(filename)
        self._save_to_path(filename, show_message=True)

    def load_order(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "加载订单", "", "JSON Files (*.json)"
        )
        if not filename:
            return
        try:
            with open(filename, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.order = self._order_from_dict(data)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "加载失败", f"无法加载订单: {exc}")
            return

        self.order_id_edit.setText(self.order.order_id)
        self.start_date_edit.setDate(_date_to_qdate(self.order.start_dt.date()))

        self.equipment = list(self.order.equipment)
        self.employees = list(self.order.employees)
        self._maybe_import_templates_from_order(data)
        self._set_autosave_path(filename)

        self._refresh_all()
        self.go_to_detail()

    def go_to_detail(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        self.stack.setCurrentWidget(self.detail_page)

    def go_to_visuals(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        self._refresh_visuals()
        self.stack.setCurrentWidget(self.visual_page)

    def open_admin_login(self) -> None:
        password, ok = QtWidgets.QInputDialog.getText(
            self, "管理员登录", "请输入管理员密码", QtWidgets.QLineEdit.EchoMode.Password
        )
        if not ok:
            return
        admin_account = next((u for u in self.user_accounts if u.username == "admin"), None)
        admin_password = admin_account.password if admin_account else ADMIN_PASSWORD
        if password != admin_password:
            QtWidgets.QMessageBox.warning(self, "验证失败", "管理员密码不正确。")
            return
        self._refresh_admin_views()
        self.stack.setCurrentWidget(self.admin_page)

    def go_to_dashboard(self) -> None:
        if not self.current_user:
            self.stack.setCurrentWidget(self.login_page)
            return
        self.stack.setCurrentWidget(self.dashboard)

    def logout_user(self) -> None:
        if not self.current_user:
            self.stack.setCurrentWidget(self.login_page)
            return
        self.current_user = ""
        self.login_user_edit.clear()
        self.login_pass_edit.clear()
        self.login_status_label.setText("已退出登录。")
        self.stack.setCurrentWidget(self.login_page)

    def switch_user(self) -> None:
        if not self.current_user:
            self.stack.setCurrentWidget(self.login_page)
            return
        self.current_user = ""
        self.login_user_edit.clear()
        self.login_pass_edit.clear()
        self.login_status_label.setText("请使用新用户登录。")
        self.stack.setCurrentWidget(self.login_page)

    def _log_change(self, content: str) -> None:
        if not self.order:
            return
        user = self.current_user or "系统"
        self.order.logs.append(LogEntry(timestamp=datetime.now(), user=user, content=content))
        self._refresh_admin_log_table()

    def _ensure_autosave_path(self, default_name: str) -> bool:
        if not default_name:
            return False
        filename = default_name
        if os.path.exists(filename):
            filename, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "保存订单", default_name, "JSON Files (*.json)"
            )
            if not filename:
                return False
        if not filename.endswith(".json"):
            filename += ".json"
        self._set_autosave_path(filename)
        return True

    def _set_autosave_path(self, filename: str) -> None:
        self.autosave_path = filename
        self.statusBar().showMessage(f"自动保存文件: {filename}", 5000)

    def _save_to_path(self, filename: str, show_message: bool = False, autosave: bool = False) -> None:
        if not self.order:
            return
        data = self._order_to_dict(self.order)
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        if show_message:
            QtWidgets.QMessageBox.information(self, "保存成功", f"订单已保存到 {filename}")
        if autosave:
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.statusBar().showMessage(f"自动保存: {os.path.basename(filename)} {timestamp}", 3000)

    def _auto_save(self) -> None:
        if not self.order or not self.autosave_path:
            return
        try:
            self._save_to_path(self.autosave_path, autosave=True)
        except Exception as exc:
            self.statusBar().showMessage(f"自动保存失败: {exc}", 5000)

    def _load_app_templates(self) -> None:
        if not os.path.exists(self.app_template_path):
            return
        try:
            with open(self.app_template_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            self.statusBar().showMessage("模板文件读取失败，已使用默认模板。", 5000)
            return

        self.event_reasons = list(data.get("event_reasons", self.event_reasons))
        self.equipment_categories = list(data.get("equipment_categories", self.equipment_categories))
        self.defect_categories = list(data.get("defect_categories", self.defect_categories))
        self.user_accounts = [
            UserAccount(username=u.get("username", ""), password=u.get("password", ""))
            for u in data.get("users", [])
            if u.get("username")
        ]
        templates = data.get("templates", {})
        self.equipment_templates = [
            Equipment(
                equipment_id=e.get("equipment_id", ""),
                category=e.get("category", ""),
                total_count=int(e.get("total_count", 1)),
                available_count=int(e.get("available_count", 1)),
            )
            for e in templates.get("equipment", [])
        ]
        self.phase_templates = [
            Phase(
                name=ph.get("name", ""),
                planned_hours=float(ph.get("planned_hours", 0)),
                parallel_group=int(ph.get("parallel_group", 0)),
                equipment_id=ph.get("equipment_id", ""),
                assigned_employee=ph.get("assigned_employee", ""),
            )
            for ph in templates.get("phases", [])
        ]
        self.employee_templates = list(templates.get("employees", []))
        self.shift_templates = []
        for tpl in templates.get("shifts", []):
            week_plan = [
                ShiftDayPlan(
                    shift_count=int(day.get("shift_count", 0)),
                    hours_per_shift=float(day.get("hours_per_shift", 0.0)),
                )
                for day in tpl.get("week_plan", [])
            ]
            if len(week_plan) < 7:
                week_plan.extend([ShiftDayPlan(0, 0.0) for _ in range(7 - len(week_plan))])
            self.shift_templates.append(ShiftTemplate(name=tpl.get("name", "班次模板"), week_plan=week_plan))
        self.active_shift_template_name = templates.get("active_shift", "")

    def _save_app_templates(self) -> None:
        data = {
            "version": 1,
            "event_reasons": self.event_reasons,
            "equipment_categories": self.equipment_categories,
            "defect_categories": self.defect_categories,
            "users": [
                {"username": u.username, "password": u.password} for u in self.user_accounts
            ],
            "templates": {
                "equipment": [
                    {
                        "equipment_id": e.equipment_id,
                        "category": e.category,
                        "total_count": e.total_count,
                        "available_count": e.available_count,
                    }
                    for e in self.equipment_templates
                ],
                "phases": [
                    {
                        "name": ph.name,
                        "planned_hours": ph.planned_hours,
                        "parallel_group": ph.parallel_group,
                        "equipment_id": ph.equipment_id,
                        "assigned_employee": ph.assigned_employee,
                    }
                    for ph in self.phase_templates
                ],
                "employees": list(self.employee_templates),
                "shifts": [
                    {
                        "name": tpl.name,
                        "week_plan": [
                            {
                                "shift_count": day.shift_count,
                                "hours_per_shift": day.hours_per_shift,
                            }
                            for day in tpl.week_plan
                        ],
                    }
                    for tpl in self.shift_templates
                ],
                "active_shift": self.active_shift_template_name,
            },
        }
        try:
            with open(self.app_template_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as exc:
            self.statusBar().showMessage(f"模板保存失败: {exc}", 5000)

    def _ensure_default_templates(self) -> None:
        if not self.event_reasons:
            self.event_reasons = ["员工请假", "设备故障", "停电", "材料短缺", "质量问题", "其他"]
        if not self.equipment_categories:
            self.equipment_categories = ["车床", "加工中心", "检验设备", "辅助设备"]
        if not self.defect_categories:
            self.defect_categories = ["设备", "原材料", "员工"]
        if not any(u.username == "admin" for u in self.user_accounts):
            self.user_accounts.append(UserAccount(username="admin", password=ADMIN_PASSWORD))
        if not self.shift_templates:
            week_plan = [
                ShiftDayPlan(1, 8.0),
                ShiftDayPlan(1, 8.0),
                ShiftDayPlan(1, 8.0),
                ShiftDayPlan(1, 8.0),
                ShiftDayPlan(1, 8.0),
                ShiftDayPlan(2, 8.0),
                ShiftDayPlan(0, 0.0),
            ]
            self.shift_templates.append(ShiftTemplate("周六两班", week_plan))
            self.active_shift_template_name = "周六两班"
        if self.active_shift_template_name and not self._current_shift_template():
            self.active_shift_template_name = self.shift_templates[0].name
        if not self.active_shift_template_name and self.shift_templates:
            self.active_shift_template_name = self.shift_templates[0].name
        self._save_app_templates()

    def _current_shift_template(self) -> Optional[ShiftTemplate]:
        for tpl in self.shift_templates:
            if tpl.name == self.active_shift_template_name:
                return tpl
        return self.shift_templates[0] if self.shift_templates else None

    def _apply_active_shift_template(self) -> None:
        self.cal.shift_template = self._current_shift_template()
        if hasattr(self, "admin_shift_active_label"):
            self._update_active_shift_label()
        self._refresh_dashboard_summary()
        self.refresh_eta()

    def _total_shift_hours(self, start_day: date, end_day: date) -> float:
        if not self.cal.shift_template or end_day < start_day:
            return 0.0
        total = 0.0
        current = start_day
        while current <= end_day:
            total += self.cal.capacity_for_day(current)
            current = current + timedelta(days=1)
        return total

    def _refresh_visuals(self) -> None:
        if not self.order:
            self.visual_progress_bar.setValue(0)
            self.visual_progress_table.setRowCount(0)
            self.visual_defect_table.setRowCount(0)
            self.visual_event_table.setRowCount(0)
            self.visual_equipment_table.setRowCount(0)
            self.visual_order_label.setText("订单: -")
            self.visual_shift_label.setText("班次: -")
            self.visual_counts_label.setText("产品/设备/员工: -/-/-")
            self.visual_eta_label.setText("预计交期: -")
            self.visual_total_hours_label.setText("总工时: -")
            self.visual_done_hours_label.setText("已完成工时: -")
            self.visual_remaining_hours_label.setText("剩余工时: -")
            self.visual_lost_hours_label.setText("累计损失: -")
            self.visual_defect_total_label.setText("不合格数量: -")
            self.visual_equipment_avg_label.setText("平均设备使用率: -")
            self.visual_equipment_peak_label.setText("最高设备使用率: -")
            return

        equipment_map = _equipment_available_map(self.order)
        total = 0.0
        done = 0.0
        for product in self.order.products:
            for phase in product.phases:
                hours = _phase_effective_hours(phase, product.quantity, equipment_map)
                total += hours
                ratio = _phase_completion_ratio(phase, product.quantity)
                done += hours * ratio
        progress = (done / total) if total > 0 else 0.0
        self.visual_progress_bar.setValue(int(progress * 100))
        remaining = max(total - done, 0.0)

        self.visual_order_label.setText(f"订单: {self.order.order_id}")
        self.visual_shift_label.setText(f"班次: {self.active_shift_template_name or '-'}")
        self.visual_counts_label.setText(
            f"产品/设备/员工: {len(self.order.products)}/{len(self.order.equipment)}/{len(self.order.employees)}"
        )
        self.visual_total_hours_label.setText(f"总工时: {total:g}h")
        self.visual_done_hours_label.setText(f"已完成工时: {done:g}h")
        self.visual_remaining_hours_label.setText(f"剩余工时: {remaining:g}h")

        self.visual_progress_table.setRowCount(0)
        for product in self.order.products:
            row = self.visual_progress_table.rowCount()
            self.visual_progress_table.insertRow(row)
            p_progress = _product_progress(product, equipment_map)
            self.visual_progress_table.setItem(row, 0, QtWidgets.QTableWidgetItem(product.product_id))
            self.visual_progress_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(product.quantity)))
            self.visual_progress_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{p_progress:.0%}"))

        defect_counts: Dict[str, int] = {c: 0 for c in self.defect_categories}
        for defect in self.order.defects:
            category = defect.category or "未分类"
            defect_counts.setdefault(category, 0)
            defect_counts[category] += defect.count
        self.visual_defect_table.setRowCount(0)
        for category, count in defect_counts.items():
            row = self.visual_defect_table.rowCount()
            self.visual_defect_table.insertRow(row)
            self.visual_defect_table.setItem(row, 0, QtWidgets.QTableWidgetItem(category))
            self.visual_defect_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(count)))
        total_defects = sum(defect.count for defect in self.order.defects)
        self.visual_defect_total_label.setText(f"不合格数量: {total_defects}")

        reason_hours: Dict[str, float] = {}
        for ev in self.order.events:
            reason = ev.reason or "事件"
            reason_hours[reason] = reason_hours.get(reason, 0.0) + ev.hours_lost
        self.visual_event_table.setRowCount(0)
        for reason, hours in sorted(reason_hours.items(), key=lambda x: x[1], reverse=True):
            row = self.visual_event_table.rowCount()
            self.visual_event_table.insertRow(row)
            self.visual_event_table.setItem(row, 0, QtWidgets.QTableWidgetItem(reason))
            self.visual_event_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{hours:g}"))
        total_lost = sum(reason_hours.values())
        self.visual_lost_hours_label.setText(f"累计损失: {total_lost:g}h")

        if not self.last_eta_dt:
            try:
                result = compute_eta(self.order, self.cal)
                self.last_eta_dt = result["eta_dt"]
                self.last_remaining_hours = result["remaining_hours"]
            except Exception:
                self.last_eta_dt = None

        total_shift_hours = 0.0
        if self.last_eta_dt:
            total_shift_hours = self._total_shift_hours(
                self.order.start_dt.date(), self.last_eta_dt.date()
            )
            self.visual_eta_label.setText(
                f"预计交期: {self.last_eta_dt.strftime('%Y-%m-%d %H:%M')}"
            )
        else:
            self.visual_eta_label.setText("预计交期: -")

        equipment_loads: Dict[str, float] = {}
        for product in self.order.products:
            for phase in product.phases:
                if not phase.equipment_id:
                    continue
                hours = phase.planned_hours * max(product.quantity, 1)
                equipment_loads[phase.equipment_id] = equipment_loads.get(phase.equipment_id, 0.0) + hours

        self.visual_equipment_table.setRowCount(0)
        usage_values: List[float] = []
        peak_usage_raw = -1.0
        peak_usage_capped = 0.0
        peak_equipment = "-"
        for eq in self.order.equipment:
            load = equipment_loads.get(eq.equipment_id, 0.0)
            capacity = total_shift_hours * max(eq.available_count, 1)
            usage_raw = (load / capacity) if capacity > 0 else 0.0
            usage_capped = min(max(usage_raw, 0.0), 1.0)
            usage_values.append(usage_capped)
            if usage_raw > peak_usage_raw:
                peak_usage_raw = usage_raw
                peak_usage_capped = usage_capped
                peak_equipment = eq.equipment_id or "-"
            row = self.visual_equipment_table.rowCount()
            self.visual_equipment_table.insertRow(row)
            self.visual_equipment_table.setItem(row, 0, QtWidgets.QTableWidgetItem(eq.equipment_id))
            self.visual_equipment_table.setItem(row, 1, QtWidgets.QTableWidgetItem(eq.category or "-"))
            self.visual_equipment_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{load:g}"))
            self.visual_equipment_table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{capacity:g}"))
            bar = QtWidgets.QProgressBar()
            usage_pct = int(round(usage_capped * 100))
            bar.setRange(0, 100)
            bar.setValue(usage_pct)
            bar.setFormat(f"{usage_pct}%")
            bar.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            bar.setProperty("complete", usage_pct >= 100)
            bar.style().unpolish(bar)
            bar.style().polish(bar)
            self.visual_equipment_table.setCellWidget(row, 4, bar)

        if usage_values:
            avg_usage = sum(usage_values) / len(usage_values)
            self.visual_equipment_avg_label.setText(f"平均设备使用率: {avg_usage:.0%}")
            over_flag = " (超负荷)" if peak_usage_raw > 1.0 else ""
            self.visual_equipment_peak_label.setText(
                f"最高设备使用率: {peak_equipment} {peak_usage_capped:.0%}{over_flag}"
            )
        else:
            self.visual_equipment_avg_label.setText("平均设备使用率: -")
            self.visual_equipment_peak_label.setText("最高设备使用率: -")

    def _maybe_import_templates_from_order(self, data: Dict[str, object]) -> None:
        if self.equipment_templates or self.phase_templates or self.employee_templates or self.shift_templates:
            return
        if "templates" not in data and "event_reasons" not in data:
            return
        self.event_reasons = list(data.get("event_reasons", self.event_reasons))
        self.equipment_categories = list(data.get("equipment_categories", self.equipment_categories))
        self.defect_categories = list(data.get("defect_categories", self.defect_categories))
        templates = data.get("templates", {})
        self.equipment_templates = [
            Equipment(
                equipment_id=e.get("equipment_id", ""),
                category=e.get("category", ""),
                total_count=int(e.get("total_count", 1)),
                available_count=int(e.get("available_count", 1)),
            )
            for e in templates.get("equipment", [])
        ]
        self.phase_templates = [
            Phase(
                name=ph.get("name", ""),
                planned_hours=float(ph.get("planned_hours", 0)),
                parallel_group=int(ph.get("parallel_group", 0)),
                equipment_id=ph.get("equipment_id", ""),
                assigned_employee=ph.get("assigned_employee", ""),
            )
            for ph in templates.get("phases", [])
        ]
        self.employee_templates = list(templates.get("employees", []))
        self.shift_templates = []
        for tpl in templates.get("shifts", []):
            week_plan = [
                ShiftDayPlan(
                    shift_count=int(day.get("shift_count", 0)),
                    hours_per_shift=float(day.get("hours_per_shift", 0.0)),
                )
                for day in tpl.get("week_plan", [])
            ]
            if len(week_plan) < 7:
                week_plan.extend([ShiftDayPlan(0, 0.0) for _ in range(7 - len(week_plan))])
            self.shift_templates.append(ShiftTemplate(name=tpl.get("name", "班次模板"), week_plan=week_plan))
        self.active_shift_template_name = templates.get("active_shift", "")
        self._ensure_default_templates()
        self.cal.shift_template = self._current_shift_template()

    def _refresh_event_reason_combo(self) -> None:
        current = self.event_reason_combo.currentText()
        self.event_reason_combo.clear()
        self.event_reason_combo.addItems(self.event_reasons)
        if current:
            self.event_reason_combo.setCurrentText(current)
        if hasattr(self, "admin_event_reason_combo"):
            admin_current = self.admin_event_reason_combo.currentText()
            self.admin_event_reason_combo.clear()
            self.admin_event_reason_combo.addItems(self.event_reasons)
            if admin_current:
                self.admin_event_reason_combo.setCurrentText(admin_current)

    def _refresh_admin_views(self) -> None:
        self._refresh_admin_events_table()
        self._refresh_admin_log_table()
        self._refresh_admin_reason_list()
        self._refresh_admin_defect_categories()
        self._refresh_admin_equipment_categories()
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_templates_table()
        self._refresh_admin_employee_templates_list()
        self._refresh_admin_shift_templates_list()
        self._refresh_admin_users_list()
        self._refresh_admin_product_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_event_reason_combo()
        self._refresh_equipment_category_combos()

    # ------------------------
    # Equipment management
    # ------------------------

    def on_equipment_select(self) -> None:
        row = self.equipment_table.currentRow()
        if row < 0 or row >= len(self.equipment):
            return
        eq = self.equipment[row]
        self.equipment_id_edit.setText(eq.equipment_id)
        self.equipment_category_combo.setCurrentText(eq.category)
        self.equipment_total_spin.setValue(eq.total_count)
        self.equipment_available_spin.setValue(eq.available_count)

    def add_or_update_equipment(self) -> None:
        eq_id = self.equipment_id_edit.text().strip()
        if not eq_id:
            QtWidgets.QMessageBox.warning(self, "无效输入", "设备编号不能为空。")
            return
        category = self.equipment_category_combo.currentText().strip()
        total = int(self.equipment_total_spin.value())
        available = int(self.equipment_available_spin.value())
        if available > total:
            available = total
        existing = next((e for e in self.equipment if e.equipment_id == eq_id), None)
        if existing:
            existing.category = category
            existing.total_count = total
            existing.available_count = available
        else:
            self.equipment.append(Equipment(eq_id, category, total, available))
        if eq_id:
            action = "更新" if existing else "添加"
            self._log_change(
                f"{action}设备 {eq_id} (类别: {category or '-'}, 总: {total}, 可用: {available})"
            )
        self._sync_equipment_to_order()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self.refresh_eta()
        self._auto_save()

    def remove_equipment(self) -> None:
        row = self.equipment_table.currentRow()
        if row < 0 or row >= len(self.equipment):
            return
        eq = self.equipment[row]
        self._log_change(
            f"删除设备 {eq.equipment_id} (类别: {eq.category or '-'}, 总: {eq.total_count}, 可用: {eq.available_count})"
        )
        del self.equipment[row]
        self._sync_equipment_to_order()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self.refresh_eta()
        self._auto_save()

    def _sync_equipment_to_order(self) -> None:
        if self.order:
            self.order.equipment = list(self.equipment)

    def _refresh_equipment_table(self) -> None:
        self.equipment_table.setRowCount(0)
        for eq in self.equipment:
            row = self.equipment_table.rowCount()
            self.equipment_table.insertRow(row)
            self.equipment_table.setItem(row, 0, QtWidgets.QTableWidgetItem(eq.equipment_id))
            self.equipment_table.setItem(row, 1, QtWidgets.QTableWidgetItem(eq.category or "-"))
            self.equipment_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(eq.total_count)))
            self.equipment_table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(eq.available_count)))

    # ------------------------
    # Employees management
    # ------------------------

    def add_employee(self) -> None:
        name = self.employee_name_edit.text().strip()
        if not name:
            return
        if name not in self.employees:
            self.employees.append(name)
            self._log_change(f"添加员工 {name}")
        self.employee_name_edit.clear()
        self._sync_employees_to_order()
        self._refresh_employee_list()
        self._refresh_phase_employee_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self._auto_save()

    def remove_employee(self) -> None:
        items = self.employee_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.employees:
            self.employees.remove(name)
            self._log_change(f"删除员工 {name}")
        self._sync_employees_to_order()
        self._refresh_employee_list()
        self._refresh_phase_employee_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self._auto_save()

    def _sync_employees_to_order(self) -> None:
        if self.order:
            self.order.employees = list(self.employees)

    def _refresh_employee_list(self) -> None:
        self.employee_list.clear()
        self.employee_list.addItems(self.employees)

    def _refresh_phase_employee_combo(self) -> None:
        current = self.phase_employee_combo.currentText()
        self.phase_employee_combo.clear()
        self.phase_employee_combo.addItems(self.employees)
        if current:
            self.phase_employee_combo.setCurrentText(current)

    def _refresh_phase_equipment_combo(self) -> None:
        current = self.phase_equipment_combo.currentText()
        self.phase_equipment_combo.clear()
        self.phase_equipment_combo.addItems([e.equipment_id for e in self.equipment])
        if current:
            self.phase_equipment_combo.setCurrentText(current)

    def _refresh_equipment_category_combos(self) -> None:
        if hasattr(self, "equipment_category_combo"):
            current = self.equipment_category_combo.currentText()
            self.equipment_category_combo.clear()
            self.equipment_category_combo.addItems(self.equipment_categories)
            if current:
                self.equipment_category_combo.setCurrentText(current)
        if hasattr(self, "admin_equipment_category_combo"):
            current_admin = self.admin_equipment_category_combo.currentText()
            self.admin_equipment_category_combo.clear()
            self.admin_equipment_category_combo.addItems(self.equipment_categories)
            if current_admin:
                self.admin_equipment_category_combo.setCurrentText(current_admin)

    # ------------------------
    # Products management
    # ------------------------

    def on_product_select(self) -> None:
        if not self.order:
            return
        row = self.products_table.currentRow()
        if row < 0 or row >= len(self.order.products):
            self._update_phase_product_label(None)
            return
        product = self.order.products[row]
        self.active_product_id = product.product_id
        self._update_phase_product_label(product)
        self.product_id_edit.setText(product.product_id)
        self.product_part_edit.setText(product.part_number)
        self.product_qty_spin.setValue(product.quantity)
        self._refresh_phase_table(product)

    def add_or_update_product(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先创建或加载订单。")
            return
        product_id = self.product_id_edit.text().strip()
        if not product_id:
            QtWidgets.QMessageBox.warning(self, "无效输入", "产品名不能为空。")
            return
        part_number = self.product_part_edit.text().strip()
        quantity = int(self.product_qty_spin.value())

        row = self.products_table.currentRow()
        if 0 <= row < len(self.order.products):
            product = self.order.products[row]
            old_id = product.product_id
            old_part = product.part_number
            old_qty = product.quantity
            product.product_id = product_id
            product.part_number = part_number
            product.quantity = quantity
            if old_id != product_id:
                self._log_change(
                    f"更新产品 {old_id} -> {product_id} (零件号: {part_number or '-'}, 数量: {quantity})"
                )
            elif old_part != part_number or old_qty != quantity:
                self._log_change(
                    f"更新产品 {product_id} (零件号: {part_number or '-'}, 数量: {quantity})"
                )
        else:
            self.order.products.append(Product(product_id, part_number, quantity))
            self._log_change(
                f"添加产品 {product_id} (零件号: {part_number or '-'}, 数量: {quantity})"
            )
        self.active_product_id = product_id
        self._refresh_products_table()
        self._refresh_admin_product_combo()
        self._refresh_defect_product_combo()
        self.refresh_eta()
        self._auto_save()

    def remove_product(self) -> None:
        if not self.order:
            return
        row = self.products_table.currentRow()
        if row < 0 or row >= len(self.order.products):
            return
        product = self.order.products[row]
        self._log_change(f"删除产品 {product.product_id}")
        del self.order.products[row]
        if self.order.products:
            next_row = min(row, len(self.order.products) - 1)
            self.active_product_id = self.order.products[next_row].product_id
        else:
            self.active_product_id = ""
        self._refresh_products_table()
        self._refresh_admin_product_combo()
        self._refresh_defect_product_combo()
        self.phases_table.setRowCount(0)
        self.refresh_eta()
        self._auto_save()

    def _refresh_products_table(self) -> None:
        self.products_table.setRowCount(0)
        if not self.order:
            self.active_product_id = ""
            self._update_phase_product_label(None)
            return
        current_id = self.active_product_id
        current_row = self.products_table.currentRow()
        if 0 <= current_row < len(self.order.products):
            current_id = self.order.products[current_row].product_id
        equipment_map = _equipment_available_map(self.order)
        for product in self.order.products:
            row = self.products_table.rowCount()
            self.products_table.insertRow(row)
            progress = _product_progress(product, equipment_map)
            self.products_table.setItem(row, 0, QtWidgets.QTableWidgetItem(product.product_id))
            self.products_table.setItem(row, 1, QtWidgets.QTableWidgetItem(product.part_number))
            self.products_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(product.quantity)))
            self.products_table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{progress:.0%}"))
        if not current_id and self.order.products:
            current_id = self.order.products[0].product_id
        if current_id:
            self.active_product_id = current_id
            self._select_product_by_id(current_id)
        else:
            self._update_phase_product_label(None)

    # ------------------------
    # Phases management
    # ------------------------

    def _current_product(self) -> Optional[Product]:
        if not self.order:
            return None
        row = self.products_table.currentRow()
        if 0 <= row < len(self.order.products):
            product = self.order.products[row]
            self.active_product_id = product.product_id
            return product
        if self.active_product_id:
            for product in self.order.products:
                if product.product_id == self.active_product_id:
                    self._select_product_by_id(self.active_product_id)
                    return product
        return None

    def on_phase_select(self) -> None:
        product = self._current_product()
        if not product:
            return
        row = self.phases_table.currentRow()
        if row < 0 or row >= len(product.phases):
            return
        phase = product.phases[row]
        self.phase_name_edit.setText(phase.name)
        self.phase_hours_spin.setValue(phase.planned_hours)
        self.phase_equipment_combo.setCurrentText(phase.equipment_id)
        self.phase_employee_combo.setCurrentText(phase.assigned_employee)
        self.phase_parallel_spin.setValue(phase.parallel_group)
        self.phase_completed_spin.setValue(phase.completed_hours)

    def add_or_update_phase(self) -> None:
        product = self._current_product()
        if not product:
            QtWidgets.QMessageBox.warning(self, "未选择产品", "请先选择一个产品。")
            return
        name = self.phase_name_edit.text().strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "无效输入", "工序名称不能为空。")
            return
        hours = float(self.phase_hours_spin.value())
        equipment_id = self.phase_equipment_combo.currentText().strip()
        employee = self.phase_employee_combo.currentText().strip()
        parallel = int(self.phase_parallel_spin.value())
        completed_hours = float(self.phase_completed_spin.value())
        planned_total = max(0.0, hours) * max(product.quantity, 1)
        if planned_total > 0:
            completed_hours = min(max(completed_hours, 0.0), planned_total)
        else:
            completed_hours = 0.0

        row = self.phases_table.currentRow()
        if 0 <= row < len(product.phases):
            phase = product.phases[row]
            phase.name = name
            phase.planned_hours = hours
            phase.equipment_id = equipment_id
            phase.assigned_employee = employee
            phase.parallel_group = parallel
            phase.completed_hours = completed_hours
            self._log_change(
                f"更新工序 {name} (产品: {product.product_id}, 完成工时: {completed_hours:g}h/{planned_total:g}h)"
            )
        else:
            product.phases.append(
                Phase(
                    name=name,
                    planned_hours=hours,
                    completed_hours=completed_hours,
                    equipment_id=equipment_id,
                    assigned_employee=employee,
                    parallel_group=parallel,
                )
            )
            self._log_change(
                f"添加工序 {name} (产品: {product.product_id}, 完成工时: {completed_hours:g}h/{planned_total:g}h)"
            )
        self.phase_completed_spin.setValue(completed_hours)
        self._refresh_phase_table(product)
        self._refresh_products_table()
        self.refresh_eta()
        self._auto_save()

    def remove_phase(self) -> None:
        product = self._current_product()
        if not product:
            return
        row = self.phases_table.currentRow()
        if row < 0 or row >= len(product.phases):
            return
        phase = product.phases[row]
        self._log_change(f"删除工序 {phase.name} (产品: {product.product_id})")
        del product.phases[row]
        self._refresh_phase_table(product)
        self._refresh_products_table()
        self.refresh_eta()
        self._auto_save()

    def set_parallel_group(self) -> None:
        product = self._current_product()
        if not product:
            return
        rows = {item.row() for item in self.phases_table.selectedItems()}
        if len(rows) < 2:
            QtWidgets.QMessageBox.information(self, "提示", "请至少选择两个工序来设置并行组。")
            return
        max_group = max((p.parallel_group for p in product.phases), default=0)
        new_group = max_group + 1
        for row in rows:
            if 0 <= row < len(product.phases):
                product.phases[row].parallel_group = new_group
        self._log_change(
            f"设置并行组 {new_group} (产品: {product.product_id}, 工序数: {len(rows)})"
        )
        self._refresh_phase_table(product)
        self.refresh_eta()
        self._auto_save()

    def clear_parallel_group(self) -> None:
        product = self._current_product()
        if not product:
            return
        rows = {item.row() for item in self.phases_table.selectedItems()}
        if not rows:
            return
        for row in rows:
            if 0 <= row < len(product.phases):
                product.phases[row].parallel_group = 0
        self._log_change(
            f"取消并行 (产品: {product.product_id}, 工序数: {len(rows)})"
        )
        self._refresh_phase_table(product)
        self.refresh_eta()
        self._auto_save()

    def _refresh_phase_table(self, product: Product) -> None:
        self.phases_table.setRowCount(0)
        for phase in product.phases:
            row = self.phases_table.rowCount()
            self.phases_table.insertRow(row)
            self.phases_table.setItem(row, 0, QtWidgets.QTableWidgetItem(phase.name))
            self.phases_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{phase.planned_hours:g}"))
            self.phases_table.setItem(row, 2, QtWidgets.QTableWidgetItem(phase.equipment_id))
            self.phases_table.setItem(row, 3, QtWidgets.QTableWidgetItem(phase.assigned_employee))
            self.phases_table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(phase.parallel_group)))
            progress = int(round(_phase_completion_ratio(phase, product.quantity) * 100))
            progress = min(max(progress, 0), 100)
            bar = QtWidgets.QProgressBar()
            bar.setRange(0, 100)
            bar.setValue(progress)
            bar.setFormat(f"{progress}%")
            bar.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            bar.setProperty("complete", progress >= 100)
            bar.style().unpolish(bar)
            bar.style().polish(bar)
            self.phases_table.setCellWidget(row, 5, bar)

    # ------------------------
    # Events
    # ------------------------

    def add_event(self) -> None:
        if not self.order:
            return
        day = _qdate_to_date(self.event_date_edit.date())
        hours = float(self.event_hours_spin.value())
        reason = self.event_reason_combo.currentText().strip() or "事件"
        self.order.events.append(Event(day=day, hours_lost=hours, reason=reason))
        self._log_change(f"添加事件 {day.isoformat()} {hours:g}h {reason}")
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def remove_event(self) -> None:
        if not self.order:
            return
        row = self.events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        self._log_change(f"删除事件 {ev.day.isoformat()} {ev.hours_lost:g}h {ev.reason}")
        del self.order.events[row]
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def _refresh_events_table(self) -> None:
        self.events_table.setRowCount(0)
        if not self.order:
            return
        for ev in self.order.events:
            row = self.events_table.rowCount()
            self.events_table.insertRow(row)
            self.events_table.setItem(row, 0, QtWidgets.QTableWidgetItem(ev.day.isoformat()))
            self.events_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{ev.hours_lost:g}"))
            self.events_table.setItem(row, 2, QtWidgets.QTableWidgetItem(ev.reason))

    # ------------------------
    # Defects
    # ------------------------

    def _refresh_defect_product_combo(self) -> None:
        current = self.defect_product_combo.currentText()
        self.defect_product_combo.clear()
        if self.order:
            self.defect_product_combo.addItems([p.product_id for p in self.order.products])
        if current:
            self.defect_product_combo.setCurrentText(current)

    def _refresh_defect_detail_combos(self) -> None:
        eq_current = self.defect_detail_equipment_combo.currentText()
        emp_current = self.defect_detail_employee_combo.currentText()
        self.defect_detail_equipment_combo.clear()
        self.defect_detail_equipment_combo.addItems([e.equipment_id for e in self.equipment])
        if eq_current:
            self.defect_detail_equipment_combo.setCurrentText(eq_current)
        self.defect_detail_employee_combo.clear()
        self.defect_detail_employee_combo.addItems(self.employees)
        if emp_current:
            self.defect_detail_employee_combo.setCurrentText(emp_current)

    def _refresh_defect_category_combo(self) -> None:
        current = self.defect_category_combo.currentText()
        self.defect_category_combo.clear()
        categories = list(self.defect_categories)
        if self.order and any(d.category == "未分类" for d in self.order.defects):
            if "未分类" not in categories:
                categories.append("未分类")
        self.defect_category_combo.addItems(categories)
        if current:
            self.defect_category_combo.setCurrentText(current)
        self.on_defect_category_change(self.defect_category_combo.currentText())

    def on_defect_category_change(self, text: str) -> None:
        if text == "设备":
            self.defect_detail_stack.setCurrentWidget(self.defect_detail_equipment_combo)
        elif text == "员工":
            self.defect_detail_stack.setCurrentWidget(self.defect_detail_employee_combo)
        else:
            self.defect_detail_stack.setCurrentWidget(self.defect_detail_edit)

    def _defect_detail_value(self) -> str:
        current = self.defect_category_combo.currentText()
        if current == "设备":
            return self.defect_detail_equipment_combo.currentText().strip()
        if current == "员工":
            return self.defect_detail_employee_combo.currentText().strip()
        return self.defect_detail_edit.text().strip()

    def add_defect(self) -> None:
        if not self.order:
            return
        product_id = self.defect_product_combo.currentText().strip()
        if not product_id:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择产品。")
            return
        count = int(self.defect_count_spin.value())
        category = self.defect_category_combo.currentText().strip()
        detail = self._defect_detail_value()
        self.order.defects.append(
            DefectRecord(product_id=product_id, count=count, category=category, detail=detail)
        )
        self._log_change(
            f"添加不合格 {product_id} 数量{count} 原因{category or '-'} {detail or '-'}"
        )
        self._refresh_defects_table()
        self._refresh_visuals()
        self._auto_save()

    def on_defect_select(self) -> None:
        if not self.order:
            return
        row = self.defects_table.currentRow()
        if row < 0 or row >= len(self.order.defects):
            return
        defect = self.order.defects[row]
        self.defect_product_combo.setCurrentText(defect.product_id)
        if defect.category:
            self.defect_category_combo.setCurrentText(defect.category)
        self.defect_count_spin.setValue(defect.count)
        self.on_defect_category_change(self.defect_category_combo.currentText())
        if defect.category == "设备":
            self.defect_detail_equipment_combo.setCurrentText(defect.detail)
        elif defect.category == "员工":
            self.defect_detail_employee_combo.setCurrentText(defect.detail)
        else:
            self.defect_detail_edit.setText(defect.detail)

    def update_defect(self) -> None:
        if not self.order:
            return
        row = self.defects_table.currentRow()
        if row < 0 or row >= len(self.order.defects):
            return
        defect = self.order.defects[row]
        defect.product_id = self.defect_product_combo.currentText().strip()
        defect.count = int(self.defect_count_spin.value())
        defect.category = self.defect_category_combo.currentText().strip()
        defect.detail = self._defect_detail_value()
        self._log_change(
            f"更新不合格 {defect.product_id} 数量{defect.count} 原因{defect.category or '-'} {defect.detail or '-'}"
        )
        self._refresh_defects_table()
        self._refresh_visuals()
        self._auto_save()

    def remove_defect(self) -> None:
        if not self.order:
            return
        row = self.defects_table.currentRow()
        if row < 0 or row >= len(self.order.defects):
            return
        defect = self.order.defects[row]
        self._log_change(
            f"删除不合格 {defect.product_id} 数量{defect.count} 原因{defect.category or '-'}"
        )
        del self.order.defects[row]
        self._refresh_defects_table()
        self._refresh_visuals()
        self._auto_save()

    def _refresh_defects_table(self) -> None:
        self.defects_table.setRowCount(0)
        if not self.order:
            return
        for defect in self.order.defects:
            row = self.defects_table.rowCount()
            self.defects_table.insertRow(row)
            self.defects_table.setItem(row, 0, QtWidgets.QTableWidgetItem(defect.product_id))
            self.defects_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(defect.count)))
            self.defects_table.setItem(
                row, 2, QtWidgets.QTableWidgetItem(defect.category or "未分类")
            )
            self.defects_table.setItem(row, 3, QtWidgets.QTableWidgetItem(defect.detail))
            self.defects_table.setItem(
                row, 4, QtWidgets.QTableWidgetItem(defect.timestamp.strftime("%Y-%m-%d %H:%M"))
            )

    # ------------------------
    # Admin: Events & Reasons
    # ------------------------

    def _refresh_admin_events_table(self) -> None:
        self.admin_events_table.setRowCount(0)
        if not self.order:
            return
        for ev in self.order.events:
            row = self.admin_events_table.rowCount()
            self.admin_events_table.insertRow(row)
            self.admin_events_table.setItem(row, 0, QtWidgets.QTableWidgetItem(ev.day.isoformat()))
            self.admin_events_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{ev.hours_lost:g}"))
            self.admin_events_table.setItem(row, 2, QtWidgets.QTableWidgetItem(ev.reason))

    def on_admin_event_select(self) -> None:
        if not self.order:
            return
        row = self.admin_events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        self.admin_event_date_edit.setDate(_date_to_qdate(ev.day))
        self.admin_event_hours_spin.setValue(ev.hours_lost)
        self.admin_event_reason_combo.setCurrentText(ev.reason)

    def admin_add_event(self) -> None:
        if not self.order:
            return
        day = _qdate_to_date(self.admin_event_date_edit.date())
        hours = float(self.admin_event_hours_spin.value())
        reason = self.admin_event_reason_combo.currentText().strip() or "事件"
        self.order.events.append(Event(day=day, hours_lost=hours, reason=reason))
        self._log_change(f"添加事件 {day.isoformat()} {hours:g}h {reason}")
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def admin_update_event(self) -> None:
        if not self.order:
            return
        row = self.admin_events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        old_day = ev.day
        old_hours = ev.hours_lost
        old_reason = ev.reason
        ev.day = _qdate_to_date(self.admin_event_date_edit.date())
        ev.hours_lost = float(self.admin_event_hours_spin.value())
        ev.reason = self.admin_event_reason_combo.currentText().strip() or "事件"
        self._log_change(
            f"更新事件 {old_day.isoformat()} {old_hours:g}h {old_reason} -> "
            f"{ev.day.isoformat()} {ev.hours_lost:g}h {ev.reason}"
        )
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def admin_remove_event(self) -> None:
        if not self.order:
            return
        row = self.admin_events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        self._log_change(f"删除事件 {ev.day.isoformat()} {ev.hours_lost:g}h {ev.reason}")
        del self.order.events[row]
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def _refresh_admin_log_table(self) -> None:
        if not hasattr(self, "admin_log_table"):
            return
        self.admin_log_table.setRowCount(0)
        if not self.order:
            return
        for entry in self.order.logs:
            row = self.admin_log_table.rowCount()
            self.admin_log_table.insertRow(row)
            self.admin_log_table.setItem(
                row, 0, QtWidgets.QTableWidgetItem(entry.timestamp.strftime("%Y-%m-%d %H:%M:%S"))
            )
            self.admin_log_table.setItem(row, 1, QtWidgets.QTableWidgetItem(entry.user))
            self.admin_log_table.setItem(row, 2, QtWidgets.QTableWidgetItem(entry.content))

    def _refresh_admin_users_list(self) -> None:
        if not hasattr(self, "admin_user_list"):
            return
        self.admin_user_list.clear()
        for account in self.user_accounts:
            self.admin_user_list.addItem(account.username)

    def on_admin_user_select(self) -> None:
        if not hasattr(self, "admin_user_list"):
            return
        items = self.admin_user_list.selectedItems()
        if not items:
            return
        username = items[0].text()
        account = next((u for u in self.user_accounts if u.username == username), None)
        if not account:
            return
        self.admin_user_name_edit.setText(account.username)
        if account.username == "admin":
            self.admin_user_pass_edit.clear()
            self.admin_user_pass_edit.setPlaceholderText("管理员密码不可显示")
        else:
            self.admin_user_pass_edit.setPlaceholderText("密码")
            self.admin_user_pass_edit.setText(account.password)
        self._toggle_admin_password_visibility(
            self.admin_show_password_check.isChecked()
            if hasattr(self, "admin_show_password_check")
            else False
        )

    def admin_add_user(self) -> None:
        username = self.admin_user_name_edit.text().strip()
        password = self.admin_user_pass_edit.text()
        if not username or not password:
            QtWidgets.QMessageBox.information(self, "提示", "用户名和密码不能为空。")
            return
        if any(u.username == username for u in self.user_accounts):
            QtWidgets.QMessageBox.information(self, "提示", "该用户名已存在。")
            return
        self.user_accounts.append(UserAccount(username=username, password=password))
        self.admin_user_name_edit.clear()
        self.admin_user_pass_edit.clear()
        self._refresh_admin_users_list()
        self._save_app_templates()

    def admin_update_user(self) -> None:
        items = self.admin_user_list.selectedItems()
        if not items:
            return
        old_username = items[0].text()
        new_username = self.admin_user_name_edit.text().strip()
        new_password = self.admin_user_pass_edit.text()
        if not new_username or not new_password:
            QtWidgets.QMessageBox.information(self, "提示", "用户名和密码不能为空。")
            return
        if old_username == "admin" and new_username != "admin":
            QtWidgets.QMessageBox.information(self, "提示", "admin 用户名不可修改。")
            return
        if new_username != old_username and any(
            u.username == new_username for u in self.user_accounts
        ):
            QtWidgets.QMessageBox.information(self, "提示", "该用户名已存在。")
            return
        account = next((u for u in self.user_accounts if u.username == old_username), None)
        if not account:
            return
        if old_username == "admin" and new_password != account.password:
            secret, ok = QtWidgets.QInputDialog.getText(
                self,
                "秘钥验证",
                "请输入管理员秘钥",
                QtWidgets.QLineEdit.EchoMode.Password,
            )
            if not ok:
                return
            if secret != ADMIN_SECRET_KEY:
                QtWidgets.QMessageBox.warning(self, "验证失败", "秘钥不正确，无法修改 admin 密码。")
                return
        account.username = new_username
        account.password = new_password
        if self.current_user == old_username:
            self.current_user = new_username
        self.admin_user_name_edit.clear()
        self.admin_user_pass_edit.clear()
        self._refresh_admin_users_list()
        self._save_app_templates()

    def admin_remove_user(self) -> None:
        items = self.admin_user_list.selectedItems()
        if not items:
            return
        username = items[0].text()
        if username == "admin":
            QtWidgets.QMessageBox.information(self, "提示", "admin 用户无法删除。")
            return
        if len(self.user_accounts) <= 1:
            QtWidgets.QMessageBox.information(self, "提示", "至少保留一个用户。")
            return
        self.user_accounts = [u for u in self.user_accounts if u.username != username]
        self.admin_user_name_edit.clear()
        self.admin_user_pass_edit.clear()
        self._refresh_admin_users_list()
        self._save_app_templates()

    def _refresh_admin_reason_list(self) -> None:
        self.admin_reason_list.clear()
        self.admin_reason_list.addItems(self.event_reasons)

    def on_admin_reason_select(self) -> None:
        items = self.admin_reason_list.selectedItems()
        if not items:
            return
        self.admin_reason_edit.setText(items[0].text())

    def admin_add_reason(self) -> None:
        reason = self.admin_reason_edit.text().strip()
        if not reason:
            return
        if reason in self.event_reasons:
            QtWidgets.QMessageBox.information(self, "提示", "该原因已存在。")
            return
        self.event_reasons.append(reason)
        self.admin_reason_edit.clear()
        self._refresh_admin_reason_list()
        self._refresh_event_reason_combo()
        self._save_app_templates()
        self._auto_save()

    def admin_update_reason(self) -> None:
        items = self.admin_reason_list.selectedItems()
        if not items:
            return
        new_reason = self.admin_reason_edit.text().strip()
        if not new_reason:
            return
        old_reason = items[0].text()
        if new_reason != old_reason and new_reason in self.event_reasons:
            QtWidgets.QMessageBox.information(self, "提示", "该原因已存在。")
            return
        idx = self.event_reasons.index(old_reason)
        self.event_reasons[idx] = new_reason
        if self.order:
            for ev in self.order.events:
                if ev.reason == old_reason:
                    ev.reason = new_reason
            self._log_change(f"更新事件原因 {old_reason} -> {new_reason}")
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self._refresh_admin_reason_list()
        self._refresh_event_reason_combo()
        self._save_app_templates()
        self._auto_save()

    def admin_remove_reason(self) -> None:
        items = self.admin_reason_list.selectedItems()
        if not items:
            return
        reason = items[0].text()
        if reason in self.event_reasons:
            self.event_reasons.remove(reason)
        self._refresh_admin_reason_list()
        self._refresh_event_reason_combo()
        self._save_app_templates()
        self._auto_save()

    def _refresh_admin_defect_categories(self) -> None:
        self.admin_defect_category_list.clear()
        self.admin_defect_category_list.addItems(self.defect_categories)

    def on_admin_defect_category_select(self) -> None:
        items = self.admin_defect_category_list.selectedItems()
        if not items:
            return
        self.admin_defect_category_edit.setText(items[0].text())

    def admin_add_defect_category(self) -> None:
        name = self.admin_defect_category_edit.text().strip()
        if not name:
            return
        if name in self.defect_categories:
            QtWidgets.QMessageBox.information(self, "提示", "该原因已存在。")
            return
        self.defect_categories.append(name)
        self.admin_defect_category_edit.clear()
        self._refresh_admin_defect_categories()
        self._refresh_defect_category_combo()
        self._save_app_templates()

    def admin_update_defect_category(self) -> None:
        items = self.admin_defect_category_list.selectedItems()
        if not items:
            return
        new_name = self.admin_defect_category_edit.text().strip()
        if not new_name:
            return
        old_name = items[0].text()
        if new_name != old_name and new_name in self.defect_categories:
            QtWidgets.QMessageBox.information(self, "提示", "该原因已存在。")
            return
        idx = self.defect_categories.index(old_name)
        self.defect_categories[idx] = new_name
        if self.order:
            for defect in self.order.defects:
                if defect.category == old_name:
                    defect.category = new_name
        self._refresh_admin_defect_categories()
        self._refresh_defect_category_combo()
        self._refresh_defects_table()
        self._refresh_visuals()
        self._save_app_templates()
        self._auto_save()

    def admin_remove_defect_category(self) -> None:
        items = self.admin_defect_category_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.defect_categories:
            self.defect_categories.remove(name)
        if self.order:
            for defect in self.order.defects:
                if defect.category == name:
                    defect.category = "未分类"
        self._refresh_admin_defect_categories()
        self._refresh_defect_category_combo()
        self._refresh_defects_table()
        self._refresh_visuals()
        self._save_app_templates()
        self._auto_save()

    def _refresh_admin_equipment_categories(self) -> None:
        self.admin_equipment_category_list.clear()
        self.admin_equipment_category_list.addItems(self.equipment_categories)

    def on_admin_equipment_category_select(self) -> None:
        items = self.admin_equipment_category_list.selectedItems()
        if not items:
            return
        self.admin_equipment_category_edit.setText(items[0].text())

    def admin_add_equipment_category(self) -> None:
        name = self.admin_equipment_category_edit.text().strip()
        if not name:
            return
        if name in self.equipment_categories:
            QtWidgets.QMessageBox.information(self, "提示", "该分类已存在。")
            return
        self.equipment_categories.append(name)
        self.admin_equipment_category_edit.clear()
        self._refresh_admin_equipment_categories()
        self._refresh_equipment_category_combos()
        self._save_app_templates()

    def admin_update_equipment_category(self) -> None:
        items = self.admin_equipment_category_list.selectedItems()
        if not items:
            return
        new_name = self.admin_equipment_category_edit.text().strip()
        if not new_name:
            return
        old_name = items[0].text()
        if new_name != old_name and new_name in self.equipment_categories:
            QtWidgets.QMessageBox.information(self, "提示", "该分类已存在。")
            return
        idx = self.equipment_categories.index(old_name)
        self.equipment_categories[idx] = new_name
        for eq in self.equipment:
            if eq.category == old_name:
                eq.category = new_name
        for eq in self.equipment_templates:
            if eq.category == old_name:
                eq.category = new_name
        self._refresh_admin_equipment_categories()
        self._refresh_equipment_table()
        self._refresh_admin_equipment_templates_table()
        self._refresh_equipment_category_combos()
        self._save_app_templates()
        self._auto_save()

    def admin_remove_equipment_category(self) -> None:
        items = self.admin_equipment_category_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.equipment_categories:
            self.equipment_categories.remove(name)
        fallback = self.equipment_categories[0] if self.equipment_categories else ""
        for eq in self.equipment:
            if eq.category == name:
                eq.category = fallback
        for eq in self.equipment_templates:
            if eq.category == name:
                eq.category = fallback
        self._refresh_admin_equipment_categories()
        self._refresh_equipment_category_combos()
        self._refresh_equipment_table()
        self._refresh_admin_equipment_templates_table()
        self._save_app_templates()

    # ------------------------
    # Admin: Templates
    # ------------------------

    def _refresh_admin_equipment_templates_table(self) -> None:
        self.admin_equipment_template_table.setRowCount(0)
        for eq in self.equipment_templates:
            row = self.admin_equipment_template_table.rowCount()
            self.admin_equipment_template_table.insertRow(row)
            self.admin_equipment_template_table.setItem(row, 0, QtWidgets.QTableWidgetItem(eq.equipment_id))
            self.admin_equipment_template_table.setItem(row, 1, QtWidgets.QTableWidgetItem(eq.category or "-"))
            self.admin_equipment_template_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(eq.total_count)))
            self.admin_equipment_template_table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(eq.available_count)))

    def on_admin_equipment_template_select(self) -> None:
        row = self.admin_equipment_template_table.currentRow()
        if row < 0 or row >= len(self.equipment_templates):
            return
        eq = self.equipment_templates[row]
        self.admin_equipment_id_edit.setText(eq.equipment_id)
        self.admin_equipment_category_combo.setCurrentText(eq.category)
        self.admin_equipment_total_spin.setValue(eq.total_count)
        self.admin_equipment_available_spin.setValue(eq.available_count)

    def admin_add_or_update_equipment_template(self) -> None:
        eq_id = self.admin_equipment_id_edit.text().strip()
        if not eq_id:
            return
        category = self.admin_equipment_category_combo.currentText().strip()
        total = int(self.admin_equipment_total_spin.value())
        available = int(self.admin_equipment_available_spin.value())
        if available > total:
            available = total
        existing = next((e for e in self.equipment_templates if e.equipment_id == eq_id), None)
        if existing:
            existing.category = category
            existing.total_count = total
            existing.available_count = available
        else:
            self.equipment_templates.append(Equipment(eq_id, category, total, available))
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()

    def admin_remove_equipment_template(self) -> None:
        row = self.admin_equipment_template_table.currentRow()
        if row < 0 or row >= len(self.equipment_templates):
            return
        del self.equipment_templates[row]
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()

    def admin_apply_equipment_template(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        self.equipment = [
            Equipment(e.equipment_id, e.category, e.total_count, e.available_count)
            for e in self.equipment_templates
        ]
        self._log_change(f"应用设备模板 (设备数: {len(self.equipment)})")
        self._sync_equipment_to_order()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self.refresh_eta()
        self._auto_save()

    def _refresh_admin_phase_templates_table(self) -> None:
        self.admin_phase_template_table.setRowCount(0)
        for phase in self.phase_templates:
            row = self.admin_phase_template_table.rowCount()
            self.admin_phase_template_table.insertRow(row)
            self.admin_phase_template_table.setItem(row, 0, QtWidgets.QTableWidgetItem(phase.name))
            self.admin_phase_template_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{phase.planned_hours:g}"))
            self.admin_phase_template_table.setItem(row, 2, QtWidgets.QTableWidgetItem(phase.equipment_id))
            self.admin_phase_template_table.setItem(row, 3, QtWidgets.QTableWidgetItem(phase.assigned_employee))
            self.admin_phase_template_table.setItem(row, 4, QtWidgets.QTableWidgetItem(str(phase.parallel_group)))

    def on_admin_phase_template_select(self) -> None:
        row = self.admin_phase_template_table.currentRow()
        if row < 0 or row >= len(self.phase_templates):
            return
        phase = self.phase_templates[row]
        self.admin_phase_name_edit.setText(phase.name)
        self.admin_phase_hours_spin.setValue(phase.planned_hours)
        self.admin_phase_equipment_combo.setCurrentText(phase.equipment_id)
        self.admin_phase_employee_combo.setCurrentText(phase.assigned_employee)
        self.admin_phase_parallel_spin.setValue(phase.parallel_group)

    def admin_add_or_update_phase_template(self) -> None:
        name = self.admin_phase_name_edit.text().strip()
        if not name:
            return
        hours = float(self.admin_phase_hours_spin.value())
        equipment_id = self.admin_phase_equipment_combo.currentText().strip()
        employee = self.admin_phase_employee_combo.currentText().strip()
        parallel = int(self.admin_phase_parallel_spin.value())

        row = self.admin_phase_template_table.currentRow()
        if 0 <= row < len(self.phase_templates):
            phase = self.phase_templates[row]
            phase.name = name
            phase.planned_hours = hours
            phase.equipment_id = equipment_id
            phase.assigned_employee = employee
            phase.parallel_group = parallel
            phase.completed_hours = 0.0
        else:
            self.phase_templates.append(
                Phase(
                    name=name,
                    planned_hours=hours,
                    equipment_id=equipment_id,
                    assigned_employee=employee,
                    parallel_group=parallel,
                )
            )
        self._refresh_admin_phase_templates_table()
        self._save_app_templates()

    def admin_remove_phase_template(self) -> None:
        row = self.admin_phase_template_table.currentRow()
        if row < 0 or row >= len(self.phase_templates):
            return
        del self.phase_templates[row]
        self._refresh_admin_phase_templates_table()
        self._save_app_templates()

    def admin_apply_phase_template(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        product_id = self.admin_phase_product_combo.currentText().strip()
        product = self._get_product_by_id(product_id)
        if not product:
            QtWidgets.QMessageBox.information(self, "提示", "请选择要应用的产品。")
            return
        product.phases = [
            Phase(
                name=ph.name,
                planned_hours=ph.planned_hours,
                equipment_id=ph.equipment_id,
                assigned_employee=ph.assigned_employee,
                parallel_group=ph.parallel_group,
            )
            for ph in self.phase_templates
        ]
        self._log_change(f"应用工序模板到产品 {product_id} (工序数: {len(product.phases)})")
        self._refresh_phase_table(product)
        self._refresh_products_table()
        self.refresh_eta()
        self._auto_save()

    def _refresh_admin_employee_templates_list(self) -> None:
        self.admin_employee_template_list.clear()
        self.admin_employee_template_list.addItems(self.employee_templates)

    def on_admin_employee_template_select(self) -> None:
        items = self.admin_employee_template_list.selectedItems()
        if not items:
            return
        self.admin_employee_name_edit.setText(items[0].text())

    def admin_add_or_update_employee_template(self) -> None:
        name = self.admin_employee_name_edit.text().strip()
        if not name:
            return
        items = self.admin_employee_template_list.selectedItems()
        if items:
            old_name = items[0].text()
            if name != old_name and name in self.employee_templates:
                QtWidgets.QMessageBox.information(self, "提示", "该员工已存在。")
                return
            idx = self.employee_templates.index(old_name)
            self.employee_templates[idx] = name
        else:
            if name in self.employee_templates:
                QtWidgets.QMessageBox.information(self, "提示", "该员工已存在。")
                return
            self.employee_templates.append(name)
        self.admin_employee_name_edit.clear()
        self._refresh_admin_employee_templates_list()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()

    def admin_remove_employee_template(self) -> None:
        items = self.admin_employee_template_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.employee_templates:
            self.employee_templates.remove(name)
        self._refresh_admin_employee_templates_list()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()

    def admin_apply_employee_template(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        self.employees = list(self.employee_templates)
        self._log_change(f"应用员工模板 (员工数: {len(self.employees)})")
        self._sync_employees_to_order()
        self._refresh_employee_list()
        self._refresh_phase_employee_combo()
        self._refresh_admin_phase_template_combos()
        self.refresh_eta()
        self._auto_save()

    def _refresh_admin_shift_templates_list(self) -> None:
        self.admin_shift_template_list.clear()
        for tpl in self.shift_templates:
            self.admin_shift_template_list.addItem(tpl.name)
        if self.active_shift_template_name:
            matches = self.admin_shift_template_list.findItems(
                self.active_shift_template_name, QtCore.Qt.MatchFlag.MatchExactly
            )
            if matches:
                self.admin_shift_template_list.setCurrentItem(matches[0])
        elif self.shift_templates:
            self.admin_shift_template_list.setCurrentRow(0)
        self._update_active_shift_label()

    def _update_active_shift_label(self) -> None:
        name = self.active_shift_template_name or "-"
        self.admin_shift_active_label.setText(f"当前班次: {name}")

    def _update_shift_row_total(self, row: int) -> None:
        if row < 0 or row >= len(self.shift_count_spins):
            return
        count = self.shift_count_spins[row].value()
        hours = self.shift_hours_spins[row].value()
        total_item = self.admin_shift_table.item(row, 3)
        if total_item:
            total_item.setText(f"{count * hours:g}")

    def _collect_shift_week_plan(self) -> List[ShiftDayPlan]:
        week_plan: List[ShiftDayPlan] = []
        for row in range(7):
            count = self.shift_count_spins[row].value()
            hours = self.shift_hours_spins[row].value()
            week_plan.append(ShiftDayPlan(count, hours))
        return week_plan

    def _load_shift_plan_to_table(self, week_plan: List[ShiftDayPlan]) -> None:
        for row in range(7):
            day = week_plan[row] if row < len(week_plan) else ShiftDayPlan(0, 0.0)
            self.shift_count_spins[row].setValue(day.shift_count)
            self.shift_hours_spins[row].setValue(day.hours_per_shift)
            self._update_shift_row_total(row)

    def on_admin_shift_template_select(self) -> None:
        items = self.admin_shift_template_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        tpl = next((t for t in self.shift_templates if t.name == name), None)
        if not tpl:
            return
        self.admin_shift_name_edit.setText(tpl.name)
        self._load_shift_plan_to_table(tpl.week_plan)

    def admin_add_or_update_shift_template(self) -> None:
        name = self.admin_shift_name_edit.text().strip()
        if not name:
            return
        week_plan = self._collect_shift_week_plan()
        items = self.admin_shift_template_list.selectedItems()
        if items:
            old_name = items[0].text()
            if name != old_name and any(t.name == name for t in self.shift_templates):
                QtWidgets.QMessageBox.information(self, "提示", "该班次模板已存在。")
                return
            tpl = next((t for t in self.shift_templates if t.name == old_name), None)
            if tpl:
                tpl.name = name
                tpl.week_plan = week_plan
                if self.active_shift_template_name == old_name:
                    self.active_shift_template_name = name
        else:
            if any(t.name == name for t in self.shift_templates):
                QtWidgets.QMessageBox.information(self, "提示", "该班次模板已存在。")
                return
            self.shift_templates.append(ShiftTemplate(name, week_plan))
        self._refresh_admin_shift_templates_list()
        self._save_app_templates()
        self._apply_active_shift_template()

    def admin_remove_shift_template(self) -> None:
        items = self.admin_shift_template_list.selectedItems()
        if not items:
            return
        if len(self.shift_templates) <= 1:
            QtWidgets.QMessageBox.information(self, "提示", "至少保留一个班次模板。")
            return
        name = items[0].text()
        self.shift_templates = [t for t in self.shift_templates if t.name != name]
        if self.active_shift_template_name == name:
            self.active_shift_template_name = self.shift_templates[0].name if self.shift_templates else ""
        self._refresh_admin_shift_templates_list()
        self._save_app_templates()
        self._apply_active_shift_template()

    def admin_set_active_shift_template(self) -> None:
        items = self.admin_shift_template_list.selectedItems()
        if not items:
            return
        self.active_shift_template_name = items[0].text()
        self._update_active_shift_label()
        self._save_app_templates()
        self._apply_active_shift_template()

    def _refresh_admin_product_combo(self) -> None:
        current = self.admin_phase_product_combo.currentText()
        self.admin_phase_product_combo.clear()
        if not self.order:
            return
        self.admin_phase_product_combo.addItems([p.product_id for p in self.order.products])
        if current:
            self.admin_phase_product_combo.setCurrentText(current)

    def _refresh_admin_phase_template_combos(self) -> None:
        equipment_ids = {
            e.equipment_id for e in self.equipment_templates if e.equipment_id
        } | {e.equipment_id for e in self.equipment if e.equipment_id}
        employee_names = set(self.employee_templates) | set(self.employees)

        current_eq = self.admin_phase_equipment_combo.currentText()
        self.admin_phase_equipment_combo.clear()
        self.admin_phase_equipment_combo.addItems(sorted(equipment_ids))
        if current_eq:
            self.admin_phase_equipment_combo.setCurrentText(current_eq)

        current_emp = self.admin_phase_employee_combo.currentText()
        self.admin_phase_employee_combo.clear()
        self.admin_phase_employee_combo.addItems(sorted(employee_names))
        if current_emp:
            self.admin_phase_employee_combo.setCurrentText(current_emp)

    def _get_product_by_id(self, product_id: str) -> Optional[Product]:
        if not self.order:
            return None
        for product in self.order.products:
            if product.product_id == product_id:
                return product
        return None

    # ------------------------
    # ETA / Progress
    # ------------------------

    def refresh_eta(self) -> None:
        if not self.order:
            self.eta_value.setText("-")
            self.remaining_value.setText("-")
            self.detail_eta_label.setText("预计交期: -")
            self.detail_progress_label.setText("总体进度: 0%")
            self.overall_progress.setValue(0)
            self.last_eta_dt = None
            self.last_remaining_hours = 0.0
            self._refresh_visuals()
            return
        if not self.cal.shift_template:
            self.eta_value.setText("请先设置班次模板")
            self.remaining_value.setText("-")
            self.detail_eta_label.setText("预计交期: 请先设置班次模板")
            self.detail_progress_label.setText("总体进度: -")
            self.overall_progress.setValue(0)
            self.last_eta_dt = None
            self.last_remaining_hours = 0.0
            self._refresh_visuals()
            return
        try:
            result = compute_eta(self.order, self.cal)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "交期计算错误", str(exc))
            return

        eta_dt: datetime = result["eta_dt"]
        remaining_hours: float = result["remaining_hours"]
        self.last_eta_dt = eta_dt
        self.last_remaining_hours = remaining_hours
        self.eta_value.setText(eta_dt.strftime("%Y-%m-%d %H:%M"))
        self.remaining_value.setText(f"{remaining_hours:g}h")
        self.detail_eta_label.setText(f"预计交期: {eta_dt.strftime('%Y-%m-%d %H:%M')}")

        equipment_map = _equipment_available_map(self.order)
        total = 0.0
        done = 0.0
        for product in self.order.products:
            for phase in product.phases:
                hours = _phase_effective_hours(phase, product.quantity, equipment_map)
                total += hours
                ratio = _phase_completion_ratio(phase, product.quantity)
                done += hours * ratio
        progress = (done / total) if total > 0 else 0.0
        self.overall_progress.setValue(int(progress * 100))
        self.detail_progress_label.setText(f"总体进度: {progress:.0%}")
        self._refresh_products_table()
        self._refresh_dashboard_product_progress()
        self._refresh_order_summary()
        self._refresh_visuals()

    def _refresh_dashboard_product_progress(self) -> None:
        self.product_progress_table.setRowCount(0)
        if not self.order:
            return
        equipment_map = _equipment_available_map(self.order)
        for product in self.order.products:
            row = self.product_progress_table.rowCount()
            self.product_progress_table.insertRow(row)
            self.product_progress_table.setItem(row, 0, QtWidgets.QTableWidgetItem(product.product_id))
            self.product_progress_table.setItem(row, 1, QtWidgets.QTableWidgetItem(product.part_number))
            progress = _product_progress(product, equipment_map)
            self.product_progress_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{progress:.0%}"))

    def _refresh_order_summary(self) -> None:
        if not self.order:
            self.order_summary_label.setText("订单: -")
            self._refresh_dashboard_summary()
            return
        self.order_summary_label.setText(
            f"订单: {self.order.order_id} | 产品数: {len(self.order.products)} | 设备数: {len(self.order.equipment)}"
        )
        self._refresh_dashboard_summary()

    def _refresh_dashboard_summary(self) -> None:
        shift_name = self.active_shift_template_name or "-"
        if hasattr(self, "shift_summary_label"):
            self.shift_summary_label.setText(f"当前班次: {shift_name}")
        if hasattr(self, "dashboard_summary_label"):
            if not self.order:
                self.dashboard_summary_label.setText("产品数: 0 | 设备数: 0")
            else:
                self.dashboard_summary_label.setText(
                    f"产品数: {len(self.order.products)} | 设备数: {len(self.order.equipment)}"
                )

    # ------------------------
    # Serialization
    # ------------------------

    def _order_to_dict(self, order: Order) -> Dict[str, object]:
        return {
            "version": 8,
            "order_id": order.order_id,
            "start_dt": order.start_dt.isoformat(),
            "equipment": [
                {
                    "equipment_id": e.equipment_id,
                    "category": e.category,
                    "total_count": e.total_count,
                    "available_count": e.available_count,
                }
                for e in order.equipment
            ],
            "employees": order.employees,
            "products": [
                {
                    "product_id": p.product_id,
                    "part_number": p.part_number,
                    "quantity": p.quantity,
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
                }
                for e in order.events
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
            "logs": [
                {
                    "timestamp": l.timestamp.isoformat(),
                    "user": l.user,
                    "content": l.content,
                }
                for l in order.logs
            ],
        }

    @staticmethod
    def _order_from_dict(data: Dict[str, object]) -> Order:
        order_id = data.get("order_id", "O-UNKNOWN")
        start_dt = datetime.fromisoformat(data.get("start_dt", datetime.now().isoformat()))

        equipment = [
            Equipment(
                equipment_id=e.get("equipment_id", ""),
                category=e.get("category", ""),
                total_count=int(e.get("total_count", 1)),
                available_count=int(e.get("available_count", 1)),
            )
            for e in data.get("equipment", [])
        ]

        employees = list(data.get("employees", []))

        products: List[Product] = []
        if "products" in data:
            for p in data.get("products", []):
                quantity = int(p.get("quantity", 1))
                phases = []
                for ph in p.get("phases", []):
                    planned_hours = float(ph.get("planned_hours", 0))
                    completed = ph.get("completed_hours", None)
                    if completed is None:
                        completed = planned_hours * max(quantity, 1) if ph.get("done", False) else 0.0
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
                        phases=phases,
                    )
                )
        elif "phases" in data:
            # Backward compatibility: old single-product format
            quantity = int(data.get("quantity", 1))
            phases = []
            for ph in data.get("phases", []):
                planned_hours = float(ph.get("planned_hours", 0))
                completed = ph.get("completed_hours", None)
                if completed is None:
                    completed = planned_hours * max(quantity, 1) if ph.get("done", False) else 0.0
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
                    phases=phases,
                )
            )

        events = [
            Event(
                day=date.fromisoformat(e.get("day")),
                hours_lost=float(e.get("hours_lost", 0)),
                reason=e.get("reason", ""),
            )
            for e in data.get("events", [])
            if e.get("day")
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
            )
            for l in data.get("logs", [])
        ]

        return Order(
            order_id=order_id,
            start_dt=start_dt,
            products=products,
            events=events,
            defects=defects,
            equipment=equipment,
            employees=employees,
            logs=logs,
        )

    # ------------------------
    # Global refresh
    # ------------------------

    def _refresh_all(self) -> None:
        self._refresh_equipment_table()
        self._refresh_employee_list()
        self._refresh_phase_equipment_combo()
        self._refresh_phase_employee_combo()
        self._refresh_equipment_category_combos()
        self._refresh_products_table()
        self._refresh_events_table()
        self._refresh_defects_table()
        self._refresh_defect_product_combo()
        self._refresh_defect_detail_combos()
        self._refresh_defect_category_combo()
        self._refresh_event_reason_combo()
        self._refresh_admin_views()
        self.refresh_eta()


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
