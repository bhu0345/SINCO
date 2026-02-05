import json
import os
from datetime import datetime, timedelta, date
from typing import List, Dict, Optional
from PySide6 import QtCore, QtGui, QtWidgets

from modules.models import (
    Phase,
    Product,
    Equipment,
    ShiftDayPlan,
    ShiftTemplate,
    Event,
    CapacityAdjustment,
    DefectRecord,
    LogEntry,
    MemoEntry,
    Order,
    UserAccount,
)
from modules.delegates import ComboBoxDelegate, SpinBoxDelegate
from modules.scheduler import (
    WorkCalendar,
    _equipment_available_map_from_list,
    _equipment_available_map,
    _normalize_equipment_ids,
    _split_equipment_ids,
    _format_equipment_ids,
    _phase_effective_hours,
    _phase_total_hours,
    _phase_completion_ratio,
    _product_remaining_hours,
    _product_progress,
    _product_quantity_progress,
    compute_eta,
)
from modules.data_io import order_to_dict, order_from_dict


QT_BINDING = "PySide6"
ADMIN_PASSWORD = "admin123"
ADMIN_SECRET_KEY = "s3cr3t_k3y"


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
        self.setWindowTitle("生产交期优化系统 V3.5")
        self.resize(1280, 820)

        self.locale_cn = _chinese_locale()
        QtCore.QLocale.setDefault(self.locale_cn)

        self.order: Optional[Order] = None
        self.orders: List[Order] = []
        self.factory_name = "默认工厂"
        self.equipment: List[Equipment] = []
        self.employees: List[str] = []
        self.event_reasons = ["员工请假", "设备故障", "停电", "材料短缺", "质量问题", "其他"]
        self.user_accounts: List[UserAccount] = []
        self.current_user = ""
        self.factory_path: Optional[str] = None
        self._updating_orders_table = False
        self.app_template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "app_templates.json"
        )
        self.equipment_categories = ["车床", "加工中心", "检验设备", "辅助设备"]
        self.defect_categories = ["设备", "原材料", "员工"]
        self.customer_codes = ["HYD", "SWI", "SCH"]
        self.shipping_methods = ["空运", "海运"]
        self.logo_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "GUI", "sinco.JPG"
        )
        self.equipment_templates: List[Equipment] = []
        self.phase_templates: List[Phase] = []
        self.phase_template_sets: Dict[str, List[Phase]] = {}
        self.active_phase_template_name = ""
        self.employee_templates: List[str] = []
        self.shift_templates: List[ShiftTemplate] = []
        self.active_shift_template_name = ""
        self.last_eta_dt: Optional[datetime] = None
        self.last_remaining_hours: float = 0.0
        self.last_capacity_map: Dict[date, float] = {}
        self.active_product_id: str = ""
        self._updating_products_table = False
        self._updating_phase_table = False
        self._updating_admin_phase_table = False
        self.app_logs: List[LogEntry] = []
        self.memos: List[MemoEntry] = []
        self._time_timer: Optional[QtCore.QTimer] = None
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
        self._setup_time_updates()

    def _setup_time_updates(self) -> None:
        if not hasattr(self, "beijing_time_label"):
            return
        self._update_beijing_time()
        timer = QtCore.QTimer(self)
        timer.timeout.connect(self._update_beijing_time)
        timer.start(60_000)
        self._time_timer = timer

    def _update_beijing_time(self) -> None:
        if not hasattr(self, "beijing_time_label"):
            return
        now = self._beijing_now()
        self.beijing_time_label.setText(f"北京时间: {now.strftime('%Y-%m-%d %H:%M')}")

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

    def _beijing_now(self) -> datetime:
        return datetime.utcnow() + timedelta(hours=8)

    def _equipment_id_list(self) -> List[str]:
        items = [e.equipment_id for e in self.equipment if e.equipment_id]
        return ["无需设备"] + sorted(set(items))

    def _phase_template_name_list(self) -> List[str]:
        seen = set()
        names: List[str] = []
        for ph in self.phase_templates:
            name = (ph.name or "").strip()
            if not name or name in seen:
                continue
            seen.add(name)
            names.append(name)
        return names

    def _employee_name_list(self) -> List[str]:
        items = [name for name in self.employees if name]
        return [""] + sorted(set(items))

    def _employee_template_list(self) -> List[str]:
        items = [name for name in self.employees if name]
        return [""] + sorted(set(items))

    def _update_phase_product_label(self, product: Optional[Product]) -> None:
        if not hasattr(self, "phase_product_label"):
            return
        if not product:
            self.phase_product_label.setText("当前产品: -")
            return
        part = product.part_number or "-"
        self.phase_product_label.setText(
            f"当前产品: {product.product_id} | 零件号: {part} | 要求数量: {product.quantity}"
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
            QLabel#detailOrderTitle { font-size: 18px; font-weight: 700; color: #0f5e9c; }
            QLabel#detailProgressBig { font-size: 16px; font-weight: 700; color: #1f2d3d; }
            QPushButton { background: #e8eef5; border: 1px solid #c8d4e2; padding: 6px 12px; border-radius: 6px; }
            QPushButton#primaryAction { background: #0f5e9c; color: #ffffff; border: none; }
            QPushButton#dangerAction { background: #c64545; color: #ffffff; border: none; }
            QPushButton#opsAction { background: #f5b942; color: #1f2d3d; border: none; }
            QPushButton#opsAction:hover { background: #e7aa2b; }
            QPushButton#phaseAction { background: #2f855a; color: #ffffff; border: none; }
            QPushButton#phaseAction:hover { background: #276749; }
            QToolButton#menuButton {
                background: #e8eef5; border: 1px solid #c8d4e2; padding: 6px 12px; border-radius: 10px;
            }
            QToolButton#menuButton:hover { background: #dde7f2; }
            QToolButton#menuButton:pressed { background: #d2ddea; }
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
        self.dashboard_summary_label = QtWidgets.QLabel("订单数: 0 | 设备数: 0")
        self.dashboard_summary_label.setObjectName("metricLabel")
        self.beijing_time_label = QtWidgets.QLabel("北京时间: -")
        self.beijing_time_label.setObjectName("metricLabel")
        stats_layout.addWidget(self.shift_summary_label)
        stats_layout.addWidget(self.dashboard_summary_label)
        stats_layout.addWidget(self.beijing_time_label)
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

        self.create_order_btn = QtWidgets.QPushButton("新增订单")
        self.create_order_btn.setObjectName("primaryAction")
        self.create_order_btn.clicked.connect(self.create_order)

        self.file_menu_btn = QtWidgets.QToolButton()
        self.file_menu_btn.setText("工厂文件")
        self.file_menu_btn.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)
        self.file_menu_btn.setObjectName("menuButton")
        file_menu = QtWidgets.QMenu(self.file_menu_btn)
        file_save_action = file_menu.addAction("保存工厂")
        file_save_action.triggered.connect(self.save_factory)
        file_load_action = file_menu.addAction("加载工厂")
        file_load_action.triggered.connect(self.load_factory)
        file_path_action = file_menu.addAction("显示工厂文件路径")
        file_path_action.triggered.connect(self.show_factory_path)
        self.file_menu_btn.setMenu(file_menu)

        self.visual_btn = QtWidgets.QPushButton("数据看板")
        self.visual_btn.clicked.connect(self.go_to_visuals)

        self.system_menu_btn = QtWidgets.QToolButton()
        self.system_menu_btn.setText("系统")
        self.system_menu_btn.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)
        self.system_menu_btn.setObjectName("menuButton")
        system_menu = QtWidgets.QMenu(self.system_menu_btn)
        admin_action = system_menu.addAction("管理员")
        admin_action.triggered.connect(self.open_admin_login)
        switch_action = system_menu.addAction("切换用户")
        switch_action.triggered.connect(self.switch_user)
        logout_action = system_menu.addAction("退出登录")
        logout_action.triggered.connect(self.logout_user)
        self.system_menu_btn.setMenu(system_menu)

        header_layout.addWidget(self.create_order_btn, 0, 4)
        header_layout.addWidget(self.file_menu_btn, 0, 5)
        header_layout.addWidget(self.visual_btn, 0, 6)
        self.log_btn = QtWidgets.QPushButton("备忘录")
        self.log_btn.clicked.connect(self.open_log_dialog)
        header_layout.addWidget(self.log_btn, 0, 7)
        header_layout.addWidget(self.system_menu_btn, 0, 8)

        layout.addWidget(header)

        order_list_group = QtWidgets.QGroupBox("订单列表")
        order_list_layout = QtWidgets.QVBoxLayout(order_list_group)

        self.orders_table = QtWidgets.QTableWidget(0, 7)
        self.orders_table.setHorizontalHeaderLabels(
            ["订单编号", "客户代码", "开始日期", "要求发货日期", "计划发货期", "剩余工时", "产品数"]
        )
        self.orders_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.orders_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.orders_table.verticalHeader().setVisible(False)
        self.orders_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.orders_table.setSortingEnabled(True)
        self.orders_table.itemSelectionChanged.connect(self.on_order_select)
        self.orders_table.itemDoubleClicked.connect(self.on_order_double_click)
        order_list_layout.addWidget(self.orders_table)

        order_actions = QtWidgets.QHBoxLayout()
        self.goto_detail_btn = QtWidgets.QPushButton("进入订单详情")
        self.goto_detail_btn.setObjectName("primaryAction")
        self.goto_detail_btn.clicked.connect(self.go_to_detail)
        self.edit_order_btn = QtWidgets.QPushButton("编辑订单")
        self.edit_order_btn.clicked.connect(self.open_order_editor)
        self.duplicate_order_btn = QtWidgets.QPushButton("复制订单")
        self.duplicate_order_btn.clicked.connect(self.duplicate_order)
        self.remove_order_btn = QtWidgets.QPushButton("删除订单")
        self.remove_order_btn.clicked.connect(self.remove_order)

        order_actions.addWidget(self.goto_detail_btn)
        order_actions.addWidget(self.edit_order_btn)
        order_actions.addWidget(self.duplicate_order_btn)
        order_actions.addWidget(self.remove_order_btn)
        order_list_layout.addLayout(order_actions)

        layout.addWidget(order_list_group)

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
        eq_form.setColumnStretch(1, 3)
        eq_form.setColumnStretch(3, 1)
        self.equipment_id_edit = QtWidgets.QLineEdit()
        self.equipment_id_edit.setMinimumWidth(160)
        self.equipment_category_combo = QtWidgets.QComboBox()
        self.equipment_category_combo.setEditable(False)
        self.equipment_category_combo.addItems(self.equipment_categories)
        self._configure_combo_popup(self.equipment_category_combo, 140)
        self.equipment_total_spin = QtWidgets.QSpinBox()
        self.equipment_total_spin.setRange(1, 9999)
        self.equipment_available_spin = QtWidgets.QSpinBox()
        self.equipment_available_spin.setRange(0, 9999)

        eq_id_label = QtWidgets.QLabel("设备编号")
        eq_cat_label = QtWidgets.QLabel("类别")
        eq_total_label = QtWidgets.QLabel("总数量")
        eq_avail_label = QtWidgets.QLabel("可用数量")
        eq_form.addWidget(eq_id_label, 0, 0)
        eq_form.addWidget(self.equipment_id_edit, 0, 1)
        eq_form.addWidget(eq_cat_label, 0, 2)
        eq_form.addWidget(self.equipment_category_combo, 0, 3)
        eq_form.addWidget(eq_total_label, 0, 4)
        eq_form.addWidget(self.equipment_total_spin, 0, 5)
        eq_form.addWidget(eq_avail_label, 0, 6)
        eq_form.addWidget(self.equipment_available_spin, 0, 7)

        self.eq_add_btn = QtWidgets.QPushButton("添加/更新设备")
        self.eq_add_btn.clicked.connect(self.add_or_update_equipment)
        self.eq_remove_btn = QtWidgets.QPushButton("删除设备")
        self.eq_remove_btn.clicked.connect(self.remove_equipment)

        eq_form.addWidget(self.eq_add_btn, 1, 6)
        eq_form.addWidget(self.eq_remove_btn, 1, 7)

        equipment_layout.addLayout(eq_form)
        self.equipment_id_edit.setEnabled(False)
        self.equipment_category_combo.setEnabled(False)
        self.equipment_total_spin.setEnabled(False)
        self.equipment_available_spin.setEnabled(False)
        self.eq_add_btn.setEnabled(False)
        self.eq_remove_btn.setEnabled(False)
        self.equipment_id_edit.setVisible(False)
        self.equipment_category_combo.setVisible(False)
        self.equipment_total_spin.setVisible(False)
        self.equipment_available_spin.setVisible(False)
        self.eq_add_btn.setVisible(False)
        self.eq_remove_btn.setVisible(False)
        eq_id_label.setVisible(False)
        eq_cat_label.setVisible(False)
        eq_total_label.setVisible(False)
        eq_avail_label.setVisible(False)
        eq_tip = QtWidgets.QLabel("设备维护请在管理员界面完成")
        eq_tip.setObjectName("metricLabel")
        equipment_layout.addWidget(eq_tip)

        mid.addWidget(equipment_group, 2)

        progress_group = QtWidgets.QGroupBox("进度与交期")
        progress_layout = QtWidgets.QVBoxLayout(progress_group)

        self.overall_progress = QtWidgets.QProgressBar()
        self.overall_progress.setValue(0)
        progress_layout.addWidget(QtWidgets.QLabel("订单总体进度"))
        progress_layout.addWidget(self.overall_progress)

        self.product_progress_table = QtWidgets.QTableWidget(0, 3)
        self.product_progress_table.setHorizontalHeaderLabels(["产品描述", "零件号", "进度"])
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

        eta_label = QtWidgets.QLabel("计划发货期")
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

        self.refresh_eta_btn = QtWidgets.QPushButton("刷新计划发货期")
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
        self.detail_eta_label = QtWidgets.QLabel("计划发货期: -")
        self.detail_progress_label = QtWidgets.QLabel("总体进度: 0%")
        self.order_summary_label.setObjectName("detailOrderTitle")
        self.detail_eta_label.setObjectName("metricValue")
        self.detail_progress_label.setObjectName("detailProgressBig")

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
        left_panel.setMinimumWidth(220)
        left_layout = QtWidgets.QVBoxLayout(left_panel)

        products_group = QtWidgets.QGroupBox("产品列表")
        products_layout = QtWidgets.QVBoxLayout(products_group)

        self.products_table = QtWidgets.QTableWidget(0, 7)
        self.products_table.setHorizontalHeaderLabels(
            ["产品描述", "零件号", "要求数量", "已生产", "单件重量(g)", "总重量(kg)", "数量进度"]
        )
        self.products_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.products_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.products_table.verticalHeader().setVisible(False)
        self.products_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.products_table.itemSelectionChanged.connect(self.on_product_select)
        self.products_table.cellChanged.connect(self.on_product_cell_changed)

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
        self.product_weight_spin = QtWidgets.QDoubleSpinBox()
        self.product_weight_spin.setRange(0, 1_000_000)
        self.product_weight_spin.setDecimals(2)
        self.product_weight_spin.setSingleStep(1.0)
        self.product_weight_spin.setMaximumWidth(140)

        prod_form.setColumnStretch(1, 3)
        prod_form.setColumnStretch(3, 2)
        prod_form.addWidget(QtWidgets.QLabel("产品描述"), 0, 0)
        prod_form.addWidget(self.product_id_edit, 0, 1)
        prod_form.addWidget(QtWidgets.QLabel("零件号"), 0, 2)
        prod_form.addWidget(self.product_part_edit, 0, 3)
        prod_form.addWidget(QtWidgets.QLabel("要求数量"), 0, 4)
        prod_form.addWidget(self.product_qty_spin, 0, 5)
        prod_form.addWidget(QtWidgets.QLabel("单件重量(g)"), 1, 0)
        prod_form.addWidget(self.product_weight_spin, 1, 1)

        self.product_add_btn = QtWidgets.QPushButton("添加产品")
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
        self.employee_name_edit.setVisible(False)
        self.employee_add_btn.setVisible(False)
        self.employee_remove_btn.setVisible(False)
        employee_tip = QtWidgets.QLabel("员工维护请在管理员界面完成")
        employee_tip.setObjectName("metricLabel")
        employees_layout.addWidget(employee_tip)
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
        self.phase_form_toggle_btn.setObjectName("phaseAction")
        self.phase_form_toggle_btn.clicked.connect(
            lambda: self._toggle_form_section(
                self.phase_form_toggle_btn,
                self.phase_form_widget,
                "工序",
            )
        )
        self.phase_hours_hint = QtWidgets.QLabel("工时按总量填写(小时)")
        self.phase_hours_hint.setObjectName("metricLabel")
        phases_toolbar.addWidget(self.phase_form_toggle_btn)
        phases_toolbar.addWidget(self.phase_hours_hint)
        phases_layout.addLayout(phases_toolbar)

        self.phases_table = QtWidgets.QTableWidget(0, 6)
        self.phases_table.setHorizontalHeaderLabels(
            ["工序名称", "总工时(小时)", "设备", "员工", "并行组", "进度"]
        )
        self.phases_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.phases_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.phases_table.verticalHeader().setVisible(False)
        self.phases_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.phases_table.itemSelectionChanged.connect(self.on_phase_select)
        self.phases_table.cellChanged.connect(self.on_phase_cell_changed)
        self.phases_table.setDragDropMode(QtWidgets.QAbstractItemView.DragDropMode.NoDragDrop)
        self.phases_table.cellDoubleClicked.connect(self.on_phase_cell_double_clicked)
        self.phases_table.setItemDelegateForColumn(
            0, ComboBoxDelegate(self._phase_template_name_list, self.phases_table)
        )
        self.phases_table.setItemDelegateForColumn(
            3, ComboBoxDelegate(self._employee_name_list, self.phases_table)
        )
        # 并行组显示 '-' 时不使用数字委托

        phases_layout.addWidget(self.phases_table)

        self.phase_form_widget = QtWidgets.QWidget()
        phase_form = QtWidgets.QGridLayout(self.phase_form_widget)
        self.phase_name_combo = QtWidgets.QComboBox()
        self.phase_name_combo.setEditable(False)
        self._configure_combo_popup(self.phase_name_combo, 200)
        self.phase_hours_spin = QtWidgets.QDoubleSpinBox()
        self.phase_hours_spin.setRange(0, 99999)
        self.phase_hours_spin.setDecimals(2)
        self.phase_equipment_display = QtWidgets.QLineEdit()
        self.phase_equipment_display.setReadOnly(True)
        self.phase_equipment_display.setText("无需设备")
        self.phase_equipment_btn = QtWidgets.QPushButton("选择设备")
        self.phase_equipment_btn.clicked.connect(self.open_phase_equipment_selector)
        self.phase_employee_combo = QtWidgets.QComboBox()
        self.phase_employee_combo.setEditable(True)
        self.phase_parallel_spin = QtWidgets.QSpinBox()
        self.phase_parallel_spin.setRange(0, 9999)
        self.phase_completed_spin = QtWidgets.QDoubleSpinBox()
        self.phase_completed_spin.setRange(0, 999999)
        self.phase_completed_spin.setDecimals(2)
        self.phase_completed_spin.setSuffix(" h")

        phase_form.addWidget(QtWidgets.QLabel("名称"), 0, 0)
        phase_form.addWidget(self.phase_name_combo, 0, 1)
        phase_form.addWidget(QtWidgets.QLabel("总工时(小时)"), 0, 2)
        phase_form.addWidget(self.phase_hours_spin, 0, 3)
        phase_form.addWidget(QtWidgets.QLabel("设备"), 0, 4)
        equipment_widget = QtWidgets.QWidget()
        equipment_layout = QtWidgets.QHBoxLayout(equipment_widget)
        equipment_layout.setContentsMargins(0, 0, 0, 0)
        equipment_layout.addWidget(self.phase_equipment_display, 1)
        equipment_layout.addWidget(self.phase_equipment_btn)
        phase_form.addWidget(equipment_widget, 0, 5)

        phase_form.addWidget(QtWidgets.QLabel("员工"), 1, 0)
        phase_form.addWidget(self.phase_employee_combo, 1, 1)
        phase_form.addWidget(QtWidgets.QLabel("并行组"), 1, 2)
        phase_form.addWidget(self.phase_parallel_spin, 1, 3)
        phase_form.addWidget(QtWidgets.QLabel("完成工时(小时)"), 1, 4)
        phase_form.addWidget(self.phase_completed_spin, 1, 5)

        self.phase_add_btn = QtWidgets.QPushButton("添加工序")
        self.phase_add_btn.setObjectName("phaseAction")
        self.phase_add_btn.clicked.connect(self.add_or_update_phase)
        self.phase_remove_btn = QtWidgets.QPushButton("删除工序")
        self.phase_remove_btn.clicked.connect(self.remove_phase)
        self.phase_parallel_btn = QtWidgets.QPushButton("设为并行组")
        self.phase_parallel_btn.clicked.connect(self.set_parallel_group)
        self.phase_parallel_clear_btn = QtWidgets.QPushButton("取消并行")
        self.phase_parallel_clear_btn.clicked.connect(self.clear_parallel_group)
        self.phase_update_btn = QtWidgets.QPushButton("更新")
        self.phase_update_btn.setObjectName("phaseAction")
        self.phase_update_btn.clicked.connect(self.update_phase_from_form)
        self.phase_move_up_btn = QtWidgets.QPushButton("上移")
        self.phase_move_up_btn.clicked.connect(lambda: self.move_phase(-1))
        self.phase_move_down_btn = QtWidgets.QPushButton("下移")
        self.phase_move_down_btn.clicked.connect(lambda: self.move_phase(1))

        phase_form.addWidget(self.phase_add_btn, 2, 4)
        phase_form.addWidget(self.phase_remove_btn, 2, 5)
        phase_form.addWidget(self.phase_parallel_btn, 3, 4)
        phase_form.addWidget(self.phase_parallel_clear_btn, 3, 5)
        phase_form.addWidget(self.phase_update_btn, 4, 4, 1, 2)
        phase_form.addWidget(self.phase_move_up_btn, 5, 4)
        phase_form.addWidget(self.phase_move_down_btn, 5, 5)

        phases_layout.addWidget(self.phase_form_widget)
        self.phase_form_widget.setVisible(False)
        right_splitter.addWidget(phases_group)

        brief_group = QtWidgets.QGroupBox("订单信息")
        brief_layout = QtWidgets.QGridLayout(brief_group)
        brief_layout.setHorizontalSpacing(14)
        brief_layout.setVerticalSpacing(6)

        self.brief_order_label = QtWidgets.QLabel("订单编号: -")
        self.brief_customer_label = QtWidgets.QLabel("客户: -")
        self.brief_ship_label = QtWidgets.QLabel("发货类别: -")
        self.brief_count_label = QtWidgets.QLabel("产品数: -")

        self.brief_order_date_label = QtWidgets.QLabel("订单日期: -")
        self.brief_start_date_label = QtWidgets.QLabel("开始日期: -")
        self.brief_due_date_label = QtWidgets.QLabel("要求发货日期: -")
        self.brief_plan_label = QtWidgets.QLabel("计划发货期: -")

        self.brief_event_count_label = QtWidgets.QLabel("事件数量: -")
        self.brief_defects_label = QtWidgets.QLabel("不合格品: -")

        self.detail_ops_btn = QtWidgets.QPushButton("事件/产能/不合格")
        self.detail_ops_btn.setObjectName("opsAction")
        self.detail_ops_btn.clicked.connect(self.open_ops_dialog)

        info_labels = (
            self.brief_order_label,
            self.brief_customer_label,
            self.brief_ship_label,
            self.brief_count_label,
            self.brief_order_date_label,
            self.brief_start_date_label,
            self.brief_due_date_label,
            self.brief_plan_label,
            self.brief_event_count_label,
            self.brief_defects_label,
        )
        for label in info_labels:
            label.setObjectName("metricLabel")

        brief_layout.addWidget(self.brief_order_label, 0, 0)
        brief_layout.addWidget(self.brief_customer_label, 0, 1)
        brief_layout.addWidget(self.brief_ship_label, 0, 2)
        brief_layout.addWidget(self.brief_count_label, 0, 3)

        brief_layout.addWidget(self.brief_order_date_label, 1, 0)
        brief_layout.addWidget(self.brief_start_date_label, 1, 1)
        brief_layout.addWidget(self.brief_due_date_label, 1, 2)
        brief_layout.addWidget(self.brief_plan_label, 1, 3)

        brief_layout.addWidget(self.brief_event_count_label, 2, 0)
        brief_layout.addWidget(self.brief_defects_label, 2, 1)

        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(self.detail_ops_btn)
        brief_layout.addLayout(button_row, 3, 0, 1, 4)
        brief_layout.setColumnStretch(0, 1)
        brief_layout.setColumnStretch(1, 1)
        brief_layout.setColumnStretch(2, 1)
        brief_layout.setColumnStretch(3, 1)
        right_splitter.addWidget(brief_group)

        right_layout.addWidget(right_splitter, 1)

        splitter.addWidget(right_panel)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)

        self._build_ops_dialog()

        return page

    def _build_ops_dialog(self) -> None:
        if hasattr(self, "ops_dialog") and self.ops_dialog:
            return
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("事件 / 产能 / 不合格")
        dialog.setMinimumSize(900, 600)

        layout = QtWidgets.QVBoxLayout(dialog)
        tabs = QtWidgets.QTabWidget()
        layout.addWidget(tabs)

        # Events tab
        events_page = QtWidgets.QWidget()
        events_page_layout = QtWidgets.QVBoxLayout(events_page)
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

        self.events_table = QtWidgets.QTableWidget(0, 4)
        self.events_table.setHorizontalHeaderLabels(["日期", "损失工时", "原因", "备注"])
        self.events_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.events_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.events_table.verticalHeader().setVisible(False)
        self.events_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.events_table.itemSelectionChanged.connect(self.on_event_select)
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
        self.event_remark_edit = QtWidgets.QLineEdit()
        self.event_remark_edit.setPlaceholderText("备注")

        ev_form.addWidget(QtWidgets.QLabel("日期"), 0, 0)
        ev_form.addWidget(self.event_date_edit, 0, 1)
        ev_form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        ev_form.addWidget(self.event_hours_spin, 0, 3)
        ev_form.addWidget(QtWidgets.QLabel("原因"), 0, 4)
        ev_form.addWidget(self.event_reason_combo, 0, 5)
        ev_form.addWidget(QtWidgets.QLabel("备注"), 0, 6)
        ev_form.addWidget(self.event_remark_edit, 0, 7)

        self.event_add_btn = QtWidgets.QPushButton("添加事件")
        self.event_add_btn.clicked.connect(self.add_event)

        ev_form.addWidget(self.event_add_btn, 1, 6)

        events_layout.addWidget(self.event_form_widget)
        self.event_form_widget.setVisible(False)
        event_tip = QtWidgets.QLabel("事件更新/删除请在管理员界面操作")
        event_tip.setObjectName("metricLabel")
        events_layout.addWidget(event_tip)
        events_page_layout.addWidget(events_group)
        tabs.addTab(events_page, "事件")

        # Capacity adjustments tab
        adjust_page = QtWidgets.QWidget()
        adjust_page_layout = QtWidgets.QVBoxLayout(adjust_page)
        adjust_group = QtWidgets.QGroupBox("日历/产能调整")
        adjust_layout = QtWidgets.QVBoxLayout(adjust_group)
        adjust_toolbar = QtWidgets.QHBoxLayout()
        adjust_toolbar.addStretch(1)
        self.adjust_form_toggle_btn = QtWidgets.QPushButton("编辑产能")
        self.adjust_form_toggle_btn.clicked.connect(
            lambda: self._toggle_form_section(
                self.adjust_form_toggle_btn,
                self.adjust_form_widget,
                "产能",
            )
        )
        adjust_toolbar.addWidget(self.adjust_form_toggle_btn)
        adjust_layout.addLayout(adjust_toolbar)

        self.adjustments_table = QtWidgets.QTableWidget(0, 4)
        self.adjustments_table.setHorizontalHeaderLabels(["日期", "加班工时", "设备", "说明"])
        self.adjustments_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.adjustments_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.adjustments_table.verticalHeader().setVisible(False)
        self.adjustments_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.adjustments_table.itemSelectionChanged.connect(self.on_adjustment_select)
        adjust_layout.addWidget(self.adjustments_table)

        self.adjust_form_widget = QtWidgets.QWidget()
        adj_form = QtWidgets.QGridLayout(self.adjust_form_widget)
        self.adjust_date_edit = QtWidgets.QDateEdit()
        self.adjust_date_edit.setDate(QtCore.QDate.currentDate())
        self._setup_date_edit(self.adjust_date_edit)
        self.adjust_hours_spin = QtWidgets.QDoubleSpinBox()
        self.adjust_hours_spin.setRange(0, 24)
        self.adjust_hours_spin.setDecimals(2)
        self.adjust_hours_spin.setValue(2.0)
        self.adjust_reason_edit = QtWidgets.QLineEdit()
        self.adjust_reason_edit.setPlaceholderText("加班说明(可选)")
        self.adjust_equipment_list = QtWidgets.QListWidget()
        self.adjust_equipment_list.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.MultiSelection
        )
        self.adjust_equipment_list.setMaximumHeight(90)

        adj_form.addWidget(QtWidgets.QLabel("日期"), 0, 0)
        adj_form.addWidget(self.adjust_date_edit, 0, 1)
        adj_form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        adj_form.addWidget(self.adjust_hours_spin, 0, 3)
        adj_form.addWidget(QtWidgets.QLabel("说明"), 0, 4)
        adj_form.addWidget(self.adjust_reason_edit, 0, 5)
        adj_form.addWidget(QtWidgets.QLabel("设备"), 1, 0)
        adj_form.addWidget(self.adjust_equipment_list, 1, 1, 1, 5)

        self.adjust_add_btn = QtWidgets.QPushButton("添加加班")
        self.adjust_add_btn.clicked.connect(self.add_capacity_adjustment)
        self.adjust_update_btn = QtWidgets.QPushButton("更新加班")
        self.adjust_update_btn.clicked.connect(self.update_capacity_adjustment)
        self.adjust_remove_btn = QtWidgets.QPushButton("删除加班")
        self.adjust_remove_btn.clicked.connect(self.remove_capacity_adjustment)

        adj_form.addWidget(self.adjust_add_btn, 2, 3)
        adj_form.addWidget(self.adjust_update_btn, 2, 4)
        adj_form.addWidget(self.adjust_remove_btn, 2, 5)

        adjust_layout.addWidget(self.adjust_form_widget)
        self.adjust_form_widget.setVisible(False)
        adjust_page_layout.addWidget(adjust_group)
        tabs.addTab(adjust_page, "产能调整")

        # Defects tab
        defects_page = QtWidgets.QWidget()
        defects_page_layout = QtWidgets.QVBoxLayout(defects_page)
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
        defects_page_layout.addWidget(defects_group)
        tabs.addTab(defects_page, "不合格品")

        self.ops_dialog = dialog

    def open_ops_dialog(self) -> None:
        if not hasattr(self, "ops_dialog") or not self.ops_dialog:
            self._build_ops_dialog()
        self.ops_dialog.show()
        self.ops_dialog.raise_()
        self.ops_dialog.activateWindow()

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
        self.admin_tabs.addTab(self._build_admin_log_tab(), "系统日志")
        self.admin_tabs.addTab(self._build_admin_users_tab(), "用户管理")
        self.admin_tabs.addTab(self._build_admin_reasons_tab(), "事件原因")
        self.admin_tabs.addTab(self._build_admin_defect_categories_tab(), "不合格原因")
        self.admin_tabs.addTab(self._build_admin_customer_codes_tab(), "客户代码")
        self.admin_tabs.addTab(self._build_admin_shipping_methods_tab(), "发货类别")
        self.admin_tabs.addTab(self._build_admin_equipment_categories_tab(), "设备分类")
        self.admin_tabs.addTab(self._build_admin_equipment_templates_tab(), "设备")
        self.admin_tabs.addTab(self._build_admin_phase_templates_tab(), "工序模板")
        self.admin_tabs.addTab(self._build_admin_employee_templates_tab(), "员工")
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

        self.visual_tabs = QtWidgets.QTabWidget()
        layout.addWidget(self.visual_tabs, 1)

        # Global overview tab
        global_tab = QtWidgets.QWidget()
        global_layout = QtWidgets.QVBoxLayout(global_tab)
        global_layout.setSpacing(12)

        global_summary = QtWidgets.QGroupBox("全局概览")
        global_summary_layout = QtWidgets.QGridLayout(global_summary)
        self.global_orders_label = QtWidgets.QLabel("订单数: 0")
        self.global_products_label = QtWidgets.QLabel("产品数: 0")
        self.global_employees_label = QtWidgets.QLabel("员工数: 0")
        self.global_events_label = QtWidgets.QLabel("事件数: 0")
        self.global_defects_label = QtWidgets.QLabel("不合格数量: 0")
        self.global_hours_total_label = QtWidgets.QLabel("总工时: 0h")
        self.global_hours_done_label = QtWidgets.QLabel("已完成: 0h")
        self.global_hours_remaining_label = QtWidgets.QLabel("剩余: 0h")
        self.global_shift_label = QtWidgets.QLabel("当前班次: -")
        for label in (
            self.global_orders_label,
            self.global_products_label,
            self.global_employees_label,
            self.global_events_label,
            self.global_defects_label,
            self.global_hours_total_label,
            self.global_hours_done_label,
            self.global_hours_remaining_label,
            self.global_shift_label,
        ):
            label.setObjectName("metricLabel")
        global_summary_layout.addWidget(self.global_orders_label, 0, 0)
        global_summary_layout.addWidget(self.global_products_label, 0, 1)
        global_summary_layout.addWidget(self.global_employees_label, 0, 2)
        global_summary_layout.addWidget(self.global_shift_label, 0, 3)
        global_summary_layout.addWidget(self.global_events_label, 1, 0)
        global_summary_layout.addWidget(self.global_defects_label, 1, 1)
        global_summary_layout.addWidget(self.global_hours_total_label, 1, 2)
        global_summary_layout.addWidget(self.global_hours_done_label, 1, 3)
        global_summary_layout.addWidget(self.global_hours_remaining_label, 2, 0)
        for col in range(4):
            global_summary_layout.setColumnStretch(col, 1)
        global_layout.addWidget(global_summary)

        global_orders_group = QtWidgets.QGroupBox("订单概览")
        global_orders_layout = QtWidgets.QVBoxLayout(global_orders_group)
        self.global_orders_table = QtWidgets.QTableWidget(0, 8)
        self.global_orders_table.setHorizontalHeaderLabels(
            ["订单", "客户", "发货类别", "完成进度", "总工时", "剩余工时", "产品数量进度", "计划发货期"]
        )
        self.global_orders_table.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.global_orders_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.global_orders_table.verticalHeader().setVisible(False)
        self.global_orders_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        global_orders_layout.addWidget(self.global_orders_table)
        view_row = QtWidgets.QHBoxLayout()
        view_row.addStretch(1)
        self.global_view_order_btn = QtWidgets.QPushButton("查看选中订单详情")
        self.global_view_order_btn.clicked.connect(self.open_selected_order_in_visuals)
        view_row.addWidget(self.global_view_order_btn)
        global_orders_layout.addLayout(view_row)
        global_layout.addWidget(global_orders_group, 1)
        self.visual_global_tab_index = self.visual_tabs.addTab(global_tab, "全局概览")

        # Order detail tab
        order_tab = QtWidgets.QWidget()
        order_layout = QtWidgets.QVBoxLayout(order_tab)
        order_layout.setSpacing(12)
        order_select_row = QtWidgets.QHBoxLayout()
        order_select_row.addWidget(QtWidgets.QLabel("订单"))
        self.visual_order_combo = QtWidgets.QComboBox()
        self.visual_order_combo.setMinimumWidth(220)
        self.visual_order_combo.currentTextChanged.connect(self.on_visual_order_select)
        order_select_row.addWidget(self.visual_order_combo)
        order_select_row.addStretch(1)
        order_layout.addLayout(order_select_row)

        summary_group = QtWidgets.QGroupBox("关键指标")
        summary_layout = QtWidgets.QGridLayout(summary_group)
        self.visual_order_label = QtWidgets.QLabel("订单: -")
        self.visual_shift_label = QtWidgets.QLabel("班次: -")
        self.visual_counts_label = QtWidgets.QLabel("产品/设备/员工: -/-/-")
        self.visual_eta_label = QtWidgets.QLabel("计划发货期: -")
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
        order_layout.addWidget(summary_group)

        upper = QtWidgets.QHBoxLayout()
        order_layout.addLayout(upper)

        progress_group = QtWidgets.QGroupBox("订单进度")
        progress_layout = QtWidgets.QVBoxLayout(progress_group)
        self.visual_progress_bar = QtWidgets.QProgressBar()
        progress_layout.addWidget(self.visual_progress_bar)
        self.visual_progress_table = QtWidgets.QTableWidget(0, 4)
        self.visual_progress_table.setHorizontalHeaderLabels(["产品", "要求数量", "已生产", "数量进度"])
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

        self.visual_order_tab_index = self.visual_tabs.addTab(order_tab, "订单详情")

        # Equipment load tab
        equipment_tab = QtWidgets.QWidget()
        equipment_layout = QtWidgets.QVBoxLayout(equipment_tab)
        equipment_layout.setSpacing(12)

        order_equipment_group = QtWidgets.QGroupBox("订单设备使用率")
        order_equipment_layout = QtWidgets.QVBoxLayout(order_equipment_group)
        self.visual_equipment_order_label = QtWidgets.QLabel("当前订单: -")
        self.visual_equipment_order_label.setObjectName("metricLabel")
        order_equipment_layout.addWidget(self.visual_equipment_order_label)
        self.visual_equipment_table = QtWidgets.QTableWidget(0, 5)
        self.visual_equipment_table.setHorizontalHeaderLabels(
            ["设备", "类别", "负载工时", "可用工时", "使用率"]
        )
        self.visual_equipment_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_equipment_table.verticalHeader().setVisible(False)
        self.visual_equipment_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        order_equipment_layout.addWidget(self.visual_equipment_table)
        equipment_layout.addWidget(order_equipment_group)

        load_group = QtWidgets.QGroupBox("综合设备负载")
        load_layout = QtWidgets.QVBoxLayout(load_group)
        self.visual_equipment_load_table = QtWidgets.QTableWidget(0, 4)
        self.visual_equipment_load_table.setHorizontalHeaderLabels(
            ["设备", "类别", "总负载工时", "关联订单数"]
        )
        self.visual_equipment_load_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_equipment_load_table.verticalHeader().setVisible(False)
        self.visual_equipment_load_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        load_layout.addWidget(self.visual_equipment_load_table)
        equipment_layout.addWidget(load_group)
        self.visual_equipment_tab_index = self.visual_tabs.addTab(equipment_tab, "设备")

        # Logs tab
        log_tab = QtWidgets.QWidget()
        log_layout = QtWidgets.QVBoxLayout(log_tab)
        self.visual_log_table = QtWidgets.QTableWidget(0, 3)
        self.visual_log_table.setHorizontalHeaderLabels(["日期", "用户", "内容"])
        self.visual_log_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.visual_log_table.verticalHeader().setVisible(False)
        self.visual_log_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        log_layout.addWidget(self.visual_log_table)
        self.visual_log_tab_index = self.visual_tabs.addTab(log_tab, "备忘录")

        return page

    def _build_admin_events_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_events_table = QtWidgets.QTableWidget(0, 4)
        self.admin_events_table.setHorizontalHeaderLabels(["日期", "损失工时", "原因", "备注"])
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
        self.admin_event_remark_edit = QtWidgets.QLineEdit()
        self.admin_event_remark_edit.setPlaceholderText("备注")

        form.addWidget(QtWidgets.QLabel("日期"), 0, 0)
        form.addWidget(self.admin_event_date_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("工时"), 0, 2)
        form.addWidget(self.admin_event_hours_spin, 0, 3)
        form.addWidget(QtWidgets.QLabel("原因"), 0, 4)
        form.addWidget(self.admin_event_reason_combo, 0, 5)
        form.addWidget(QtWidgets.QLabel("备注"), 0, 6)
        form.addWidget(self.admin_event_remark_edit, 0, 7)

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

        tip = QtWidgets.QLabel("系统日志仅管理员可查看")
        tip.setObjectName("metricLabel")
        layout.addWidget(tip)

        filter_row = QtWidgets.QHBoxLayout()
        self.admin_log_order_combo = QtWidgets.QComboBox()
        self.admin_log_user_combo = QtWidgets.QComboBox()
        self.admin_log_date_check = QtWidgets.QCheckBox("按日期筛选")
        self.admin_log_start_date = QtWidgets.QDateEdit()
        self._setup_date_edit(self.admin_log_start_date)
        self.admin_log_end_date = QtWidgets.QDateEdit()
        self._setup_date_edit(self.admin_log_end_date)
        self.admin_log_start_date.setDate(QtCore.QDate.currentDate().addMonths(-1))
        self.admin_log_end_date.setDate(QtCore.QDate.currentDate())
        self.admin_log_clear_btn = QtWidgets.QPushButton("清除筛选")

        filter_row.addWidget(QtWidgets.QLabel("订单"))
        filter_row.addWidget(self.admin_log_order_combo)
        filter_row.addWidget(QtWidgets.QLabel("用户"))
        filter_row.addWidget(self.admin_log_user_combo)
        filter_row.addWidget(self.admin_log_date_check)
        filter_row.addWidget(QtWidgets.QLabel("从"))
        filter_row.addWidget(self.admin_log_start_date)
        filter_row.addWidget(QtWidgets.QLabel("到"))
        filter_row.addWidget(self.admin_log_end_date)
        filter_row.addWidget(self.admin_log_clear_btn)
        filter_row.addStretch(1)
        layout.addLayout(filter_row)

        self.admin_log_table = QtWidgets.QTableWidget(0, 4)
        self.admin_log_table.setHorizontalHeaderLabels(["时间", "订单", "用户", "内容"])
        self.admin_log_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.admin_log_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.admin_log_table.verticalHeader().setVisible(False)
        self.admin_log_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.admin_log_table.setSortingEnabled(True)
        self.admin_log_order_combo.currentTextChanged.connect(self._refresh_admin_log_table)
        self.admin_log_user_combo.currentTextChanged.connect(self._refresh_admin_log_table)
        self.admin_log_date_check.toggled.connect(self._refresh_admin_log_table)
        self.admin_log_start_date.dateChanged.connect(self._refresh_admin_log_table)
        self.admin_log_end_date.dateChanged.connect(self._refresh_admin_log_table)
        self.admin_log_clear_btn.clicked.connect(self._reset_admin_log_filters)
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

    def _build_admin_customer_codes_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_customer_code_list = QtWidgets.QListWidget()
        self.admin_customer_code_list.itemSelectionChanged.connect(
            self.on_admin_customer_code_select
        )
        layout.addWidget(self.admin_customer_code_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_customer_code_edit = QtWidgets.QLineEdit()
        self.admin_customer_code_edit.setPlaceholderText("客户代码")
        self.admin_customer_code_add_btn = QtWidgets.QPushButton("添加")
        self.admin_customer_code_add_btn.clicked.connect(self.admin_add_customer_code)
        self.admin_customer_code_update_btn = QtWidgets.QPushButton("更新选中")
        self.admin_customer_code_update_btn.clicked.connect(self.admin_update_customer_code)
        self.admin_customer_code_remove_btn = QtWidgets.QPushButton("删除选中")
        self.admin_customer_code_remove_btn.clicked.connect(self.admin_remove_customer_code)

        form.addWidget(self.admin_customer_code_edit)
        form.addWidget(self.admin_customer_code_add_btn)
        form.addWidget(self.admin_customer_code_update_btn)
        form.addWidget(self.admin_customer_code_remove_btn)

        layout.addLayout(form)
        return page

    def _build_admin_shipping_methods_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        self.admin_shipping_method_list = QtWidgets.QListWidget()
        self.admin_shipping_method_list.itemSelectionChanged.connect(
            self.on_admin_shipping_method_select
        )
        layout.addWidget(self.admin_shipping_method_list)

        form = QtWidgets.QHBoxLayout()
        self.admin_shipping_method_edit = QtWidgets.QLineEdit()
        self.admin_shipping_method_edit.setPlaceholderText("发货类别")
        self.admin_shipping_method_add_btn = QtWidgets.QPushButton("添加")
        self.admin_shipping_method_add_btn.clicked.connect(self.admin_add_shipping_method)
        self.admin_shipping_method_update_btn = QtWidgets.QPushButton("更新选中")
        self.admin_shipping_method_update_btn.clicked.connect(self.admin_update_shipping_method)
        self.admin_shipping_method_remove_btn = QtWidgets.QPushButton("删除选中")
        self.admin_shipping_method_remove_btn.clicked.connect(self.admin_remove_shipping_method)

        form.addWidget(self.admin_shipping_method_edit)
        form.addWidget(self.admin_shipping_method_add_btn)
        form.addWidget(self.admin_shipping_method_update_btn)
        form.addWidget(self.admin_shipping_method_remove_btn)

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

        self.admin_equipment_template_table = QtWidgets.QTableWidget(0, 5)
        self.admin_equipment_template_table.setHorizontalHeaderLabels(
            ["设备编号", "类别", "总数量", "可用数量", "班次模板"]
        )
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
        self.admin_equipment_shift_combo = QtWidgets.QComboBox()
        form.addWidget(QtWidgets.QLabel("总数量"), 0, 4)
        form.addWidget(self.admin_equipment_total_spin, 0, 5)
        form.addWidget(QtWidgets.QLabel("可用数量"), 0, 6)
        form.addWidget(self.admin_equipment_available_spin, 0, 7)
        form.addWidget(QtWidgets.QLabel("班次模板"), 0, 8)
        form.addWidget(self.admin_equipment_shift_combo, 0, 9)

        self.admin_equipment_add_btn = QtWidgets.QPushButton("添加/更新设备")
        self.admin_equipment_add_btn.clicked.connect(self.admin_add_or_update_equipment_template)
        self.admin_equipment_remove_btn = QtWidgets.QPushButton("删除设备")
        self.admin_equipment_remove_btn.clicked.connect(self.admin_remove_equipment_template)
        self.admin_equipment_apply_btn = QtWidgets.QPushButton("同步到订单")
        self.admin_equipment_apply_btn.clicked.connect(self.admin_apply_equipment_template)

        form.addWidget(self.admin_equipment_add_btn, 1, 7)
        form.addWidget(self.admin_equipment_remove_btn, 1, 8)
        form.addWidget(self.admin_equipment_apply_btn, 1, 9)

        layout.addLayout(form)
        return page

    def _build_admin_phase_templates_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        template_row = QtWidgets.QHBoxLayout()
        template_row.addWidget(QtWidgets.QLabel("工序模版"))
        self.admin_phase_template_combo = QtWidgets.QComboBox()
        self.admin_phase_template_combo.setMinimumWidth(200)
        self.admin_phase_template_combo.currentTextChanged.connect(
            self.on_admin_phase_template_set_change
        )
        self.admin_phase_template_name_edit = QtWidgets.QLineEdit()
        self.admin_phase_template_name_edit.setPlaceholderText("模版名称")
        self.admin_phase_template_save_btn = QtWidgets.QPushButton("保存模版")
        self.admin_phase_template_save_btn.clicked.connect(self.admin_save_phase_template_set)
        self.admin_phase_template_delete_btn = QtWidgets.QPushButton("删除模版")
        self.admin_phase_template_delete_btn.clicked.connect(self.admin_delete_phase_template_set)
        template_row.addWidget(self.admin_phase_template_combo)
        template_row.addWidget(self.admin_phase_template_name_edit)
        template_row.addWidget(self.admin_phase_template_save_btn)
        template_row.addWidget(self.admin_phase_template_delete_btn)
        template_row.addStretch(1)
        layout.addLayout(template_row)

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
            ["工序名称", "总工时(小时)", "设备", "员工", "并行组"]
        )
        self.admin_phase_template_table.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.admin_phase_template_table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.admin_phase_template_table.verticalHeader().setVisible(False)
        self.admin_phase_template_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        self.admin_phase_template_table.itemSelectionChanged.connect(
            self.on_admin_phase_template_select
        )
        self.admin_phase_template_table.cellChanged.connect(
            self.on_admin_phase_template_cell_changed
        )
        self.admin_phase_template_table.cellDoubleClicked.connect(
            self.on_admin_phase_template_cell_double_clicked
        )
        self.admin_phase_template_table.setItemDelegateForColumn(
            3, ComboBoxDelegate(self._employee_template_list, self.admin_phase_template_table)
        )
        # 并行组显示 '-' 时不使用数字委托
        layout.addWidget(self.admin_phase_template_table)

        form = QtWidgets.QGridLayout()
        self.admin_phase_name_edit = QtWidgets.QLineEdit()
        self.admin_phase_hours_spin = QtWidgets.QDoubleSpinBox()
        self.admin_phase_hours_spin.setRange(0, 99999)
        self.admin_phase_hours_spin.setDecimals(2)
        self.admin_phase_equipment_display = QtWidgets.QLineEdit()
        self.admin_phase_equipment_display.setReadOnly(True)
        self.admin_phase_equipment_display.setText("无需设备")
        self.admin_phase_equipment_btn = QtWidgets.QPushButton("选择设备")
        self.admin_phase_equipment_btn.clicked.connect(self.open_admin_phase_equipment_selector)
        self.admin_phase_employee_combo = QtWidgets.QComboBox()
        self.admin_phase_employee_combo.setEditable(True)
        self.admin_phase_parallel_spin = QtWidgets.QSpinBox()
        self.admin_phase_parallel_spin.setRange(0, 9999)

        form.addWidget(QtWidgets.QLabel("名称"), 0, 0)
        form.addWidget(self.admin_phase_name_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("总工时(小时)"), 0, 2)
        form.addWidget(self.admin_phase_hours_spin, 0, 3)
        form.addWidget(QtWidgets.QLabel("设备"), 0, 4)
        admin_equipment_widget = QtWidgets.QWidget()
        admin_equipment_layout = QtWidgets.QHBoxLayout(admin_equipment_widget)
        admin_equipment_layout.setContentsMargins(0, 0, 0, 0)
        admin_equipment_layout.addWidget(self.admin_phase_equipment_display, 1)
        admin_equipment_layout.addWidget(self.admin_phase_equipment_btn)
        form.addWidget(admin_equipment_widget, 0, 5)

        form.addWidget(QtWidgets.QLabel("员工"), 1, 0)
        form.addWidget(self.admin_phase_employee_combo, 1, 1)
        form.addWidget(QtWidgets.QLabel("并行组"), 1, 2)
        form.addWidget(self.admin_phase_parallel_spin, 1, 3)

        self.admin_phase_add_btn = QtWidgets.QPushButton("添加工序")
        self.admin_phase_add_btn.clicked.connect(self.admin_add_or_update_phase_template)
        self.admin_phase_remove_btn = QtWidgets.QPushButton("删除工序")
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
        self.admin_employee_add_btn = QtWidgets.QPushButton("添加/更新员工")
        self.admin_employee_add_btn.clicked.connect(self.admin_add_or_update_employee_template)
        self.admin_employee_remove_btn = QtWidgets.QPushButton("删除员工")
        self.admin_employee_remove_btn.clicked.connect(self.admin_remove_employee_template)
        self.admin_employee_apply_btn = QtWidgets.QPushButton("同步到订单")
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
        order_date = date.today()
        due_date = order_date
        customer_code = ""
        shipping_method = self.shipping_methods[0] if self.shipping_methods else ""
        if any(o.order_id == order_id for o in self.orders):
            QtWidgets.QMessageBox.warning(self, "重复订单", "订单编号已存在，请使用其他编号。")
            return
        self.order = Order(
            order_id=order_id,
            start_dt=start_dt,
            order_date=order_date,
            customer_code=customer_code,
            shipping_method=shipping_method,
            due_date=due_date,
            products=[],
            events=[],
            defects=[],
            equipment=list(self.equipment),
            employees=list(self.employees),
        )
        self.orders.append(self.order)
        self.employees = list(self.order.employees)
        self._log_change(f"创建订单 {order_id} (开始日期: {start_dt.date().isoformat()})")
        self._refresh_all()
        self._auto_save()

    def open_order_editor(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先选择订单。")
            return

        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("编辑订单")
        dialog.setModal(True)

        layout = QtWidgets.QVBoxLayout(dialog)
        form = QtWidgets.QGridLayout()

        order_id_edit = QtWidgets.QLineEdit(self.order.order_id)
        start_date_edit = QtWidgets.QDateEdit()
        start_date_edit.setDate(_date_to_qdate(self.order.start_dt.date()))
        self._setup_date_edit(start_date_edit)

        order_date_edit = QtWidgets.QDateEdit()
        order_date_edit.setDate(_date_to_qdate(self.order.order_date))
        self._setup_date_edit(order_date_edit)

        due_date_edit = QtWidgets.QDateEdit()
        due_date_edit.setDate(_date_to_qdate(self.order.due_date))
        self._setup_date_edit(due_date_edit)

        customer_code_combo = QtWidgets.QComboBox()
        customer_code_combo.setEditable(False)
        codes = list(self.customer_codes)
        if self.order.customer_code and self.order.customer_code not in codes:
            codes.append(self.order.customer_code)
        customer_code_combo.addItems(codes)
        customer_code_combo.setCurrentText(self.order.customer_code)

        shipping_method_combo = QtWidgets.QComboBox()
        shipping_method_combo.setEditable(False)
        ship_methods = list(self.shipping_methods)
        if self.order.shipping_method and self.order.shipping_method not in ship_methods:
            ship_methods.append(self.order.shipping_method)
        shipping_method_combo.addItems(ship_methods)
        shipping_method_combo.setCurrentText(self.order.shipping_method)

        form.addWidget(QtWidgets.QLabel("订单编号"), 0, 0)
        form.addWidget(order_id_edit, 0, 1)
        form.addWidget(QtWidgets.QLabel("开始日期"), 1, 0)
        form.addWidget(start_date_edit, 1, 1)
        form.addWidget(QtWidgets.QLabel("订单日期"), 2, 0)
        form.addWidget(order_date_edit, 2, 1)
        form.addWidget(QtWidgets.QLabel("要求发货日期"), 3, 0)
        form.addWidget(due_date_edit, 3, 1)
        form.addWidget(QtWidgets.QLabel("客户代码"), 4, 0)
        form.addWidget(customer_code_combo, 4, 1)
        form.addWidget(QtWidgets.QLabel("发货类别"), 5, 0)
        form.addWidget(shipping_method_combo, 5, 1)

        layout.addLayout(form)

        btn_row = QtWidgets.QHBoxLayout()
        save_btn = QtWidgets.QPushButton("保存")
        cancel_btn = QtWidgets.QPushButton("取消")
        btn_row.addStretch(1)
        btn_row.addWidget(save_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        def on_save() -> None:
            new_order_id = order_id_edit.text().strip() or "O-UNKNOWN"
            if new_order_id != self.order.order_id and any(
                o.order_id == new_order_id for o in self.orders
            ):
                QtWidgets.QMessageBox.warning(self, "重复订单", "订单编号已存在，请使用其他编号。")
                return

            new_start_date = _qdate_to_date(start_date_edit.date())
            new_order_date = _qdate_to_date(order_date_edit.date())
            new_due_date = _qdate_to_date(due_date_edit.date())
            new_customer = customer_code_combo.currentText().strip()
            new_shipping_method = shipping_method_combo.currentText().strip()

            changes: List[str] = []
            if self.order.order_id != new_order_id:
                old_id = self.order.order_id
                self.order.order_id = new_order_id
                changes.append(f"订单编号: {old_id} -> {new_order_id}")
            if self.order.start_dt.date() != new_start_date:
                self.order.start_dt = datetime.combine(new_start_date, datetime.min.time())
                changes.append(f"开始日期: {new_start_date.isoformat()}")
            if self.order.order_date != new_order_date:
                self.order.order_date = new_order_date
                changes.append(f"订单日期: {new_order_date.isoformat()}")
            if self.order.due_date != new_due_date:
                self.order.due_date = new_due_date
                changes.append(f"要求发货日期: {new_due_date.isoformat()}")
            if self.order.customer_code != new_customer:
                self.order.customer_code = new_customer
                changes.append(f"客户代码: {new_customer or '-'}")
            if self.order.shipping_method != new_shipping_method:
                self.order.shipping_method = new_shipping_method
                changes.append(f"发货类别: {new_shipping_method or '-'}")

            if changes:
                self._log_change("更新订单信息 (" + ", ".join(changes) + ")")
            self._refresh_orders_table()
            self._refresh_order_summary()
            self.refresh_eta()
            self._auto_save()
            dialog.accept()

        save_btn.clicked.connect(on_save)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.exec()

    def save_factory(self) -> None:
        if not self.orders:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先创建或导入订单。")
            return
        if not self.factory_path:
            filename, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "保存工厂文件", "factory.json", "JSON Files (*.json)"
            )
            if not filename:
                return
            if not filename.endswith(".json"):
                filename += ".json"
            self._set_factory_path(filename)
        self._save_factory_to_path(self.factory_path, show_message=True)

    def load_factory(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "加载工厂文件", "", "JSON Files (*.json)"
        )
        if not filename:
            return
        try:
            with open(filename, "r", encoding="utf-8") as f:
                data = json.load(f)
            if "orders" not in data:
                QtWidgets.QMessageBox.critical(self, "加载失败", "文件不是工厂数据格式。")
                return
            self.factory_name = data.get("factory_name", self.factory_name)
            self.equipment = [
                Equipment(
                    equipment_id=e.get("equipment_id", ""),
                    category=e.get("category", ""),
                    total_count=int(e.get("total_count", 1)),
                    available_count=int(e.get("available_count", 1)),
                    shift_template_name=e.get("shift_template_name", ""),
                )
                for e in data.get("equipment", [])
            ]
            self.employees = list(data.get("employees", self.employees))
            self.orders = [self._order_from_dict(o) for o in data.get("orders", [])]
            self._sync_equipment_to_orders()
            self._sync_employees_to_order()
            self.app_logs = []
            for entry in data.get("app_logs", []):
                ts_raw = entry.get("timestamp", "")
                if ts_raw:
                    try:
                        ts = datetime.fromisoformat(ts_raw)
                    except Exception:
                        ts = datetime.now()
                else:
                    ts = datetime.now()
                self.app_logs.append(
                    LogEntry(
                        timestamp=ts,
                        user=entry.get("user", ""),
                        content=entry.get("content", ""),
                        order_id=entry.get("order_id", ""),
                    )
                )
            self.memos = []
            for entry in data.get("memos", []):
                day_raw = entry.get("day", "")
                if not day_raw:
                    continue
                try:
                    day_val = date.fromisoformat(day_raw)
                except Exception:
                    continue
                user = entry.get("user", "")
                content = entry.get("content", "")
                existing = next(
                    (memo for memo in self.memos if memo.day == day_val and memo.user == user),
                    None,
                )
                if existing:
                    if content:
                        if existing.content:
                            existing.content += "\n" + content
                        else:
                            existing.content = content
                else:
                    self.memos.append(
                        MemoEntry(
                            day=day_val,
                            user=user,
                            content=content,
                        )
                    )
            if not self.app_logs:
                for order in self.orders:
                    for log in order.logs:
                        if not log.order_id:
                            log.order_id = order.order_id
                        self.app_logs.append(log)
                    order.logs = []
            if not self.employees:
                employee_set = set()
                for order in self.orders:
                    employee_set.update(order.employees)
                self.employees = sorted(employee_set)
            active_id = data.get("active_order_id", "")
            self.order = next((o for o in self.orders if o.order_id == active_id), None)
            if not self.order and self.orders:
                self.order = self.orders[0]
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "加载失败", f"无法加载工厂文件: {exc}")
            return

        if self.order:
            self.order_id_edit.setText(self.order.order_id)
            self.start_date_edit.setDate(_date_to_qdate(self.order.start_dt.date()))
        else:
            self.order_id_edit.clear()
        self._maybe_import_templates_from_order(data)
        self._set_factory_path(filename)

        self._refresh_all()
        if self.order:
            self.go_to_detail()

    def duplicate_order(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先选择订单。")
            return
        base_id = self.order.order_id or "O-UNKNOWN"
        suffix = 1
        new_id = f"{base_id}-副本"
        while any(o.order_id == new_id for o in self.orders):
            suffix += 1
            new_id = f"{base_id}-副本{suffix}"
        new_order = self._order_from_dict(self._order_to_dict(self.order))
        new_order.order_id = new_id
        new_order.logs = []
        self.orders.append(new_order)
        self._sync_equipment_to_orders()
        self.order = new_order
        self.employees = list(new_order.employees)
        self._refresh_all()
        self._auto_save()

    def remove_order(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先选择订单。")
            return
        resp = QtWidgets.QMessageBox.question(
            self,
            "确认删除",
            f"确定删除订单 {self.order.order_id} 吗？此操作不可恢复。",
        )
        if resp != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        try:
            idx = self.orders.index(self.order)
        except ValueError:
            return
        del self.orders[idx]
        if self.orders:
            new_idx = min(idx, len(self.orders) - 1)
            self.order = self.orders[new_idx]
            self.employees = list(self.order.employees)
        else:
            self.order = None
            self.employees = []
        self._refresh_all()
        self._auto_save()

    def on_order_select(self) -> None:
        if self._updating_orders_table:
            return
        row = self.orders_table.currentRow()
        if row < 0:
            return
        item = self.orders_table.item(row, 0)
        if not item:
            return
        order_id = item.data(QtCore.Qt.ItemDataRole.UserRole) or item.text()
        order = next((o for o in self.orders if o.order_id == order_id), None)
        if not order:
            return
        self.order = order
        self.order_id_edit.setText(self.order.order_id)
        self.start_date_edit.setDate(_date_to_qdate(self.order.start_dt.date()))
        self._refresh_all()

    def on_order_double_click(self, item: QtWidgets.QTableWidgetItem) -> None:
        row = item.row() if item else self.orders_table.currentRow()
        if row < 0 or row >= len(self.orders):
            return
        self.orders_table.selectRow(row)
        self.on_order_select()
        self.go_to_detail()

    def go_to_detail(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或选择订单。")
            return
        self.stack.setCurrentWidget(self.detail_page)

    def go_to_visuals(self) -> None:
        if not self.orders:
            QtWidgets.QMessageBox.information(self, "提示", "请先创建或加载订单。")
            return
        if not self.order and self.orders:
            self.order = self.orders[0]
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

    def open_calendar_dialog(self) -> None:
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("日历")
        layout = QtWidgets.QVBoxLayout(dialog)
        calendar = QtWidgets.QCalendarWidget()
        calendar.setSelectedDate(QtCore.QDate.currentDate())
        layout.addWidget(calendar)
        close_btn = QtWidgets.QPushButton("关闭")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()

    def open_log_dialog(self) -> None:
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("备忘录")
        dialog.setMinimumSize(760, 420)
        layout = QtWidgets.QVBoxLayout(dialog)

        info_row = QtWidgets.QHBoxLayout()
        today = self._beijing_now().date()
        user = self.current_user or "系统"
        self.memo_day_label = QtWidgets.QLabel(f"日期: {today.isoformat()}")
        self.memo_user_label = QtWidgets.QLabel(f"用户: {user}")
        self.memo_day_label.setObjectName("metricLabel")
        self.memo_user_label.setObjectName("metricLabel")
        info_row.addWidget(self.memo_day_label)
        info_row.addWidget(self.memo_user_label)
        info_row.addStretch(1)
        layout.addLayout(info_row)

        form = QtWidgets.QHBoxLayout()
        self.log_content_edit = QtWidgets.QPlainTextEdit()
        self.log_content_edit.setPlaceholderText("输入备忘录内容...")
        self.log_content_edit.setMinimumHeight(80)
        add_btn = QtWidgets.QPushButton("保存备忘录")

        def on_add() -> None:
            text = self.log_content_edit.toPlainText().strip()
            if not text:
                return
            self._save_memo_entry(text)

        add_btn.clicked.connect(on_add)
        form.addWidget(self.log_content_edit, 1)
        form.addWidget(add_btn)
        layout.addLayout(form)

        self.log_table = QtWidgets.QTableWidget(0, 3)
        self.log_table.setHorizontalHeaderLabels(["日期", "用户", "内容"])
        self.log_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.log_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.log_table.verticalHeader().setVisible(False)
        self.log_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.log_table)

        entry = self._find_memo_entry(today, user)
        if entry:
            self.log_content_edit.setPlainText(entry.content)
        self._refresh_log_dialog_table()
        dialog.exec()

    def _refresh_log_dialog_table(self) -> None:
        if not hasattr(self, "log_table"):
            return
        self.log_table.setRowCount(0)
        memos = sorted(self.memos, key=lambda x: (x.day, x.user), reverse=True)
        for entry in memos:
            row = self.log_table.rowCount()
            self.log_table.insertRow(row)
            self.log_table.setItem(
                row, 0, QtWidgets.QTableWidgetItem(entry.day.isoformat())
            )
            self.log_table.setItem(row, 1, QtWidgets.QTableWidgetItem(entry.user))
            self.log_table.setItem(row, 2, QtWidgets.QTableWidgetItem(entry.content))

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
        user = self.current_user or "系统"
        order_id = self.order.order_id if self.order else ""
        self.app_logs.append(
            LogEntry(timestamp=self._beijing_now(), user=user, content=content, order_id=order_id)
        )
        self._refresh_admin_log_table()
        self._refresh_visual_logs()

    def _find_memo_entry(self, day: date, user: str) -> Optional[MemoEntry]:
        for entry in self.memos:
            if entry.day == day and entry.user == user:
                return entry
        return None

    def _save_memo_entry(self, content: str) -> None:
        day = self._beijing_now().date()
        user = self.current_user or "系统"
        entry = self._find_memo_entry(day, user)
        if entry:
            entry.content = content
        else:
            self.memos.append(MemoEntry(day=day, user=user, content=content))
        self._log_change(f"更新备忘录 {day.isoformat()} ({user})")
        self._refresh_log_dialog_table()
        self._auto_save()

    def _set_factory_path(self, filename: str) -> None:
        self.factory_path = filename
        if hasattr(self, "factory_path_label"):
            self.factory_path_label.setText(f"工厂文件: {filename}")
        self.statusBar().showMessage(f"工厂文件: {filename}", 5000)

    def show_factory_path(self) -> None:
        path = self.factory_path or "未设置"
        QtWidgets.QMessageBox.information(self, "工厂文件路径", path)

    def _save_factory_to_path(
        self, filename: str, show_message: bool = False, autosave: bool = False
    ) -> None:
        data = self._factory_to_dict()
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        if show_message:
            QtWidgets.QMessageBox.information(self, "保存成功", f"工厂已保存到 {filename}")
        if autosave:
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.statusBar().showMessage(f"自动保存: {os.path.basename(filename)} {timestamp}", 3000)

    def _auto_save(self) -> None:
        if not self.orders or not self.factory_path:
            return
        try:
            self._save_factory_to_path(self.factory_path, autosave=True)
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
        self.customer_codes = list(data.get("customer_codes", self.customer_codes))
        self.shipping_methods = list(data.get("shipping_methods", self.shipping_methods))
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
                shift_template_name=e.get("shift_template_name", ""),
            )
            for e in templates.get("equipment", [])
        ]
        self.phase_template_sets = {}
        phase_sets = templates.get("phase_templates", [])
        if isinstance(phase_sets, list) and phase_sets:
            for tpl in phase_sets:
                name = (tpl.get("name") or "默认模板").strip() or "默认模板"
                phases = [
                    Phase(
                        name=ph.get("name", ""),
                        planned_hours=float(ph.get("planned_hours", 0)),
                        parallel_group=int(ph.get("parallel_group", 0)),
                        equipment_id=ph.get("equipment_id", ""),
                        assigned_employee=ph.get("assigned_employee", ""),
                    )
                    for ph in tpl.get("phases", [])
                ]
                self.phase_template_sets[name] = phases
        if not self.phase_template_sets:
            legacy_phases = templates.get("phases", [])
            legacy_list = [
                Phase(
                    name=ph.get("name", ""),
                    planned_hours=float(ph.get("planned_hours", 0)),
                    parallel_group=int(ph.get("parallel_group", 0)),
                    equipment_id=ph.get("equipment_id", ""),
                    assigned_employee=ph.get("assigned_employee", ""),
                )
                for ph in legacy_phases
            ]
            self.phase_template_sets["默认模板"] = legacy_list
        active_phase_tpl = templates.get("active_phase_template", "").strip()
        if not active_phase_tpl or active_phase_tpl not in self.phase_template_sets:
            active_phase_tpl = next(iter(self.phase_template_sets.keys()), "默认模板")
        self.active_phase_template_name = active_phase_tpl
        self.phase_templates = list(self.phase_template_sets.get(active_phase_tpl, []))
        self.employees = list(templates.get("employees", self.employees))
        self.employee_templates = list(self.employees)
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
            "customer_codes": self.customer_codes,
            "shipping_methods": self.shipping_methods,
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
                        "shift_template_name": e.shift_template_name,
                    }
                    for e in self.equipment_templates
                ],
                "phase_templates": [
                    {
                        "name": name,
                        "phases": [
                            {
                                "name": ph.name,
                                "planned_hours": ph.planned_hours,
                                "parallel_group": ph.parallel_group,
                                "equipment_id": ph.equipment_id,
                                "assigned_employee": ph.assigned_employee,
                            }
                            for ph in phases
                        ],
                    }
                    for name, phases in self.phase_template_sets.items()
                ],
                "active_phase_template": self.active_phase_template_name,
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
                "employees": list(self.employees),
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
        if not self.customer_codes:
            self.customer_codes = ["HYD", "SWI", "SCH"]
        if not self.shipping_methods:
            self.shipping_methods = ["空运", "海运"]
        if not any(u.username == "admin" for u in self.user_accounts):
            self.user_accounts.append(UserAccount(username="admin", password=ADMIN_PASSWORD))
        if not self.phase_template_sets:
            base = list(self.phase_templates) if self.phase_templates else []
            self.phase_template_sets = {"默认模板": base}
        if not self.active_phase_template_name or self.active_phase_template_name not in self.phase_template_sets:
            self.active_phase_template_name = next(iter(self.phase_template_sets.keys()), "默认模板")
        self.phase_templates = list(self.phase_template_sets.get(self.active_phase_template_name, []))
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
        if not self.equipment and self.equipment_templates:
            self.equipment = [
                Equipment(
                    e.equipment_id,
                    e.category,
                    e.total_count,
                    e.available_count,
                    e.shift_template_name,
                )
                for e in self.equipment_templates
            ]
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

    def _shift_template_by_name(self, name: str) -> Optional[ShiftTemplate]:
        for tpl in self.shift_templates:
            if tpl.name == name:
                return tpl
        return None

    def _total_shift_hours_for_template(
        self, template: Optional[ShiftTemplate], start_day: date, end_day: date
    ) -> float:
        if not template or end_day < start_day:
            return 0.0
        total = 0.0
        current = start_day
        while current <= end_day:
            total += template.hours_for_weekday(current.weekday())
            current = current + timedelta(days=1)
        return total

    def _total_shift_hours(self, start_day: date, end_day: date) -> float:
        return self._total_shift_hours_for_template(self.cal.shift_template, start_day, end_day)

    def _refresh_visuals(self) -> None:
        self._refresh_visual_order_combo()
        self._refresh_visuals_global()
        self._refresh_visuals_order()
        self._refresh_visuals_equipment()
        self._refresh_visual_logs()

    def _refresh_visual_order_combo(self) -> None:
        if not hasattr(self, "visual_order_combo"):
            return
        current = self.visual_order_combo.currentText()
        self.visual_order_combo.blockSignals(True)
        self.visual_order_combo.clear()
        self.visual_order_combo.addItems([o.order_id for o in self.orders])
        if current and current in [o.order_id for o in self.orders]:
            self.visual_order_combo.setCurrentText(current)
        elif self.order:
            self.visual_order_combo.setCurrentText(self.order.order_id)
        self.visual_order_combo.blockSignals(False)

    def _refresh_visuals_global(self) -> None:
        if not hasattr(self, "global_orders_table"):
            return
        total_orders = len(self.orders)
        total_products = sum(len(o.products) for o in self.orders)
        total_events = sum(len(o.events) for o in self.orders)
        total_defects = sum(sum(d.count for d in o.defects) for o in self.orders)
        total_hours = 0.0
        done_hours = 0.0
        remaining_hours = 0.0
        for order in self.orders:
            equipment_map = _equipment_available_map(order)
            total = 0.0
            done = 0.0
            for product in order.products:
                for phase in product.phases:
                    hours = _phase_effective_hours(phase, product.quantity, equipment_map)
                    total += hours
                    ratio = _phase_completion_ratio(phase, product.quantity)
                    done += hours * ratio
            total_hours += total
            done_hours += done
            remaining_hours += max(total - done, 0.0)

        self.global_orders_label.setText(f"订单数: {total_orders}")
        self.global_products_label.setText(f"产品数: {total_products}")
        self.global_employees_label.setText(f"员工数: {len(self.employees)}")
        self.global_events_label.setText(f"事件数: {total_events}")
        self.global_defects_label.setText(f"不合格数量: {total_defects}")
        self.global_hours_total_label.setText(f"总工时: {total_hours:g}h")
        self.global_hours_done_label.setText(f"已完成: {done_hours:g}h")
        self.global_hours_remaining_label.setText(f"剩余: {remaining_hours:g}h")
        self.global_shift_label.setText(f"当前班次: {self.active_shift_template_name or '-'}")

        self.global_orders_table.setRowCount(0)
        for order in self.orders:
            equipment_map = _equipment_available_map(order)
            total = 0.0
            done = 0.0
            for product in order.products:
                for phase in product.phases:
                    hours = _phase_effective_hours(phase, product.quantity, equipment_map)
                    total += hours
                    ratio = _phase_completion_ratio(phase, product.quantity)
                    done += hours * ratio
            remaining = max(total - done, 0.0)
            progress = (done / total) if total > 0 else 0.0
            product_progress_parts: List[str] = []
            for product in order.products:
                qty = max(int(product.quantity), 0)
                produced = max(0, int(product.produced_qty))
                if qty > 0:
                    produced = min(produced, qty)
                    pct = produced / qty
                    product_progress_parts.append(
                        f"{product.product_id} {produced}/{qty}({pct:.0%})"
                    )
                else:
                    product_progress_parts.append(f"{product.product_id} -")
            product_progress_text = " | ".join(product_progress_parts) if product_progress_parts else "-"
            eta_text = "-"
            if self.cal.shift_template and order.products:
                try:
                    result = compute_eta(order, self.cal)
                    eta_dt: datetime = result["eta_dt"]
                    eta_text = eta_dt.strftime("%Y-%m-%d %H:%M")
                except Exception:
                    eta_text = "计算失败"
            row = self.global_orders_table.rowCount()
            self.global_orders_table.insertRow(row)
            order_item = QtWidgets.QTableWidgetItem(order.order_id)
            order_item.setData(QtCore.Qt.ItemDataRole.UserRole, order.order_id)
            self.global_orders_table.setItem(row, 0, order_item)
            self.global_orders_table.setItem(row, 1, QtWidgets.QTableWidgetItem(order.customer_code or "-"))
            self.global_orders_table.setItem(row, 2, QtWidgets.QTableWidgetItem(order.shipping_method or "-"))
            self.global_orders_table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{progress:.0%}"))
            self.global_orders_table.setItem(row, 4, QtWidgets.QTableWidgetItem(f"{total:g}h"))
            self.global_orders_table.setItem(row, 5, QtWidgets.QTableWidgetItem(f"{remaining:g}h"))
            self.global_orders_table.setItem(row, 6, QtWidgets.QTableWidgetItem(product_progress_text))
            self.global_orders_table.setItem(row, 7, QtWidgets.QTableWidgetItem(eta_text))

    def _refresh_visuals_order(self) -> None:
        if not hasattr(self, "visual_progress_bar"):
            return
        if not self.order:
            self.visual_progress_bar.setValue(0)
            self.visual_progress_table.setRowCount(0)
            self.visual_defect_table.setRowCount(0)
            self.visual_event_table.setRowCount(0)
            self.visual_equipment_table.setRowCount(0)
            self.visual_order_label.setText("订单: -")
            self.visual_shift_label.setText("班次: -")
            self.visual_counts_label.setText("产品/设备/员工: -/-/-")
            self.visual_eta_label.setText("计划发货期: -")
            self.visual_total_hours_label.setText("总工时: -")
            self.visual_done_hours_label.setText("已完成工时: -")
            self.visual_remaining_hours_label.setText("剩余工时: -")
            self.visual_lost_hours_label.setText("累计损失: -")
            self.visual_defect_total_label.setText("不合格数量: -")
            self.visual_equipment_avg_label.setText("平均设备使用率: -")
            self.visual_equipment_peak_label.setText("最高设备使用率: -")
            if hasattr(self, "visual_equipment_order_label"):
                self.visual_equipment_order_label.setText("当前订单: -")
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
        if hasattr(self, "visual_equipment_order_label"):
            self.visual_equipment_order_label.setText(f"当前订单: {self.order.order_id}")
        self.visual_shift_label.setText(f"班次: {self.active_shift_template_name or '-'}")
        self.visual_counts_label.setText(
            f"产品/设备/员工: {len(self.order.products)}/{len(self.order.equipment)}/{len(self.employees)}"
        )
        self.visual_total_hours_label.setText(f"总工时: {total:g}h")
        self.visual_done_hours_label.setText(f"已完成工时: {done:g}h")
        self.visual_remaining_hours_label.setText(f"剩余工时: {remaining:g}h")

        self.visual_progress_table.setRowCount(0)
        for product in self.order.products:
            row = self.visual_progress_table.rowCount()
            self.visual_progress_table.insertRow(row)
            p_progress = _product_quantity_progress(product)
            produced = max(0, int(product.produced_qty))
            if produced > product.quantity:
                produced = product.quantity
            self.visual_progress_table.setItem(row, 0, QtWidgets.QTableWidgetItem(product.product_id))
            self.visual_progress_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(product.quantity)))
            self.visual_progress_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(produced)))
            self.visual_progress_table.setItem(row, 3, QtWidgets.QTableWidgetItem(f"{p_progress:.0%}"))

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

        capacity_map = self.last_capacity_map
        if not self.last_eta_dt:
            try:
                result = compute_eta(self.order, self.cal)
                self.last_eta_dt = result["eta_dt"]
                self.last_remaining_hours = result["remaining_hours"]
                capacity_map = result.get("daily_capacity_map", {})
                self.last_capacity_map = dict(capacity_map)
            except Exception:
                self.last_eta_dt = None
                self.last_capacity_map = {}

        start_day = self.order.start_dt.date()
        end_day = self.last_eta_dt.date() if self.last_eta_dt else None
        if self.last_eta_dt:
            self.visual_eta_label.setText(
                f"计划发货期: {self.last_eta_dt.strftime('%Y-%m-%d %H:%M')}"
            )
        else:
            self.visual_eta_label.setText("计划发货期: -")

        equipment_loads: Dict[str, float] = {}
        for product in self.order.products:
            for phase in product.phases:
                eq_ids = _split_equipment_ids(phase.equipment_id)
                if not eq_ids:
                    continue
                hours = phase.planned_hours
                share = hours / max(len(eq_ids), 1)
                for eq_id in eq_ids:
                    equipment_loads[eq_id] = equipment_loads.get(eq_id, 0.0) + share

        self.visual_equipment_table.setRowCount(0)
        usage_values: List[float] = []
        peak_usage_raw = -1.0
        peak_usage_capped = 0.0
        peak_equipment = "-"
        for eq in self.order.equipment:
            load = equipment_loads.get(eq.equipment_id, 0.0)
            eq_template = (
                self._shift_template_by_name(eq.shift_template_name)
                if eq.shift_template_name
                else self.cal.shift_template
            )
            total_shift_hours = (
                self._total_shift_hours_for_template(eq_template, start_day, end_day)
                if end_day
                else 0.0
            )
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

    def _refresh_visuals_equipment(self) -> None:
        if not hasattr(self, "visual_equipment_load_table"):
            return
        self.visual_equipment_load_table.setRowCount(0)
        equipment_loads: Dict[str, float] = {}
        equipment_orders: Dict[str, set] = {}
        for order in self.orders:
            for product in order.products:
                for phase in product.phases:
                    eq_ids = _split_equipment_ids(phase.equipment_id)
                    if not eq_ids:
                        continue
                    share = phase.planned_hours / max(len(eq_ids), 1)
                    for eq_id in eq_ids:
                        equipment_loads[eq_id] = equipment_loads.get(eq_id, 0.0) + share
                        equipment_orders.setdefault(eq_id, set()).add(order.order_id)

        eq_map: Dict[str, Equipment] = {e.equipment_id: e for e in self.equipment}
        for eq_id, load in sorted(equipment_loads.items()):
            row = self.visual_equipment_load_table.rowCount()
            self.visual_equipment_load_table.insertRow(row)
            eq = eq_map.get(eq_id)
            self.visual_equipment_load_table.setItem(row, 0, QtWidgets.QTableWidgetItem(eq_id))
            self.visual_equipment_load_table.setItem(
                row, 1, QtWidgets.QTableWidgetItem(eq.category if eq else "-")
            )
            self.visual_equipment_load_table.setItem(row, 2, QtWidgets.QTableWidgetItem(f"{load:g}"))
            order_count = len(equipment_orders.get(eq_id, set()))
            self.visual_equipment_load_table.setItem(
                row, 3, QtWidgets.QTableWidgetItem(str(order_count))
            )

    def _refresh_visual_logs(self) -> None:
        if not hasattr(self, "visual_log_table"):
            return
        self.visual_log_table.setRowCount(0)
        memos = sorted(self.memos, key=lambda x: (x.day, x.user), reverse=True)
        for entry in memos:
            row = self.visual_log_table.rowCount()
            self.visual_log_table.insertRow(row)
            self.visual_log_table.setItem(
                row, 0, QtWidgets.QTableWidgetItem(entry.day.isoformat())
            )
            self.visual_log_table.setItem(row, 1, QtWidgets.QTableWidgetItem(entry.user))
            self.visual_log_table.setItem(row, 2, QtWidgets.QTableWidgetItem(entry.content))

    def open_selected_order_in_visuals(self) -> None:
        if not hasattr(self, "global_orders_table"):
            return
        row = self.global_orders_table.currentRow()
        if row < 0:
            return
        item = self.global_orders_table.item(row, 0)
        if not item:
            return
        order_id = item.text()
        self.set_active_order_by_id(order_id)
        if hasattr(self, "visual_tabs") and hasattr(self, "visual_order_tab_index"):
            self.visual_tabs.setCurrentIndex(self.visual_order_tab_index)

    def on_visual_order_select(self, order_id: str) -> None:
        if not order_id:
            return
        self.set_active_order_by_id(order_id)

    def set_active_order_by_id(self, order_id: str) -> None:
        order = next((o for o in self.orders if o.order_id == order_id), None)
        if not order:
            return
        self.order = order
        self.order_id_edit.setText(order.order_id)
        self.start_date_edit.setDate(_date_to_qdate(order.start_dt.date()))
        self._refresh_all()

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
                shift_template_name=e.get("shift_template_name", ""),
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
        if self.phase_templates:
            self.phase_template_sets = {"默认模板": self._clone_phase_list(self.phase_templates)}
            self.active_phase_template_name = "默认模板"
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
        self._refresh_admin_customer_codes()
        self._refresh_admin_shipping_methods()
        self._refresh_admin_equipment_categories()
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_templates_table()
        self._refresh_admin_phase_template_sets()
        self._refresh_admin_employee_templates_list()
        self._refresh_admin_shift_templates_list()
        self._refresh_admin_equipment_shift_combo()
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
            self.equipment.append(Equipment(eq_id, category, total, available, ""))
        if eq_id:
            action = "更新" if existing else "添加"
            self._log_change(
                f"{action}设备 {eq_id} (类别: {category or '-'}, 总: {total}, 可用: {available})"
            )
        self._sync_equipment_to_orders()
        self._sync_equipment_templates_from_system()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self.refresh_eta()
        self._refresh_orders_table()
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
        self._sync_equipment_to_orders()
        self._sync_equipment_templates_from_system()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self._refresh_defect_detail_combos()
        self.refresh_eta()
        self._refresh_orders_table()
        self._auto_save()

    def _sync_equipment_to_orders(self) -> None:
        for order in self.orders:
            order.equipment = [
                Equipment(
                    e.equipment_id,
                    e.category,
                    e.total_count,
                    e.available_count,
                    e.shift_template_name,
                )
                for e in self.equipment
            ]

    def _sync_equipment_templates_from_system(self) -> None:
        self.equipment_templates = [
            Equipment(
                e.equipment_id,
                e.category,
                e.total_count,
                e.available_count,
                e.shift_template_name,
            )
            for e in self.equipment
        ]
        if hasattr(self, "admin_equipment_template_table"):
            self._refresh_admin_equipment_templates_table()
        self._save_app_templates()

    def _refresh_equipment_table(self) -> None:
        self.equipment_table.setRowCount(0)
        for eq in self.equipment:
            row = self.equipment_table.rowCount()
            self.equipment_table.insertRow(row)
            self.equipment_table.setItem(row, 0, QtWidgets.QTableWidgetItem(eq.equipment_id))
            self.equipment_table.setItem(row, 1, QtWidgets.QTableWidgetItem(eq.category or "-"))
            self.equipment_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(eq.total_count)))
            self.equipment_table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(eq.available_count)))
        self._refresh_adjustment_equipment_list()

    def _refresh_orders_table(self) -> None:
        if not hasattr(self, "orders_table"):
            return
        self._updating_orders_table = True
        self.orders_table.blockSignals(True)
        self.orders_table.setSortingEnabled(False)
        self.orders_table.setRowCount(0)
        for order in self.orders:
            row = self.orders_table.rowCount()
            self.orders_table.insertRow(row)
            order_item = QtWidgets.QTableWidgetItem(order.order_id)
            order_item.setData(QtCore.Qt.ItemDataRole.UserRole, order.order_id)
            self.orders_table.setItem(row, 0, order_item)
            self.orders_table.setItem(row, 1, QtWidgets.QTableWidgetItem(order.customer_code or "-"))
            self.orders_table.setItem(row, 2, QtWidgets.QTableWidgetItem(order.start_dt.date().isoformat()))
            due_text = order.due_date.isoformat() if order.due_date else "-"
            self.orders_table.setItem(row, 3, QtWidgets.QTableWidgetItem(due_text))

            eta_text = "-"
            remaining_text = "-"
            if self.cal.shift_template and order.products:
                try:
                    result = compute_eta(order, self.cal)
                    eta_dt: datetime = result["eta_dt"]
                    remaining_hours: float = result["remaining_hours"]
                    eta_text = eta_dt.strftime("%Y-%m-%d %H:%M")
                    remaining_text = f"{remaining_hours:g}h"
                except Exception:
                    eta_text = "计算失败"
                    remaining_text = "-"

            self.orders_table.setItem(row, 4, QtWidgets.QTableWidgetItem(eta_text))
            self.orders_table.setItem(row, 5, QtWidgets.QTableWidgetItem(remaining_text))
            self.orders_table.setItem(row, 6, QtWidgets.QTableWidgetItem(str(len(order.products))))

        if self.order:
            for row in range(self.orders_table.rowCount()):
                item = self.orders_table.item(row, 0)
                if item and item.data(QtCore.Qt.ItemDataRole.UserRole) == self.order.order_id:
                    self.orders_table.selectRow(row)
                    break
        self.orders_table.blockSignals(False)
        self.orders_table.setSortingEnabled(True)
        self._updating_orders_table = False

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
        for order in self.orders:
            order.employees = list(self.employees)

    def _refresh_employee_list(self) -> None:
        self.employee_list.clear()
        self.employee_list.addItems(self.employees)

    def _refresh_phase_employee_combo(self) -> None:
        current = self.phase_employee_combo.currentText()
        self.phase_employee_combo.clear()
        items = [""] + [name for name in self.employees if name]
        self.phase_employee_combo.addItems(items)
        if current in items:
            self.phase_employee_combo.setCurrentText(current)
        else:
            self.phase_employee_combo.setCurrentIndex(0)

    def _refresh_phase_name_combo(self) -> None:
        if not hasattr(self, "phase_name_combo"):
            return
        current = self.phase_name_combo.currentText()
        items = self._phase_template_name_list()
        self.phase_name_combo.clear()
        self.phase_name_combo.addItems(items)
        if current in items:
            self.phase_name_combo.setCurrentText(current)
        elif items:
            self.phase_name_combo.setCurrentIndex(0)

    def _refresh_phase_equipment_combo(self) -> None:
        if not hasattr(self, "phase_equipment_display"):
            return
        current_ids = _split_equipment_ids(self.phase_equipment_display.text())
        available_ids = {e.equipment_id for e in self.equipment if e.equipment_id}
        filtered = [eq_id for eq_id in current_ids if eq_id in available_ids]
        self._set_phase_form_equipment_ids(filtered)

    def _set_phase_form_equipment_ids(self, equipment_ids: List[str]) -> None:
        if not hasattr(self, "phase_equipment_display"):
            return
        normalized = _normalize_equipment_ids(equipment_ids)
        if normalized:
            self.phase_equipment_display.setText(", ".join(normalized))
        else:
            self.phase_equipment_display.setText("无需设备")

    def _get_phase_form_equipment_ids(self) -> List[str]:
        if not hasattr(self, "phase_equipment_display"):
            return []
        text = self.phase_equipment_display.text().strip()
        return _split_equipment_ids(text)

    def _choose_equipment_ids(self, current_ids: List[str]) -> List[str]:
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("选择设备")
        dialog.setMinimumSize(360, 420)
        layout = QtWidgets.QVBoxLayout(dialog)
        list_widget = QtWidgets.QListWidget()
        list_widget.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.MultiSelection
        )
        current_set = set(current_ids)
        for eq in self.equipment:
            if not eq.equipment_id:
                continue
            item = QtWidgets.QListWidgetItem(eq.equipment_id)
            if eq.equipment_id in current_set:
                item.setSelected(True)
            list_widget.addItem(item)
        layout.addWidget(list_widget, 1)
        if list_widget.count() == 0:
            QtWidgets.QMessageBox.information(self, "提示", "请先在管理员界面添加设备。")
            return []

        btn_row = QtWidgets.QHBoxLayout()
        clear_btn = QtWidgets.QPushButton("清空")
        ok_btn = QtWidgets.QPushButton("确定")
        cancel_btn = QtWidgets.QPushButton("取消")
        btn_row.addWidget(clear_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        def on_clear() -> None:
            list_widget.clearSelection()

        clear_btn.clicked.connect(on_clear)
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)

        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return current_ids
        selected = [item.text() for item in list_widget.selectedItems()]
        return _normalize_equipment_ids(selected)

    def open_phase_equipment_selector(self) -> None:
        current_ids = self._get_phase_form_equipment_ids()
        selected = self._choose_equipment_ids(current_ids)
        self._set_phase_form_equipment_ids(selected)

    def _set_admin_phase_form_equipment_ids(self, equipment_ids: List[str]) -> None:
        if not hasattr(self, "admin_phase_equipment_display"):
            return
        normalized = _normalize_equipment_ids(equipment_ids)
        if normalized:
            self.admin_phase_equipment_display.setText(", ".join(normalized))
        else:
            self.admin_phase_equipment_display.setText("无需设备")

    def _get_admin_phase_form_equipment_ids(self) -> List[str]:
        if not hasattr(self, "admin_phase_equipment_display"):
            return []
        text = self.admin_phase_equipment_display.text().strip()
        return _split_equipment_ids(text)

    def open_admin_phase_equipment_selector(self) -> None:
        current_ids = self._get_admin_phase_form_equipment_ids()
        selected = self._choose_equipment_ids(current_ids)
        self._set_admin_phase_form_equipment_ids(selected)

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
        self.product_weight_spin.setValue(product.unit_weight_g)
        self._refresh_phase_table(product)

    def on_product_cell_changed(self, row: int, col: int) -> None:
        if self._updating_products_table or not self.order:
            return
        if row < 0 or row >= len(self.order.products):
            return
        product = self.order.products[row]
        item = self.products_table.item(row, col)
        if not item:
            return
        text = item.text().strip()
        if col == 0:
            if not text:
                self.products_table.blockSignals(True)
                item.setText(product.product_id)
                self.products_table.blockSignals(False)
                return
            if any(p is not product and p.product_id == text for p in self.order.products):
                QtWidgets.QMessageBox.information(self, "提示", "产品描述已存在。")
                self.products_table.blockSignals(True)
                item.setText(product.product_id)
                self.products_table.blockSignals(False)
                return
            if text != product.product_id:
                old_id = product.product_id
                product.product_id = text
                self.active_product_id = text
                if self.order:
                    for defect in self.order.defects:
                        if defect.product_id == old_id:
                            defect.product_id = text
                self._log_change(f"更新产品描述 {old_id} -> {text}")
                self._refresh_defect_product_combo()
                self._refresh_admin_product_combo()
                self._refresh_products_table()
                self._auto_save()
        elif col == 1:
            if text != product.part_number:
                old_part = product.part_number
                product.part_number = text
                self._log_change(
                    f"更新零件号 {product.product_id} {old_part or '-'} -> {text or '-'}"
                )
                self._refresh_products_table()
                self._auto_save()
        elif col == 2:
            try:
                qty = int(text)
            except ValueError:
                qty = product.quantity
                self.products_table.blockSignals(True)
                item.setText(str(product.quantity))
                self.products_table.blockSignals(False)
                return
            if qty < 1:
                qty = 1
                self.products_table.blockSignals(True)
                item.setText(str(qty))
                self.products_table.blockSignals(False)
            if qty != product.quantity:
                old_qty = product.quantity
                product.quantity = qty
                if product.produced_qty > qty:
                    product.produced_qty = qty
                self._log_change(
                    f"更新产品数量 {product.product_id} {old_qty} -> {qty}"
                )
                self.refresh_eta()
                self._refresh_products_table()
                self._auto_save()
        elif col == 3:
            try:
                produced = int(text)
            except ValueError:
                produced = product.produced_qty
                self.products_table.blockSignals(True)
                item.setText(str(product.produced_qty))
                self.products_table.blockSignals(False)
                return
            if produced < 0:
                produced = 0
                self.products_table.blockSignals(True)
                item.setText("0")
                self.products_table.blockSignals(False)
            if produced > product.quantity:
                produced = product.quantity
                self.products_table.blockSignals(True)
                item.setText(str(produced))
                self.products_table.blockSignals(False)
            if produced != product.produced_qty:
                old_val = product.produced_qty
                product.produced_qty = produced
                self._log_change(
                    f"更新已生产数量 {product.product_id} {old_val} -> {produced}"
                )
                self._refresh_products_table()
                self._refresh_visuals()
                self._auto_save()
        elif col == 4:
            try:
                weight = float(text)
            except ValueError:
                weight = product.unit_weight_g
                self.products_table.blockSignals(True)
                item.setText(f"{product.unit_weight_g:g}")
                self.products_table.blockSignals(False)
                return
            if weight < 0:
                weight = 0.0
                self.products_table.blockSignals(True)
                item.setText("0")
                self.products_table.blockSignals(False)
            if abs(weight - product.unit_weight_g) > 1e-6:
                old_weight = product.unit_weight_g
                product.unit_weight_g = weight
                self._log_change(
                    f"更新产品重量 {product.product_id} {old_weight:g}g -> {weight:g}g"
                )
                self._refresh_products_table()
                self._auto_save()

    def add_or_update_product(self) -> None:
        if not self.order:
            QtWidgets.QMessageBox.warning(self, "无订单", "请先创建或加载订单。")
            return
        product_id = self.product_id_edit.text().strip()
        if not product_id:
            QtWidgets.QMessageBox.warning(self, "无效输入", "产品描述不能为空。")
            return
        if any(p.product_id == product_id for p in self.order.products):
            QtWidgets.QMessageBox.information(self, "提示", "产品描述已存在，请使用其他名称。")
            return
        part_number = self.product_part_edit.text().strip()
        quantity = int(self.product_qty_spin.value())
        unit_weight_g = float(self.product_weight_spin.value())

        self.order.products.append(
            Product(
                product_id=product_id,
                part_number=part_number,
                quantity=quantity,
                produced_qty=0,
                unit_weight_g=unit_weight_g,
            )
        )
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
        self._updating_products_table = True
        self.products_table.blockSignals(True)
        self.products_table.setRowCount(0)
        if not self.order:
            self.active_product_id = ""
            self._update_phase_product_label(None)
            self.products_table.blockSignals(False)
            self._updating_products_table = False
            return
        current_id = self.active_product_id
        current_row = self.products_table.currentRow()
        if 0 <= current_row < len(self.order.products):
            current_id = self.order.products[current_row].product_id
        for product in self.order.products:
            row = self.products_table.rowCount()
            self.products_table.insertRow(row)
            progress = _product_quantity_progress(product)
            id_item = QtWidgets.QTableWidgetItem(product.product_id)
            id_item.setFlags(id_item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            part_item = QtWidgets.QTableWidgetItem(product.part_number)
            part_item.setFlags(part_item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            qty_item = QtWidgets.QTableWidgetItem(str(product.quantity))
            qty_item.setFlags(qty_item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            produced_item = QtWidgets.QTableWidgetItem(str(product.produced_qty))
            produced_item.setFlags(produced_item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            weight_item = QtWidgets.QTableWidgetItem(f"{float(product.unit_weight_g):g}")
            weight_item.setFlags(weight_item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 0, id_item)
            self.products_table.setItem(row, 1, part_item)
            self.products_table.setItem(row, 2, qty_item)
            self.products_table.setItem(row, 3, produced_item)
            unit_weight_g = max(0.0, float(product.unit_weight_g))
            total_weight_kg = unit_weight_g * max(product.quantity, 1) / 1000.0
            total_weight_item = QtWidgets.QTableWidgetItem(f"{total_weight_kg:g}")
            total_weight_item.setFlags(total_weight_item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            progress_item = QtWidgets.QTableWidgetItem(f"{progress:.0%}")
            progress_item.setFlags(progress_item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 4, weight_item)
            self.products_table.setItem(row, 5, total_weight_item)
            self.products_table.setItem(row, 6, progress_item)
        if not current_id and self.order.products:
            current_id = self.order.products[0].product_id
        if current_id:
            self.active_product_id = current_id
            self._select_product_by_id(current_id)
        else:
            self._update_phase_product_label(None)
        self.products_table.blockSignals(False)
        self._updating_products_table = False

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
        if self.phase_name_combo.findText(phase.name) < 0:
            self.phase_name_combo.addItem(phase.name)
        self.phase_name_combo.setCurrentText(phase.name)
        self.phase_hours_spin.setValue(phase.planned_hours)
        self._set_phase_form_equipment_ids(_split_equipment_ids(phase.equipment_id))
        self.phase_employee_combo.setCurrentText(phase.assigned_employee)
        self.phase_parallel_spin.setValue(phase.parallel_group)
        self.phase_completed_spin.setValue(phase.completed_hours)

    def on_phase_cell_double_clicked(self, row: int, col: int) -> None:
        if col != 2:
            return
        product = self._current_product()
        if not product or row < 0 or row >= len(product.phases):
            return
        phase = product.phases[row]
        current_ids = _split_equipment_ids(phase.equipment_id)
        selected = self._choose_equipment_ids(current_ids)
        new_text = _format_equipment_ids(selected)
        if new_text != phase.equipment_id:
            old_text = phase.equipment_id
            phase.equipment_id = new_text
            self._log_change(
                f"更新设备 {phase.name} {old_text or '-'} -> {new_text or '-'} (产品: {product.product_id})"
            )
            if row == self.phases_table.currentRow():
                self._set_phase_form_equipment_ids(selected)
            self._refresh_phase_table(product)
            self.refresh_eta()
            self._auto_save()

    def on_phase_cell_changed(self, row: int, col: int) -> None:
        if self._updating_phase_table:
            return
        product = self._current_product()
        if not product or row < 0 or row >= len(product.phases):
            return
        phase = product.phases[row]
        item = self.phases_table.item(row, col)
        if not item:
            return
        text = item.text().strip()
        changed = False
        if col == 0:
            if not text:
                item.setText(phase.name)
                return
            template_names = self._phase_template_name_list()
            if template_names and text not in template_names:
                QtWidgets.QMessageBox.information(self, "提示", "工序只能从管理员设置的模板中选择。")
                self.phases_table.blockSignals(True)
                item.setText(phase.name)
                self.phases_table.blockSignals(False)
                return
            if any(p is not phase and p.name == text for p in product.phases):
                QtWidgets.QMessageBox.information(self, "提示", "同一产品的工序不能重复。")
                self.phases_table.blockSignals(True)
                item.setText(phase.name)
                self.phases_table.blockSignals(False)
                return
            if text != phase.name:
                old = phase.name
                phase.name = text
                self._log_change(
                    f"更新工序名称 {old} -> {text} (产品: {product.product_id})"
                )
                changed = True
        elif col == 1:
            try:
                hours = float(text)
            except ValueError:
                self.phases_table.blockSignals(True)
                item.setText(f"{phase.planned_hours:g}")
                self.phases_table.blockSignals(False)
                return
            if hours < 0:
                hours = 0.0
                self.phases_table.blockSignals(True)
                item.setText("0")
                self.phases_table.blockSignals(False)
            if abs(hours - phase.planned_hours) > 1e-6:
                old = phase.planned_hours
                phase.planned_hours = hours
                if phase.completed_hours > hours:
                    phase.completed_hours = hours
                self._log_change(
                    f"更新工时 {phase.name} {old:g}h -> {hours:g}h (产品: {product.product_id})"
                )
                changed = True
        elif col == 2:
            if text in ("无需设备", "-"):
                new_eq = ""
            else:
                ids = _split_equipment_ids(text)
                if ids:
                    available = {e.equipment_id for e in self.equipment if e.equipment_id}
                    invalid = [eq_id for eq_id in ids if eq_id not in available]
                    if invalid:
                        QtWidgets.QMessageBox.information(
                            self, "提示", f"设备不存在: {', '.join(invalid)}"
                        )
                        self.phases_table.blockSignals(True)
                        item.setText(", ".join(_split_equipment_ids(phase.equipment_id)) or "-")
                        self.phases_table.blockSignals(False)
                        return
                    new_eq = _format_equipment_ids(ids)
                else:
                    new_eq = ""
            if new_eq != phase.equipment_id:
                old = phase.equipment_id
                phase.equipment_id = new_eq
                if new_eq == "":
                    self.phases_table.blockSignals(True)
                    item.setText("-")
                    self.phases_table.blockSignals(False)
                self._log_change(
                    f"更新设备 {phase.name} {old or '-'} -> {new_eq or '-'} (产品: {product.product_id})"
                )
                changed = True
        elif col == 3:
            if text != phase.assigned_employee:
                old = phase.assigned_employee
                phase.assigned_employee = text
                self._log_change(
                    f"更新员工 {phase.name} {old or '-'} -> {text or '-'} (产品: {product.product_id})"
                )
                changed = True
        elif col == 4:
            try:
                group = int(text)
            except ValueError:
                self.phases_table.blockSignals(True)
                item.setText(str(phase.parallel_group))
                self.phases_table.blockSignals(False)
                return
            if group < 0:
                group = 0
                self.phases_table.blockSignals(True)
                item.setText("-")
                self.phases_table.blockSignals(False)
            if group == 0 and text != "-":
                self.phases_table.blockSignals(True)
                item.setText("-")
                self.phases_table.blockSignals(False)
            if group != phase.parallel_group:
                old = phase.parallel_group
                phase.parallel_group = group
                if group == 0:
                    self.phases_table.blockSignals(True)
                    item.setText("-")
                    self.phases_table.blockSignals(False)
                self._log_change(
                    f"更新并行组 {phase.name} {old} -> {group} (产品: {product.product_id})"
                )
                changed = True

        if changed:
            self._refresh_phase_table(product)
            self.refresh_eta()
            self._auto_save()

    def on_phase_rows_moved(self, *args) -> None:
        if self._updating_phase_table:
            return
        product = self._current_product()
        if not product:
            return
        name_to_phase = {ph.name: ph for ph in product.phases}
        new_order: List[Phase] = []
        for r in range(self.phases_table.rowCount()):
            item = self.phases_table.item(r, 0)
            name = item.text().strip() if item else ""
            phase_obj = name_to_phase.get(name)
            if phase_obj:
                new_order.append(phase_obj)
        if len(new_order) != len(product.phases):
            self._refresh_phase_table(product)
            return
        product.phases = new_order
        self._log_change(f"调整工序顺序 (产品: {product.product_id})")
        self._auto_save()

    def move_phase(self, offset: int) -> None:
        product = self._current_product()
        if not product:
            return
        row = self.phases_table.currentRow()
        if row < 0 or row >= len(product.phases):
            return
        new_row = row + offset
        if new_row < 0 or new_row >= len(product.phases):
            return
        product.phases[row], product.phases[new_row] = product.phases[new_row], product.phases[row]
        self._log_change(
            f"调整工序顺序 {row + 1} -> {new_row + 1} (产品: {product.product_id})"
        )
        self._refresh_phase_table(product)
        self.phases_table.selectRow(new_row)
        self._auto_save()

    def add_or_update_phase(self) -> None:
        product = self._current_product()
        if not product:
            QtWidgets.QMessageBox.warning(self, "未选择产品", "请先选择一个产品。")
            return
        template_names = self._phase_template_name_list()
        if not template_names:
            QtWidgets.QMessageBox.information(self, "提示", "请先在管理员界面添加工序模板。")
            return
        name = self.phase_name_combo.currentText().strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "无效输入", "工序名称不能为空。")
            return
        if name not in template_names:
            QtWidgets.QMessageBox.information(self, "提示", "工序只能从管理员设置的模板中选择。")
            return
        if any(p.name == name for p in product.phases):
            QtWidgets.QMessageBox.information(self, "提示", "同一产品的工序不能重复。")
            return
        hours = float(self.phase_hours_spin.value())
        equipment_ids = self._get_phase_form_equipment_ids()
        equipment_id = _format_equipment_ids(equipment_ids)
        employee = self.phase_employee_combo.currentText().strip()
        parallel = int(self.phase_parallel_spin.value())
        completed_hours = float(self.phase_completed_spin.value())
        planned_total = max(0.0, hours)
        if planned_total > 0:
            completed_hours = min(max(completed_hours, 0.0), planned_total)
        else:
            completed_hours = 0.0

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
        self._set_phase_form_equipment_ids([])
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

    def update_phase_from_form(self) -> None:
        product = self._current_product()
        if not product:
            return
        row = self.phases_table.currentRow()
        if row < 0 or row >= len(product.phases):
            return
        phase = product.phases[row]
        template_names = self._phase_template_name_list()
        name = self.phase_name_combo.currentText().strip() or phase.name
        if template_names and name not in template_names:
            QtWidgets.QMessageBox.information(self, "提示", "工序只能从管理员设置的模板中选择。")
            return
        if name != phase.name and any(p is not phase and p.name == name for p in product.phases):
            QtWidgets.QMessageBox.information(self, "提示", "同一产品的工序不能重复。")
            return
        hours = float(self.phase_hours_spin.value())
        equipment_ids = self._get_phase_form_equipment_ids()
        equipment_id = _format_equipment_ids(equipment_ids)
        employee = self.phase_employee_combo.currentText().strip()
        parallel = int(self.phase_parallel_spin.value())
        completed_hours = float(self.phase_completed_spin.value())
        planned_total = max(0.0, hours)
        if planned_total > 0:
            completed_hours = min(max(completed_hours, 0.0), planned_total)
        else:
            completed_hours = 0.0

        changes = []
        if phase.name != name:
            changes.append(f"名称: {phase.name} -> {name}")
            phase.name = name
        if abs(phase.planned_hours - hours) > 1e-6:
            changes.append(f"工时: {phase.planned_hours:g}h -> {hours:g}h")
            phase.planned_hours = hours
        if phase.equipment_id != equipment_id:
            changes.append(f"设备: {phase.equipment_id or '-'} -> {equipment_id or '-'}")
            phase.equipment_id = equipment_id
        if phase.assigned_employee != employee:
            changes.append(f"员工: {phase.assigned_employee or '-'} -> {employee or '-'}")
            phase.assigned_employee = employee
        if phase.parallel_group != parallel:
            changes.append(f"并行组: {phase.parallel_group} -> {parallel}")
            phase.parallel_group = parallel
        if abs(phase.completed_hours - completed_hours) > 1e-6:
            changes.append(
                f"完成工时: {phase.completed_hours:g}h -> {completed_hours:g}h"
            )
            phase.completed_hours = completed_hours
        if changes:
            self._log_change(
                f"更新工序 {phase.name} ({' | '.join(changes)}) (产品: {product.product_id})"
            )
            self._refresh_phase_table(product)
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
        self._updating_phase_table = True
        self.phases_table.blockSignals(True)
        self.phases_table.setRowCount(0)
        for phase in product.phases:
            row = self.phases_table.rowCount()
            self.phases_table.insertRow(row)
            name_item = QtWidgets.QTableWidgetItem(phase.name)
            name_item.setData(QtCore.Qt.ItemDataRole.UserRole, phase)
            name_item.setData(QtCore.Qt.ItemDataRole.EditRole, phase.name)
            hours_item = QtWidgets.QTableWidgetItem(f"{phase.planned_hours:g}")
            equipment_ids = _split_equipment_ids(phase.equipment_id)
            equipment_text = ", ".join(equipment_ids) if equipment_ids else "-"
            equipment_item = QtWidgets.QTableWidgetItem(equipment_text)
            employee_item = QtWidgets.QTableWidgetItem(phase.assigned_employee)
            group_text = "-" if phase.parallel_group == 0 else str(phase.parallel_group)
            group_item = QtWidgets.QTableWidgetItem(group_text)
            group_item.setData(QtCore.Qt.ItemDataRole.EditRole, str(phase.parallel_group))
            for item in (name_item, hours_item, employee_item, group_item):
                item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            equipment_item.setFlags(equipment_item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            self.phases_table.setItem(row, 0, name_item)
            self.phases_table.setItem(row, 1, hours_item)
            self.phases_table.setItem(row, 2, equipment_item)
            self.phases_table.setItem(row, 3, employee_item)
            self.phases_table.setItem(row, 4, group_item)
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
        self.phases_table.blockSignals(False)
        self._updating_phase_table = False

    # ------------------------
    # Events
    # ------------------------

    def add_event(self) -> None:
        if not self.order:
            return
        day = _qdate_to_date(self.event_date_edit.date())
        hours = float(self.event_hours_spin.value())
        reason = self.event_reason_combo.currentText().strip() or "事件"
        remark = self.event_remark_edit.text().strip()
        self.order.events.append(Event(day=day, hours_lost=hours, reason=reason, remark=remark))
        self._log_change(
            f"添加事件 {day.isoformat()} {hours:g}h {reason} {remark or '-'}"
        )
        self._refresh_events_table()
        self._refresh_admin_events_table()
        self.refresh_eta()
        self._auto_save()

    def update_event(self) -> None:
        if not self.order:
            return
        row = self.events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        old_day = ev.day
        old_hours = ev.hours_lost
        old_reason = ev.reason
        old_remark = ev.remark
        ev.day = _qdate_to_date(self.event_date_edit.date())
        ev.hours_lost = float(self.event_hours_spin.value())
        ev.reason = self.event_reason_combo.currentText().strip() or "事件"
        ev.remark = self.event_remark_edit.text().strip()
        self._log_change(
            f"更新事件 {old_day.isoformat()} {old_hours:g}h {old_reason} {old_remark or '-'} -> "
            f"{ev.day.isoformat()} {ev.hours_lost:g}h {ev.reason} {ev.remark or '-'}"
        )
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

    def on_event_select(self) -> None:
        if not self.order:
            return
        row = self.events_table.currentRow()
        if row < 0 or row >= len(self.order.events):
            return
        ev = self.order.events[row]
        self.event_date_edit.setDate(_date_to_qdate(ev.day))
        self.event_hours_spin.setValue(ev.hours_lost)
        self.event_reason_combo.setCurrentText(ev.reason)
        self.event_remark_edit.setText(ev.remark)

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
            self.events_table.setItem(row, 3, QtWidgets.QTableWidgetItem(ev.remark))

    # ------------------------
    # Capacity adjustments
    # ------------------------

    def add_capacity_adjustment(self) -> None:
        if not self.order:
            return
        day = _qdate_to_date(self.adjust_date_edit.date())
        hours = float(self.adjust_hours_spin.value())
        if hours <= 0:
            QtWidgets.QMessageBox.information(self, "提示", "加班工时必须大于 0。")
            return
        reason = self.adjust_reason_edit.text().strip()
        equipment_ids = self._selected_adjustment_equipment_ids()
        if not equipment_ids:
            QtWidgets.QMessageBox.information(self, "提示", "请选择需要加班的设备。")
            return
        equipment_ids = sorted(set(equipment_ids))
        existing = next(
            (
                a
                for a in self.order.adjustments
                if a.day == day and set(a.equipment_ids) == set(equipment_ids)
            ),
            None,
        )
        if existing:
            existing.extra_hours += hours
            if reason:
                if existing.reason:
                    if reason not in existing.reason:
                        existing.reason = f"{existing.reason}; {reason}"
                else:
                    existing.reason = reason
            existing.equipment_ids = list(equipment_ids)
        else:
            self.order.adjustments.append(
                CapacityAdjustment(
                    day=day,
                    extra_hours=hours,
                    reason=reason,
                    equipment_ids=list(equipment_ids),
                )
            )
        eq_text = ",".join(equipment_ids)
        self._log_change(
            f"添加加班 {day.isoformat()} {hours:g}h 设备:{eq_text} {reason or '-'}"
        )
        self._refresh_adjustments_table()
        self.refresh_eta()
        self._auto_save()

    def update_capacity_adjustment(self) -> None:
        if not self.order:
            return
        row = self.adjustments_table.currentRow()
        if row < 0 or row >= len(self.order.adjustments):
            return
        adj = self.order.adjustments[row]
        day = _qdate_to_date(self.adjust_date_edit.date())
        hours = float(self.adjust_hours_spin.value())
        if hours <= 0:
            QtWidgets.QMessageBox.information(self, "提示", "加班工时必须大于 0。")
            return
        equipment_ids = self._selected_adjustment_equipment_ids()
        if not equipment_ids:
            QtWidgets.QMessageBox.information(self, "提示", "请选择需要加班的设备。")
            return
        equipment_ids = sorted(set(equipment_ids))
        reason = self.adjust_reason_edit.text().strip()
        old_text = f"{adj.day.isoformat()} {adj.extra_hours:g}h 设备:{','.join(adj.equipment_ids) or '-'}"
        adj.day = day
        adj.extra_hours = hours
        adj.reason = reason
        adj.equipment_ids = list(equipment_ids)
        self._log_change(
            f"更新加班 {old_text} -> {day.isoformat()} {hours:g}h 设备:{','.join(equipment_ids)}"
        )
        self._refresh_adjustments_table()
        self.refresh_eta()
        self._auto_save()

    def remove_capacity_adjustment(self) -> None:
        if not self.order:
            return
        row = self.adjustments_table.currentRow()
        if row < 0 or row >= len(self.order.adjustments):
            return
        adj = self.order.adjustments[row]
        eq_text = ",".join(adj.equipment_ids) if adj.equipment_ids else "-"
        self._log_change(
            f"删除加班 {adj.day.isoformat()} {adj.extra_hours:g}h 设备:{eq_text} {adj.reason or '-'}"
        )
        del self.order.adjustments[row]
        self._refresh_adjustments_table()
        self.refresh_eta()
        self._auto_save()

    def on_adjustment_select(self) -> None:
        if not self.order:
            return
        row = self.adjustments_table.currentRow()
        if row < 0 or row >= len(self.order.adjustments):
            return
        adj = self.order.adjustments[row]
        self.adjust_date_edit.setDate(_date_to_qdate(adj.day))
        self.adjust_hours_spin.setValue(adj.extra_hours)
        self.adjust_reason_edit.setText(adj.reason)
        self._set_adjustment_equipment_checks(adj.equipment_ids)

    def _refresh_adjustments_table(self) -> None:
        self.adjustments_table.setRowCount(0)
        if not self.order:
            return
        for adj in self.order.adjustments:
            row = self.adjustments_table.rowCount()
            self.adjustments_table.insertRow(row)
            self.adjustments_table.setItem(row, 0, QtWidgets.QTableWidgetItem(adj.day.isoformat()))
            self.adjustments_table.setItem(row, 1, QtWidgets.QTableWidgetItem(f"{adj.extra_hours:g}"))
            eq_text = ",".join(adj.equipment_ids) if adj.equipment_ids else "-"
            self.adjustments_table.setItem(row, 2, QtWidgets.QTableWidgetItem(eq_text))
            self.adjustments_table.setItem(row, 3, QtWidgets.QTableWidgetItem(adj.reason))

    def _selected_adjustment_equipment_ids(self) -> List[str]:
        if not hasattr(self, "adjust_equipment_list"):
            return []
        ids: List[str] = []
        for i in range(self.adjust_equipment_list.count()):
            item = self.adjust_equipment_list.item(i)
            if item.checkState() == QtCore.Qt.CheckState.Checked:
                ids.append(item.text())
        return ids

    def _set_adjustment_equipment_checks(self, equipment_ids: List[str]) -> None:
        if not hasattr(self, "adjust_equipment_list"):
            return
        selected = set(equipment_ids or [])
        for i in range(self.adjust_equipment_list.count()):
            item = self.adjust_equipment_list.item(i)
            state = (
                QtCore.Qt.CheckState.Checked
                if item.text() in selected
                else QtCore.Qt.CheckState.Unchecked
            )
            item.setCheckState(state)

    def _refresh_adjustment_equipment_list(self) -> None:
        if not hasattr(self, "adjust_equipment_list"):
            return
        selected = set(self._selected_adjustment_equipment_ids())
        self.adjust_equipment_list.clear()
        for eq in self.equipment:
            if not eq.equipment_id:
                continue
            item = QtWidgets.QListWidgetItem(eq.equipment_id)
            item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                QtCore.Qt.CheckState.Checked
                if eq.equipment_id in selected
                else QtCore.Qt.CheckState.Unchecked
            )
            self.adjust_equipment_list.addItem(item)

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
            self.admin_events_table.setItem(row, 3, QtWidgets.QTableWidgetItem(ev.remark))

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
        self.admin_event_remark_edit.setText(ev.remark)

    def admin_add_event(self) -> None:
        if not self.order:
            return
        day = _qdate_to_date(self.admin_event_date_edit.date())
        hours = float(self.admin_event_hours_spin.value())
        reason = self.admin_event_reason_combo.currentText().strip() or "事件"
        remark = self.admin_event_remark_edit.text().strip()
        self.order.events.append(Event(day=day, hours_lost=hours, reason=reason, remark=remark))
        self._log_change(
            f"添加事件 {day.isoformat()} {hours:g}h {reason} {remark or '-'}"
        )
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
        old_remark = ev.remark
        ev.day = _qdate_to_date(self.admin_event_date_edit.date())
        ev.hours_lost = float(self.admin_event_hours_spin.value())
        ev.reason = self.admin_event_reason_combo.currentText().strip() or "事件"
        ev.remark = self.admin_event_remark_edit.text().strip()
        self._log_change(
            f"更新事件 {old_day.isoformat()} {old_hours:g}h {old_reason} {old_remark or '-'} -> "
            f"{ev.day.isoformat()} {ev.hours_lost:g}h {ev.reason} {ev.remark or '-'}"
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
        self._refresh_admin_log_filters()
        order_filter = (
            self.admin_log_order_combo.currentText()
            if hasattr(self, "admin_log_order_combo")
            else "全部订单"
        )
        user_filter = (
            self.admin_log_user_combo.currentText()
            if hasattr(self, "admin_log_user_combo")
            else "全部用户"
        )
        use_date = (
            self.admin_log_date_check.isChecked()
            if hasattr(self, "admin_log_date_check")
            else False
        )
        start_date = (
            _qdate_to_date(self.admin_log_start_date.date())
            if hasattr(self, "admin_log_start_date")
            else None
        )
        end_date = (
            _qdate_to_date(self.admin_log_end_date.date())
            if hasattr(self, "admin_log_end_date")
            else None
        )
        logs = list(self.app_logs)
        logs.sort(key=lambda x: x.timestamp, reverse=True)

        self.admin_log_table.blockSignals(True)
        self.admin_log_table.setRowCount(0)
        for entry in logs:
            order_id = entry.order_id or "无订单"
            if order_filter not in ("全部订单", "") and order_id != order_filter:
                continue
            if user_filter not in ("全部用户", "") and entry.user != user_filter:
                continue
            if use_date and start_date and end_date:
                entry_date = entry.timestamp.date()
                if entry_date < start_date or entry_date > end_date:
                    continue
            row = self.admin_log_table.rowCount()
            self.admin_log_table.insertRow(row)
            self.admin_log_table.setItem(
                row, 0, QtWidgets.QTableWidgetItem(entry.timestamp.strftime("%Y-%m-%d %H:%M:%S"))
            )
            self.admin_log_table.setItem(row, 1, QtWidgets.QTableWidgetItem(order_id))
            self.admin_log_table.setItem(row, 2, QtWidgets.QTableWidgetItem(entry.user))
            self.admin_log_table.setItem(row, 3, QtWidgets.QTableWidgetItem(entry.content))
        self.admin_log_table.blockSignals(False)

    def _refresh_admin_log_filters(self) -> None:
        if not hasattr(self, "admin_log_order_combo") or not hasattr(
            self, "admin_log_user_combo"
        ):
            return
        order_current = self.admin_log_order_combo.currentText()
        user_current = self.admin_log_user_combo.currentText()
        order_items = ["全部订单"]
        order_ids = {entry.order_id for entry in self.app_logs if entry.order_id}
        order_items.extend(sorted(order_ids))
        if any(not entry.order_id for entry in self.app_logs):
            order_items.append("无订单")
        self.admin_log_order_combo.blockSignals(True)
        self.admin_log_order_combo.clear()
        self.admin_log_order_combo.addItems(order_items)
        if order_current in order_items:
            self.admin_log_order_combo.setCurrentText(order_current)
        self.admin_log_order_combo.blockSignals(False)

        user_items = ["全部用户"]
        user_items.extend(sorted({entry.user for entry in self.app_logs if entry.user}))
        self.admin_log_user_combo.blockSignals(True)
        self.admin_log_user_combo.clear()
        self.admin_log_user_combo.addItems(user_items)
        if user_current in user_items:
            self.admin_log_user_combo.setCurrentText(user_current)
        self.admin_log_user_combo.blockSignals(False)

    def _reset_admin_log_filters(self) -> None:
        if not hasattr(self, "admin_log_order_combo"):
            return
        self.admin_log_order_combo.setCurrentText("全部订单")
        self.admin_log_user_combo.setCurrentText("全部用户")
        self.admin_log_date_check.setChecked(False)
        self.admin_log_start_date.setDate(QtCore.QDate.currentDate().addMonths(-1))
        self.admin_log_end_date.setDate(QtCore.QDate.currentDate())
        self._refresh_admin_log_table()

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

    def _refresh_admin_customer_codes(self) -> None:
        self.admin_customer_code_list.clear()
        self.admin_customer_code_list.addItems(self.customer_codes)

    def _refresh_admin_shipping_methods(self) -> None:
        self.admin_shipping_method_list.clear()
        self.admin_shipping_method_list.addItems(self.shipping_methods)

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

    def on_admin_customer_code_select(self) -> None:
        items = self.admin_customer_code_list.selectedItems()
        if not items:
            return
        self.admin_customer_code_edit.setText(items[0].text())

    def on_admin_shipping_method_select(self) -> None:
        items = self.admin_shipping_method_list.selectedItems()
        if not items:
            return
        self.admin_shipping_method_edit.setText(items[0].text())

    def admin_add_customer_code(self) -> None:
        code = self.admin_customer_code_edit.text().strip()
        if not code:
            return
        if code in self.customer_codes:
            QtWidgets.QMessageBox.information(self, "提示", "该客户代码已存在。")
            return
        self.customer_codes.append(code)
        self.admin_customer_code_edit.clear()
        self._refresh_admin_customer_codes()
        self._save_app_templates()

    def admin_add_shipping_method(self) -> None:
        name = self.admin_shipping_method_edit.text().strip()
        if not name:
            return
        if name in self.shipping_methods:
            QtWidgets.QMessageBox.information(self, "提示", "该发货类别已存在。")
            return
        self.shipping_methods.append(name)
        self.admin_shipping_method_edit.clear()
        self._refresh_admin_shipping_methods()
        self._save_app_templates()

    def admin_update_customer_code(self) -> None:
        items = self.admin_customer_code_list.selectedItems()
        if not items:
            return
        old_code = items[0].text()
        new_code = self.admin_customer_code_edit.text().strip()
        if not new_code:
            return
        if new_code != old_code and new_code in self.customer_codes:
            QtWidgets.QMessageBox.information(self, "提示", "该客户代码已存在。")
            return
        idx = self.customer_codes.index(old_code)
        self.customer_codes[idx] = new_code
        for order in self.orders:
            if order.customer_code == old_code:
                order.customer_code = new_code
        self._refresh_admin_customer_codes()
        self._refresh_orders_table()
        self._refresh_order_summary()
        self._save_app_templates()
        self._auto_save()

    def admin_update_shipping_method(self) -> None:
        items = self.admin_shipping_method_list.selectedItems()
        if not items:
            return
        old_name = items[0].text()
        new_name = self.admin_shipping_method_edit.text().strip()
        if not new_name:
            return
        if new_name != old_name and new_name in self.shipping_methods:
            QtWidgets.QMessageBox.information(self, "提示", "该发货类别已存在。")
            return
        idx = self.shipping_methods.index(old_name)
        self.shipping_methods[idx] = new_name
        for order in self.orders:
            if order.shipping_method == old_name:
                order.shipping_method = new_name
        self._refresh_admin_shipping_methods()
        self._refresh_orders_table()
        self._refresh_order_summary()
        self._save_app_templates()
        self._auto_save()

    def admin_remove_customer_code(self) -> None:
        items = self.admin_customer_code_list.selectedItems()
        if not items:
            return
        code = items[0].text()
        if code in self.customer_codes:
            self.customer_codes.remove(code)
        for order in self.orders:
            if order.customer_code == code:
                order.customer_code = ""
        self._refresh_admin_customer_codes()
        self._refresh_orders_table()
        self._refresh_order_summary()
        self._save_app_templates()
        self._auto_save()

    def admin_remove_shipping_method(self) -> None:
        items = self.admin_shipping_method_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.shipping_methods:
            self.shipping_methods.remove(name)
        for order in self.orders:
            if order.shipping_method == name:
                order.shipping_method = ""
        self._refresh_admin_shipping_methods()
        self._refresh_orders_table()
        self._refresh_order_summary()
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
        self._sync_equipment_to_orders()
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
        self._sync_equipment_to_orders()
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
            self.admin_equipment_template_table.setItem(
                row, 4, QtWidgets.QTableWidgetItem(eq.shift_template_name or "-")
            )

    def _refresh_admin_equipment_shift_combo(self) -> None:
        if not hasattr(self, "admin_equipment_shift_combo"):
            return
        valid_templates = {tpl.name for tpl in self.shift_templates}
        changed = False
        for eq in self.equipment_templates:
            if eq.shift_template_name and eq.shift_template_name not in valid_templates:
                eq.shift_template_name = ""
                changed = True
        for eq in self.equipment:
            if eq.shift_template_name and eq.shift_template_name not in valid_templates:
                eq.shift_template_name = ""
                changed = True
        current = self.admin_equipment_shift_combo.currentText()
        items = ["默认班次"] + [tpl.name for tpl in self.shift_templates]
        self.admin_equipment_shift_combo.blockSignals(True)
        self.admin_equipment_shift_combo.clear()
        self.admin_equipment_shift_combo.addItems(items)
        if current in items:
            self.admin_equipment_shift_combo.setCurrentText(current)
        else:
            self.admin_equipment_shift_combo.setCurrentIndex(0)
        self.admin_equipment_shift_combo.blockSignals(False)
        if changed:
            self._refresh_admin_equipment_templates_table()

    def on_admin_equipment_template_select(self) -> None:
        row = self.admin_equipment_template_table.currentRow()
        if row < 0 or row >= len(self.equipment_templates):
            return
        eq = self.equipment_templates[row]
        self.admin_equipment_id_edit.setText(eq.equipment_id)
        self.admin_equipment_category_combo.setCurrentText(eq.category)
        self.admin_equipment_total_spin.setValue(eq.total_count)
        self.admin_equipment_available_spin.setValue(eq.available_count)
        if hasattr(self, "admin_equipment_shift_combo"):
            self.admin_equipment_shift_combo.setCurrentText(
                eq.shift_template_name or "默认班次"
            )

    def admin_add_or_update_equipment_template(self) -> None:
        eq_id = self.admin_equipment_id_edit.text().strip()
        if not eq_id:
            return
        category = self.admin_equipment_category_combo.currentText().strip()
        total = int(self.admin_equipment_total_spin.value())
        available = int(self.admin_equipment_available_spin.value())
        if available > total:
            available = total
        shift_name = ""
        if hasattr(self, "admin_equipment_shift_combo"):
            shift_text = self.admin_equipment_shift_combo.currentText().strip()
            if shift_text and shift_text != "默认班次":
                shift_name = shift_text
        existing = next((e for e in self.equipment_templates if e.equipment_id == eq_id), None)
        if existing:
            existing.category = category
            existing.total_count = total
            existing.available_count = available
            existing.shift_template_name = shift_name
        else:
            self.equipment_templates.append(
                Equipment(eq_id, category, total, available, shift_name)
            )
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()
        self.admin_apply_equipment_template()

    def admin_remove_equipment_template(self) -> None:
        row = self.admin_equipment_template_table.currentRow()
        if row < 0 or row >= len(self.equipment_templates):
            return
        del self.equipment_templates[row]
        self._refresh_admin_equipment_templates_table()
        self._refresh_admin_phase_template_combos()
        self._save_app_templates()
        self.admin_apply_equipment_template()

    def admin_apply_equipment_template(self) -> None:
        self.equipment = [
            Equipment(
                e.equipment_id,
                e.category,
                e.total_count,
                e.available_count,
                e.shift_template_name,
            )
            for e in self.equipment_templates
        ]
        self._log_change(f"同步设备列表 (设备数: {len(self.equipment)})")
        self._sync_equipment_to_orders()
        self._refresh_equipment_table()
        self._refresh_phase_equipment_combo()
        self._refresh_admin_phase_template_combos()
        self.refresh_eta()
        self._refresh_orders_table()
        self._auto_save()

    def _refresh_admin_phase_templates_table(self) -> None:
        self._updating_admin_phase_table = True
        self.admin_phase_template_table.blockSignals(True)
        self.admin_phase_template_table.setRowCount(0)
        for phase in self.phase_templates:
            row = self.admin_phase_template_table.rowCount()
            self.admin_phase_template_table.insertRow(row)
            name_item = QtWidgets.QTableWidgetItem(phase.name)
            hours_item = QtWidgets.QTableWidgetItem(f"{phase.planned_hours:g}")
            equipment_ids = _split_equipment_ids(phase.equipment_id)
            equipment_text = ", ".join(equipment_ids) if equipment_ids else "-"
            equipment_item = QtWidgets.QTableWidgetItem(equipment_text)
            employee_item = QtWidgets.QTableWidgetItem(phase.assigned_employee)
            group_text = "-" if phase.parallel_group == 0 else str(phase.parallel_group)
            group_item = QtWidgets.QTableWidgetItem(group_text)
            group_item.setData(QtCore.Qt.ItemDataRole.EditRole, str(phase.parallel_group))
            for item in (name_item, hours_item, employee_item, group_item):
                item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            equipment_item.setFlags(equipment_item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            self.admin_phase_template_table.setItem(row, 0, name_item)
            self.admin_phase_template_table.setItem(row, 1, hours_item)
            self.admin_phase_template_table.setItem(row, 2, equipment_item)
            self.admin_phase_template_table.setItem(row, 3, employee_item)
            self.admin_phase_template_table.setItem(row, 4, group_item)
        self.admin_phase_template_table.blockSignals(False)
        self._updating_admin_phase_table = False

    def on_admin_phase_template_select(self) -> None:
        row = self.admin_phase_template_table.currentRow()
        if row < 0 or row >= len(self.phase_templates):
            return
        phase = self.phase_templates[row]
        self.admin_phase_name_edit.setText(phase.name)
        self.admin_phase_hours_spin.setValue(phase.planned_hours)
        self._set_admin_phase_form_equipment_ids(_split_equipment_ids(phase.equipment_id))
        self.admin_phase_employee_combo.setCurrentText(phase.assigned_employee)
        self.admin_phase_parallel_spin.setValue(phase.parallel_group)

    def on_admin_phase_template_cell_double_clicked(self, row: int, col: int) -> None:
        if col != 2:
            return
        if row < 0 or row >= len(self.phase_templates):
            return
        phase = self.phase_templates[row]
        current_ids = _split_equipment_ids(phase.equipment_id)
        selected = self._choose_equipment_ids(current_ids)
        new_text = _format_equipment_ids(selected)
        if new_text != phase.equipment_id:
            phase.equipment_id = new_text
            self._refresh_admin_phase_templates_table()
            self._save_app_templates()

    def on_admin_phase_template_cell_changed(self, row: int, col: int) -> None:
        if self._updating_admin_phase_table:
            return
        if row < 0 or row >= len(self.phase_templates):
            return
        phase = self.phase_templates[row]
        item = self.admin_phase_template_table.item(row, col)
        if not item:
            return
        text = item.text().strip()
        changed = False
        if col == 0:
            if not text:
                self.admin_phase_template_table.blockSignals(True)
                item.setText(phase.name)
                self.admin_phase_template_table.blockSignals(False)
                return
            if any(p is not phase and p.name == text for p in self.phase_templates):
                QtWidgets.QMessageBox.information(self, "提示", "工序模板名称已存在。")
                self.admin_phase_template_table.blockSignals(True)
                item.setText(phase.name)
                self.admin_phase_template_table.blockSignals(False)
                return
            if text != phase.name:
                phase.name = text
                changed = True
        elif col == 1:
            try:
                hours = float(text)
            except ValueError:
                self.admin_phase_template_table.blockSignals(True)
                item.setText(f"{phase.planned_hours:g}")
                self.admin_phase_template_table.blockSignals(False)
                return
            if hours < 0:
                hours = 0.0
                self.admin_phase_template_table.blockSignals(True)
                item.setText("0")
                self.admin_phase_template_table.blockSignals(False)
            if abs(hours - phase.planned_hours) > 1e-6:
                phase.planned_hours = hours
                changed = True
        elif col == 2:
            if text in ("无需设备", "-"):
                new_eq = ""
            else:
                ids = _split_equipment_ids(text)
                if ids:
                    available = {e.equipment_id for e in self.equipment if e.equipment_id}
                    invalid = [eq_id for eq_id in ids if eq_id not in available]
                    if invalid:
                        QtWidgets.QMessageBox.information(
                            self, "提示", f"设备不存在: {', '.join(invalid)}"
                        )
                        self.admin_phase_template_table.blockSignals(True)
                        item.setText(", ".join(_split_equipment_ids(phase.equipment_id)) or "-")
                        self.admin_phase_template_table.blockSignals(False)
                        return
                    new_eq = _format_equipment_ids(ids)
                else:
                    new_eq = ""
            if new_eq != phase.equipment_id:
                phase.equipment_id = new_eq
                if new_eq == "":
                    self.admin_phase_template_table.blockSignals(True)
                    item.setText("-")
                    self.admin_phase_template_table.blockSignals(False)
                changed = True
        elif col == 3:
            if text != phase.assigned_employee:
                phase.assigned_employee = text
                changed = True
        elif col == 4:
            try:
                group = int(text)
            except ValueError:
                self.admin_phase_template_table.blockSignals(True)
                item.setText(str(phase.parallel_group))
                self.admin_phase_template_table.blockSignals(False)
                return
            if group < 0:
                group = 0
                self.admin_phase_template_table.blockSignals(True)
                item.setText("-")
                self.admin_phase_template_table.blockSignals(False)
            if group == 0 and text != "-":
                self.admin_phase_template_table.blockSignals(True)
                item.setText("-")
                self.admin_phase_template_table.blockSignals(False)
            if group != phase.parallel_group:
                phase.parallel_group = group
                if group == 0:
                    self.admin_phase_template_table.blockSignals(True)
                    item.setText("-")
                    self.admin_phase_template_table.blockSignals(False)
                changed = True

        if changed:
            self._refresh_phase_name_combo()
            self._save_app_templates()
    def admin_add_or_update_phase_template(self) -> None:
        name = self.admin_phase_name_edit.text().strip()
        if not name:
            return
        if any(ph.name == name for ph in self.phase_templates):
            QtWidgets.QMessageBox.information(self, "提示", "工序模板已存在，请使用其他名称。")
            return
        hours = float(self.admin_phase_hours_spin.value())
        equipment_ids = self._get_admin_phase_form_equipment_ids()
        equipment_id = _format_equipment_ids(equipment_ids)
        employee = self.admin_phase_employee_combo.currentText().strip()
        parallel = int(self.admin_phase_parallel_spin.value())
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
        self._refresh_phase_name_combo()
        self._save_app_templates()

    def admin_remove_phase_template(self) -> None:
        row = self.admin_phase_template_table.currentRow()
        if row < 0 or row >= len(self.phase_templates):
            return
        del self.phase_templates[row]
        self._refresh_admin_phase_templates_table()
        self._refresh_phase_name_combo()
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
        self.admin_employee_template_list.addItems(self.employees)

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
            if name != old_name and name in self.employees:
                QtWidgets.QMessageBox.information(self, "提示", "该员工已存在。")
                return
            idx = self.employees.index(old_name)
            self.employees[idx] = name
            self._log_change(f"更新员工 {old_name} -> {name}")
        else:
            if name in self.employees:
                QtWidgets.QMessageBox.information(self, "提示", "该员工已存在。")
                return
            self.employees.append(name)
            self._log_change(f"添加员工 {name}")
        self.admin_employee_name_edit.clear()
        self._refresh_admin_employee_templates_list()
        self._sync_employees_to_order()
        self._refresh_employee_list()
        self._refresh_phase_employee_combo()
        self._refresh_admin_phase_template_combos()
        self._auto_save()

    def admin_remove_employee_template(self) -> None:
        items = self.admin_employee_template_list.selectedItems()
        if not items:
            return
        name = items[0].text()
        if name in self.employees:
            self.employees.remove(name)
            self._log_change(f"删除员工 {name}")
        self._refresh_admin_employee_templates_list()
        self._sync_employees_to_order()
        self._refresh_employee_list()
        self._refresh_phase_employee_combo()
        self._refresh_admin_phase_template_combos()
        self._auto_save()

    def admin_apply_employee_template(self) -> None:
        self._log_change(f"同步员工列表 (员工数: {len(self.employees)})")
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
        self._refresh_admin_equipment_shift_combo()
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
        self._refresh_admin_equipment_shift_combo()
        self._save_app_templates()
        self._apply_active_shift_template()

    def admin_set_active_shift_template(self) -> None:
        items = self.admin_shift_template_list.selectedItems()
        if not items:
            return
        self.active_shift_template_name = items[0].text()
        self._update_active_shift_label()
        self._refresh_admin_equipment_shift_combo()
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
        employee_names = set(self.employees)
        if hasattr(self, "admin_phase_equipment_display"):
            current_ids = _split_equipment_ids(self.admin_phase_equipment_display.text())
            available = {e.equipment_id for e in self.equipment if e.equipment_id}
            filtered = [eq_id for eq_id in current_ids if eq_id in available]
            self._set_admin_phase_form_equipment_ids(filtered)

        current_emp = self.admin_phase_employee_combo.currentText()
        self.admin_phase_employee_combo.clear()
        emp_items = [""] + sorted(employee_names)
        self.admin_phase_employee_combo.addItems(emp_items)
        if current_emp in emp_items:
            self.admin_phase_employee_combo.setCurrentText(current_emp)
        else:
            self.admin_phase_employee_combo.setCurrentIndex(0)

    def _clone_phase_list(self, phases: List[Phase]) -> List[Phase]:
        return [
            Phase(
                name=ph.name,
                planned_hours=ph.planned_hours,
                completed_hours=ph.completed_hours,
                parallel_group=ph.parallel_group,
                equipment_id=ph.equipment_id,
                assigned_employee=ph.assigned_employee,
            )
            for ph in phases
        ]

    def _refresh_admin_phase_template_sets(self) -> None:
        if not hasattr(self, "admin_phase_template_combo"):
            return
        current = self.admin_phase_template_combo.currentText()
        names = list(self.phase_template_sets.keys())
        self.admin_phase_template_combo.blockSignals(True)
        self.admin_phase_template_combo.clear()
        self.admin_phase_template_combo.addItems(names)
        target = self.active_phase_template_name or current
        if target in names:
            self.admin_phase_template_combo.setCurrentText(target)
        elif names:
            self.admin_phase_template_combo.setCurrentIndex(0)
            self.active_phase_template_name = self.admin_phase_template_combo.currentText()
        self.admin_phase_template_combo.blockSignals(False)
        if hasattr(self, "admin_phase_template_name_edit"):
            self.admin_phase_template_name_edit.setText(self.active_phase_template_name)

    def on_admin_phase_template_set_change(self, name: str) -> None:
        name = (name or "").strip()
        if not name or name not in self.phase_template_sets:
            return
        self.active_phase_template_name = name
        self.phase_templates = self._clone_phase_list(self.phase_template_sets[name])
        if hasattr(self, "admin_phase_template_name_edit"):
            self.admin_phase_template_name_edit.setText(name)
        self._refresh_admin_phase_templates_table()
        self._refresh_phase_name_combo()
        self._save_app_templates()

    def admin_save_phase_template_set(self) -> None:
        name = ""
        if hasattr(self, "admin_phase_template_name_edit"):
            name = self.admin_phase_template_name_edit.text().strip()
        if not name:
            name = self.active_phase_template_name or ""
        if not name:
            QtWidgets.QMessageBox.information(self, "提示", "请输入模版名称。")
            return
        self.phase_template_sets[name] = self._clone_phase_list(self.phase_templates)
        self.active_phase_template_name = name
        self._refresh_admin_phase_template_sets()
        self._refresh_phase_name_combo()
        self._save_app_templates()

    def admin_delete_phase_template_set(self) -> None:
        if not self.phase_template_sets:
            return
        name = self.active_phase_template_name
        if hasattr(self, "admin_phase_template_combo") and self.admin_phase_template_combo.currentText():
            name = self.admin_phase_template_combo.currentText()
        if not name:
            return
        if len(self.phase_template_sets) <= 1:
            QtWidgets.QMessageBox.information(self, "提示", "至少保留一个工序模版。")
            return
        del self.phase_template_sets[name]
        self.active_phase_template_name = next(iter(self.phase_template_sets.keys()))
        self.phase_templates = self._clone_phase_list(
            self.phase_template_sets[self.active_phase_template_name]
        )
        self._refresh_admin_phase_template_sets()
        self._refresh_admin_phase_templates_table()
        self._refresh_phase_name_combo()
        self._save_app_templates()

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
            self.detail_eta_label.setText("计划发货期: -")
            self.detail_progress_label.setText("总体进度: 0%")
            self.overall_progress.setValue(0)
            self.last_eta_dt = None
            self.last_remaining_hours = 0.0
            self.last_capacity_map = {}
            self._refresh_order_brief()
            self._refresh_visuals()
            return
        if not self.cal.shift_template:
            self.eta_value.setText("请先设置班次模板")
            self.remaining_value.setText("-")
            self.detail_eta_label.setText("计划发货期: 请先设置班次模板")
            self.detail_progress_label.setText("总体进度: -")
            self.overall_progress.setValue(0)
            self.last_eta_dt = None
            self.last_remaining_hours = 0.0
            self.last_capacity_map = {}
            self._refresh_order_brief()
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
        self.last_capacity_map = dict(result.get("daily_capacity_map", {}))
        self.eta_value.setText(eta_dt.strftime("%Y-%m-%d %H:%M"))
        self.remaining_value.setText(f"{remaining_hours:g}h")
        self.detail_eta_label.setText(f"计划发货期: {eta_dt.strftime('%Y-%m-%d %H:%M')}")

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
        self._refresh_orders_table()
        self._refresh_order_brief()
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
        self.order_summary_label.setText(f"订单: {self.order.order_id}")
        self._refresh_dashboard_summary()

    def _refresh_order_brief(self) -> None:
        if not hasattr(self, "brief_order_label"):
            return
        if not self.order:
            self.brief_order_label.setText("订单编号: -")
            self.brief_customer_label.setText("客户: -")
            self.brief_ship_label.setText("发货类别: -")
            self.brief_count_label.setText("产品数: -")
            self.brief_order_date_label.setText("订单日期: -")
            self.brief_start_date_label.setText("开始日期: -")
            self.brief_due_date_label.setText("要求发货日期: -")
            self.brief_plan_label.setText("计划发货期: -")
            self.brief_event_count_label.setText("事件数量: -")
            self.brief_defects_label.setText("不合格品: -")
            return
        order_date_text = self.order.order_date.isoformat() if self.order.order_date else "-"
        start_date_text = self.order.start_dt.date().isoformat() if self.order.start_dt else "-"
        due_date_text = self.order.due_date.isoformat() if self.order.due_date else "-"
        plan_text = self.last_eta_dt.strftime("%Y-%m-%d %H:%M") if self.last_eta_dt else "-"
        customer = self.order.customer_code or "-"
        ship_method = self.order.shipping_method or "-"
        self.brief_order_label.setText(f"订单编号: {self.order.order_id}")
        self.brief_customer_label.setText(f"客户: {customer}")
        self.brief_ship_label.setText(f"发货类别: {ship_method}")
        self.brief_count_label.setText(f"产品数: {len(self.order.products)}")
        self.brief_order_date_label.setText(f"订单日期: {order_date_text}")
        self.brief_start_date_label.setText(f"开始日期: {start_date_text}")
        self.brief_due_date_label.setText(f"要求发货日期: {due_date_text}")
        self.brief_plan_label.setText(f"计划发货期: {plan_text}")
        self.brief_event_count_label.setText(f"事件数量: {len(self.order.events)}")
        total_defects = sum(d.count for d in self.order.defects)
        self.brief_defects_label.setText(f"不合格品: {total_defects}")

    def _refresh_dashboard_summary(self) -> None:
        shift_name = self.active_shift_template_name or "-"
        if hasattr(self, "shift_summary_label"):
            self.shift_summary_label.setText(f"当前班次: {shift_name}")
        if hasattr(self, "dashboard_summary_label"):
            if not self.order:
                self.dashboard_summary_label.setText(
                    f"订单数: {len(self.orders)} | 设备数: {len(self.equipment)}"
                )
            else:
                self.dashboard_summary_label.setText(
                    f"订单数: {len(self.orders)} | 设备数: {len(self.equipment)}"
                )

    # ------------------------
    # Serialization
    # ------------------------

    def _factory_to_dict(self) -> Dict[str, object]:
        return {
            "version": 3,
            "factory_name": self.factory_name,
            "employees": list(self.employees),
            "equipment": [
                {
                    "equipment_id": e.equipment_id,
                    "category": e.category,
                    "total_count": e.total_count,
                    "available_count": e.available_count,
                    "shift_template_name": e.shift_template_name,
                }
                for e in self.equipment
            ],
            "orders": [
                self._order_to_dict(order, include_equipment=False) for order in self.orders
            ],
            "active_order_id": self.order.order_id if self.order else "",
            "app_logs": [
                {
                    "timestamp": entry.timestamp.isoformat(),
                    "user": entry.user,
                    "content": entry.content,
                    "order_id": entry.order_id,
                }
                for entry in self.app_logs
            ],
            "memos": [
                {
                    "day": entry.day.isoformat(),
                    "user": entry.user,
                    "content": entry.content,
                }
                for entry in self.memos
            ],
        }

    def _order_to_dict(self, order: Order, include_equipment: bool = True) -> Dict[str, object]:
        return order_to_dict(order, include_equipment=include_equipment)

    @staticmethod
    def _order_from_dict(data: Dict[str, object]) -> Order:
        return order_from_dict(data)

    # ------------------------
    # Global refresh
    # ------------------------

    def _refresh_all(self) -> None:
        self._refresh_orders_table()
        self._refresh_equipment_table()
        self._refresh_employee_list()
        self._refresh_phase_equipment_combo()
        self._refresh_phase_employee_combo()
        self._refresh_phase_name_combo()
        self._refresh_equipment_category_combos()
        self._refresh_products_table()
        self._refresh_events_table()
        self._refresh_adjustments_table()
        self._refresh_defects_table()
        self._refresh_defect_product_combo()
        self._refresh_defect_detail_combos()
        self._refresh_defect_category_combo()
        self._refresh_event_reason_combo()
        self._refresh_admin_views()
        self.refresh_eta()


def main() -> None:
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
