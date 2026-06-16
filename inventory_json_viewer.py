from __future__ import annotations

import ctypes
import json
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QObject, QSettings, Qt, QSortFilterProxyModel
from PySide6.QtGui import QAction, QColor, QDragEnterEvent, QDropEvent, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QStyle,
    QStyledItemDelegate,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


APP_ORG = "WarehouseTools"
APP_NAME = "InventoryJsonViewer"
APP_USER_MODEL_ID = "WarehouseTools.InventoryJsonViewer"

DETAIL_COLUMNS = [
    "品番",
    "サイズ",
    "厚み",
    "加工/裏表",
    "グレード",
    "枚数",
    "ロット",
    "備考",
    "入庫日",
    "在庫日数",
    "パレット番号",
    "保管場所",
    "色",
    "更新日時",
]
SUMMARY_COLUMNS = ["品番", "サイズ", "厚み", "加工/裏表", "グレード", "ロット", "合計枚数"]
PALLET_COLUMNS = ["パレット番号", "保管場所", "入庫日", "在庫日数", "色", "更新日時", "明細数", "合計枚数"]
MAJOR_COLUMN_PRIORITY = [
    "品番",
    "part_code",
    "items.part_code",
    "サイズ",
    "size",
    "items.size",
    "厚み",
    "thickness_mm",
    "items.thickness_mm",
    "加工/裏表",
    "finish_text",
    "items.finish_text",
    "グレード",
    "grade",
    "items.grade",
    "枚数",
    "sheet_count",
    "items.sheet_count",
    "ロット",
    "lot",
    "items.lot",
    "備考",
    "note",
    "items.note",
    "入庫日",
    "received_date",
    "在庫日数",
    "パレット番号",
    "pallet_number",
    "保管場所",
    "location_code",
    "色",
    "color_key",
    "更新日時",
    "updated_at",
]
SEARCH_COLUMN_CANDIDATES = [
    "全列",
    "パレット番号",
    "保管場所",
    "品番",
    "サイズ",
    "厚み",
    "加工/裏表",
    "グレード",
    "ロット",
    "備考",
    "入庫日",
    "在庫日数",
]
NUMERIC_COLUMNS = {"枚数", "合計枚数", "明細数", "在庫日数"}
TEXT_COLUMNS = {"パレット番号"}


@dataclass
class TableData:
    name: str
    columns: list[str]
    rows: list[dict[str, Any]]


def resource_path(filename: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / filename


def set_windows_app_user_model_id() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def app_icon() -> QIcon:
    icon_path = resource_path("icon.ico")
    return QIcon(str(icon_path)) if icon_path.exists() else QIcon()


def as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return str(value)


def as_int(value: Any) -> int:
    if value in (None, ""):
        return 0
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def inventory_days(received_date: Any, today: date) -> int | None:
    text = as_text(received_date).strip()
    if not text:
        return None
    try:
        parsed = datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None
    return (today - parsed).days


def first_value(source: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in source:
            return source.get(key)
    return ""


def flatten_value(value: Any, prefix: str = "") -> dict[str, Any]:
    if isinstance(value, dict):
        result: dict[str, Any] = {}
        for key, child in value.items():
            name = f"{prefix}.{key}" if prefix else str(key)
            result.update(flatten_value(child, name))
        return result
    if isinstance(value, list):
        return {prefix: json.dumps(value, ensure_ascii=False, separators=(",", ":"))}
    return {prefix: value}


def ordered_columns(rows: list[dict[str, Any]], preferred: list[str] | None = None) -> list[str]:
    seen: set[str] = set()
    columns: list[str] = []
    present = {column for row in rows for column in row}
    for column in preferred or []:
        if column in present and column not in seen:
            seen.add(column)
            columns.append(column)
    for row in rows:
        for column in row:
            if column not in seen:
                seen.add(column)
                columns.append(column)
    return columns


def build_inventory_detail(pallets: list[Any], today: date) -> TableData:
    rows: list[dict[str, Any]] = []
    for pallet in pallets:
        if not isinstance(pallet, dict):
            continue
        items = pallet.get("items")
        if not isinstance(items, list) or not items:
            continue
        for item in items:
            if not isinstance(item, dict):
                item = {}
            received_date = first_value(pallet, "received_date", "received_at")
            rows.append(
                {
                    "品番": as_text(first_value(item, "part_code", "product_code", "品番")),
                    "サイズ": as_text(first_value(item, "size", "サイズ")),
                    "厚み": as_text(first_value(item, "thickness_mm", "thickness", "厚み")),
                    "加工/裏表": as_text(first_value(item, "finish_text", "finish", "加工/裏表")),
                    "グレード": as_text(first_value(item, "grade", "グレード")),
                    "枚数": as_int(first_value(item, "sheet_count", "count", "枚数")),
                    "ロット": as_text(first_value(item, "lot", "ロット")),
                    "備考": as_text(first_value(item, "note", "備考")),
                    "入庫日": as_text(received_date),
                    "在庫日数": inventory_days(received_date, today),
                    "パレット番号": as_text(first_value(pallet, "pallet_number", "pallet_no")),
                    "保管場所": as_text(first_value(pallet, "location_code", "location")),
                    "色": as_text(first_value(pallet, "color_key", "color", "color_mode")),
                    "更新日時": as_text(first_value(pallet, "updated_at")),
                }
            )
    return TableData("在庫明細", DETAIL_COLUMNS, rows)


def build_inventory_summary(detail: TableData) -> TableData:
    totals: dict[tuple[str, ...], int] = defaultdict(int)
    key_columns = ["品番", "サイズ", "厚み", "加工/裏表", "グレード", "ロット"]
    for row in detail.rows:
        key = tuple(as_text(row.get(column)) for column in key_columns)
        totals[key] += as_int(row.get("枚数"))

    rows = []
    for key in sorted(totals):
        row = dict(zip(key_columns, key, strict=True))
        row["合計枚数"] = totals[key]
        rows.append(row)
    return TableData("在庫集計", SUMMARY_COLUMNS, rows)


def build_pallet_list(pallets: list[Any], today: date) -> TableData:
    rows: list[dict[str, Any]] = []
    for pallet in pallets:
        if not isinstance(pallet, dict):
            continue
        items = pallet.get("items")
        item_list = items if isinstance(items, list) else []
        received_date = first_value(pallet, "received_date", "received_at")
        rows.append(
            {
                "パレット番号": as_text(first_value(pallet, "pallet_number", "pallet_no")),
                "保管場所": as_text(first_value(pallet, "location_code", "location")),
                "入庫日": as_text(received_date),
                "在庫日数": inventory_days(received_date, today),
                "色": as_text(first_value(pallet, "color_key", "color", "color_mode")),
                "更新日時": as_text(first_value(pallet, "updated_at")),
                "明細数": len(item_list),
                "合計枚数": sum(as_int(item.get("sheet_count")) for item in item_list if isinstance(item, dict)),
            }
        )
    return TableData("パレット一覧", PALLET_COLUMNS, rows)


def build_expanded_table(name: str, values: Any) -> TableData:
    rows: list[dict[str, Any]] = []
    source = values if isinstance(values, list) else []
    for value in source:
        if not isinstance(value, dict):
            rows.append({"値": value})
            continue

        base = {k: v for k, v in value.items() if k != "items"}
        flat_base = flatten_value(base)
        items = value.get("items")
        if isinstance(items, list) and items:
            for item in items:
                flat_item = flatten_value(item, "items") if isinstance(item, dict) else {"items": item}
                rows.append({**flat_base, **flat_item})
        else:
            rows.append(flat_base)

    return TableData(name, ordered_columns(rows, MAJOR_COLUMN_PRIORITY), rows)


def build_tables(data: Any, today: date | None = None) -> list[TableData]:
    if not isinstance(data, dict):
        raise ValueError("JSONの最上位はオブジェクトである必要があります。")
    today = today or date.today()
    pallets_value = data.get("pallets", [])
    pallets = pallets_value if isinstance(pallets_value, list) else []
    detail = build_inventory_detail(pallets, today)
    return [
        detail,
        build_inventory_summary(detail),
        build_pallet_list(pallets, today),
        build_expanded_table("出庫履歴", data.get("shipments", [])),
        build_expanded_table("マップメモ", data.get("map_notes", [])),
    ]


class ReadOnlyDelegate(QStyledItemDelegate):
    def createEditor(self, parent: QWidget, option, index: QModelIndex) -> None:  # type: ignore[override]
        return None


class TableModel(QAbstractTableModel):
    def __init__(self, table: TableData, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.table = table

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.table.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.table.columns)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid():
            return None
        column = self.table.columns[index.column()]
        value = self.table.rows[index.row()].get(column, "")
        if role == Qt.DisplayRole:
            return as_text(value)
        if role == Qt.EditRole:
            return None
        if role == Qt.ForegroundRole:
            return QColor("#111827")
        if role == Qt.BackgroundRole:
            return QColor("#F9FAFB") if index.row() % 2 else QColor("#FFFFFF")
        if role == Qt.TextAlignmentRole and column in NUMERIC_COLUMNS:
            return Qt.AlignRight | Qt.AlignVCenter
        if role == Qt.UserRole:
            return value
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.table.columns[section]
        return section + 1

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class InventoryProxyModel(QSortFilterProxyModel):
    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.search_column = "全列"
        self.search_text = ""

    def set_search(self, column: str, text: str) -> None:
        self.search_column = column
        self.search_text = text.casefold().strip()
        self.invalidateRowsFilter()

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        if not self.search_text:
            return True
        source = self.sourceModel()
        if source is None:
            return True

        if self.search_column != "全列":
            for column in range(source.columnCount()):
                if source.headerData(column, Qt.Horizontal) == self.search_column:
                    value = source.index(source_row, column, source_parent).data(Qt.DisplayRole)
                    return self.search_text in as_text(value).casefold()
            return True

        for column in range(source.columnCount()):
            value = source.index(source_row, column, source_parent).data(Qt.DisplayRole)
            if self.search_text in as_text(value).casefold():
                return True
        return False

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        source = self.sourceModel()
        header = source.headerData(left.column(), Qt.Horizontal) if source else ""
        left_value = left.data(Qt.UserRole)
        right_value = right.data(Qt.UserRole)
        if header in NUMERIC_COLUMNS:
            return as_int(left_value) < as_int(right_value)
        return as_text(left_value).casefold() < as_text(right_value).casefold()


class InventoryTableView(QTableView):
    def __init__(self, table_name: str, settings: QSettings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.table_name = table_name
        self.settings = settings
        self.setItemDelegate(ReadOnlyDelegate(self))
        self.setEditTriggers(QTableView.NoEditTriggers)
        self.setSelectionMode(QTableView.ExtendedSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self.setAlternatingRowColors(False)
        self.setSortingEnabled(True)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)
        self.horizontalHeader().setSectionsMovable(True)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.horizontalHeader().setStretchLastSection(False)
        self.horizontalHeader().sectionMoved.connect(self.save_header_state)
        self.horizontalHeader().sectionResized.connect(self.save_header_state)
        self.verticalHeader().setDefaultSectionSize(24)
        self.setHorizontalScrollMode(QTableView.ScrollPerPixel)
        self.setVerticalScrollMode(QTableView.ScrollPerPixel)
        self.setStyleSheet(
            """
            QTableView {
                color: #111827;
                background-color: #FFFFFF;
                alternate-background-color: #F9FAFB;
                selection-color: #FFFFFF;
                selection-background-color: #2563EB;
                gridline-color: #E5E7EB;
            }
            QHeaderView::section {
                color: #111827;
                background-color: #E5E7EB;
                border: 1px solid #D1D5DB;
                padding: 4px;
            }
            """
        )

        copy_action = QAction("コピー", self)
        copy_action.setShortcut(QKeySequence.Copy)
        copy_action.triggered.connect(lambda: self.copy_selection(False))
        self.addAction(copy_action)

    def setModel(self, model) -> None:  # type: ignore[override]
        super().setModel(model)
        self.restore_header_state()

    def settings_key(self) -> str:
        return f"tables/{self.table_name}/header_state_v2"

    def save_header_state(self, *args: Any) -> None:
        self.settings.setValue(self.settings_key(), self.horizontalHeader().saveState())

    def restore_header_state(self) -> None:
        state = self.settings.value(self.settings_key())
        self.horizontalHeader().blockSignals(True)
        if state:
            self.horizontalHeader().restoreState(state)
        else:
            self.resizeColumnsToContents()
            for logical in range(self.model().columnCount() if self.model() else 0):
                current = self.columnWidth(logical)
                self.setColumnWidth(logical, min(max(current, 80), 260))
        self.horizontalHeader().blockSignals(False)

    def open_context_menu(self, position) -> None:
        menu = QMenu(self)
        menu.addAction("コピー", lambda: self.copy_selection(False))
        menu.addAction("ヘッダー付きコピー", lambda: self.copy_selection(True))
        menu.exec(self.viewport().mapToGlobal(position))

    def copy_selection(self, with_header: bool) -> None:
        model = self.model()
        if model is None:
            return
        selected = self.selectionModel().selectedIndexes() if self.selectionModel() else []
        if not selected:
            return
        rows = sorted({index.row() for index in selected})
        columns = sorted({index.column() for index in selected})
        selected_set = {(index.row(), index.column()) for index in selected}
        visual_columns = sorted(columns, key=lambda c: self.horizontalHeader().visualIndex(c))

        lines: list[str] = []
        if with_header:
            lines.append("\t".join(as_text(model.headerData(column, Qt.Horizontal)) for column in visual_columns))
        for row in rows:
            values = []
            for column in visual_columns:
                values.append(as_text(model.index(row, column).data(Qt.DisplayRole)) if (row, column) in selected_set else "")
            lines.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(lines))

    def current_export_columns(self) -> list[int]:
        model = self.model()
        if model is None:
            return []
        columns = list(range(model.columnCount()))
        return sorted(columns, key=lambda c: self.horizontalHeader().visualIndex(c))


class MainWindow(QMainWindow):
    def __init__(self, icon: QIcon | None = None) -> None:
        super().__init__()
        self.setWindowTitle("倉庫 inventory-data JSON ビューア")
        if icon is not None and not icon.isNull():
            self.setWindowIcon(icon)
        self.resize(1280, 820)
        self.setAcceptDrops(True)
        self.settings = QSettings(APP_ORG, APP_NAME)
        self.today = date.today()
        self.current_file: Path | None = None
        self.tables: dict[str, InventoryTableView] = {}

        self.select_button = QPushButton("ファイルを選択")
        self.select_button.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.select_button.clicked.connect(self.choose_file)
        self.export_visible_only = QCheckBox("検索結果のみ出力")
        self.export_visible_only.setChecked(True)
        self.export_button = QPushButton("Excel出力")
        self.export_button.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.export_button.clicked.connect(self.export_excel)
        self.export_button.setEnabled(False)
        self.path_label = QLabel("inventory-data-*.json をドラッグ&ドロップ、またはファイルを選択してください")
        self.path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top = QHBoxLayout()
        top.addWidget(self.select_button)
        top.addWidget(self.path_label, 1)
        top.addWidget(self.export_visible_only)
        top.addWidget(self.export_button)

        self.search_column_combo = QComboBox()
        self.search_text = QLineEdit()
        self.search_text.setPlaceholderText("検索文字を入力")
        self.clear_search_button = QPushButton("検索クリア")
        self.clear_search_button.clicked.connect(self.clear_search)
        self.search_column_combo.currentTextChanged.connect(self.apply_search)
        self.search_text.textChanged.connect(self.apply_search)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("検索対象"))
        search_row.addWidget(self.search_column_combo)
        search_row.addWidget(self.search_text, 1)
        search_row.addWidget(self.clear_search_button)

        self.drop_label = QLabel("ここへ inventory-data-*.json をドラッグ&ドロップできます")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setFrameShape(QFrame.StyledPanel)
        self.drop_label.setStyleSheet(
            "QLabel { border: 1px dashed #64748B; color: #334155; padding: 10px; background: #F8FAFC; }"
        )

        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self.on_tab_changed)
        self.count_label = QLabel("0件")

        layout = QVBoxLayout()
        layout.addLayout(top)
        layout.addLayout(search_row)
        layout.addWidget(self.drop_label)
        layout.addWidget(self.tabs, 1)
        layout.addWidget(self.count_label)

        root = QWidget()
        root.setLayout(layout)
        self.setCentralWidget(root)
        self.refresh_search_columns()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if any(url.toLocalFile().lower().endswith(".json") for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        for url in event.mimeData().urls():
            local_file = url.toLocalFile()
            if local_file.lower().endswith(".json"):
                self.load_file(Path(local_file))
                event.acceptProposedAction()
                return

    def choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "JSONファイルを選択",
            str(Path.home() / "Downloads"),
            "JSON files (*.json)",
        )
        if filename:
            self.load_file(Path(filename))

    def show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)
        self.count_label.setText(message)

    def load_file(self, path: Path) -> None:
        try:
            text = path.read_text(encoding="utf-8")
            data = json.loads(text)
            tables = build_tables(data, self.today)
        except UnicodeDecodeError:
            self.show_error("読み込みエラー", "JSONファイルをUTF-8として読み込めませんでした。")
            return
        except json.JSONDecodeError as exc:
            self.show_error("JSON形式エラー", f"JSON形式が不正です。\n{exc}")
            return
        except Exception as exc:
            self.show_error("読み込みエラー", f"読み込み中にエラーが発生しました。\n{exc}")
            return

        self.current_file = path
        self.path_label.setText(str(path))
        self.populate_tabs(tables)
        self.export_button.setEnabled(True)
        self.on_tab_changed()

    def populate_tabs(self, tables: list[TableData]) -> None:
        self.tabs.clear()
        self.tables.clear()
        for table in tables:
            source_model = TableModel(table, self)
            proxy = InventoryProxyModel(self)
            proxy.setSourceModel(source_model)
            view = InventoryTableView(table.name, self.settings, self)
            view.setModel(proxy)
            self.tables[table.name] = view
            self.tabs.addTab(view, table.name)

    def current_view(self) -> InventoryTableView | None:
        widget = self.tabs.currentWidget()
        return widget if isinstance(widget, InventoryTableView) else None

    def current_proxy(self) -> InventoryProxyModel | None:
        view = self.current_view()
        model = view.model() if view else None
        return model if isinstance(model, InventoryProxyModel) else None

    def on_tab_changed(self) -> None:
        self.refresh_search_columns()
        self.apply_search()
        self.update_count_label()

    def refresh_search_columns(self) -> None:
        current = self.search_column_combo.currentText() or "全列"
        view = self.current_view()
        model = view.model() if view else None
        available = ["全列"]
        if model is not None:
            headers = [as_text(model.headerData(column, Qt.Horizontal)) for column in range(model.columnCount())]
            available.extend([column for column in SEARCH_COLUMN_CANDIDATES[1:] if column in headers])
            available.extend([column for column in headers if column not in available])

        self.search_column_combo.blockSignals(True)
        self.search_column_combo.clear()
        self.search_column_combo.addItems(available)
        self.search_column_combo.setCurrentText(current if current in available else "全列")
        self.search_column_combo.blockSignals(False)

    def apply_search(self) -> None:
        proxy = self.current_proxy()
        if proxy is not None:
            proxy.set_search(self.search_column_combo.currentText() or "全列", self.search_text.text())
        self.update_count_label()

    def clear_search(self) -> None:
        self.search_text.clear()
        self.search_column_combo.setCurrentText("全列")
        self.apply_search()

    def update_count_label(self) -> None:
        view = self.current_view()
        model = view.model() if view else None
        if model is None:
            self.count_label.setText("0件")
            return
        source_count = model.sourceModel().rowCount() if isinstance(model, InventoryProxyModel) and model.sourceModel() else model.rowCount()
        visible_count = model.rowCount()
        name = self.tabs.tabText(self.tabs.currentIndex()) if self.tabs.currentIndex() >= 0 else ""
        if visible_count == source_count:
            self.count_label.setText(f"{name}: {visible_count}件")
        else:
            self.count_label.setText(f"{name}: {visible_count}件 / 全{source_count}件")

    def export_excel(self) -> None:
        if self.current_file is None:
            return
        default_path = str(self.current_file.with_suffix(".xlsx"))
        output_path, _ = QFileDialog.getSaveFileName(self, "Excel出力", default_path, "Excel files (*.xlsx)")
        if not output_path:
            return
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        workbook = Workbook()
        workbook.remove(workbook.active)
        visible_only = self.export_visible_only.isChecked()
        for index in range(self.tabs.count()):
            name = self.tabs.tabText(index)
            view = self.tabs.widget(index)
            if not isinstance(view, InventoryTableView) or view.model() is None:
                continue
            model = view.model()
            sheet = workbook.create_sheet(title=name)
            columns = view.current_export_columns()
            headers = [as_text(model.headerData(column, Qt.Horizontal)) for column in columns]
            sheet.append(headers)

            row_indexes = range(model.rowCount())
            if not visible_only and isinstance(model, InventoryProxyModel) and model.sourceModel() is not None:
                row_indexes = range(model.sourceModel().rowCount())

            for row in row_indexes:
                values = []
                for column in columns:
                    header = as_text(model.headerData(column, Qt.Horizontal))
                    if visible_only:
                        cell_index = model.index(row, column)
                        value = cell_index.data(Qt.DisplayRole)
                        raw_value = cell_index.data(Qt.UserRole)
                    elif isinstance(model, InventoryProxyModel) and model.sourceModel() is not None:
                        source_index = model.sourceModel().index(row, column)
                        value = source_index.data(Qt.DisplayRole)
                        raw_value = source_index.data(Qt.UserRole)
                    else:
                        cell_index = model.index(row, column)
                        value = cell_index.data(Qt.DisplayRole)
                        raw_value = cell_index.data(Qt.UserRole)
                    if header in NUMERIC_COLUMNS:
                        values.append("" if raw_value in (None, "") else as_int(raw_value))
                    elif header in TEXT_COLUMNS:
                        values.append(as_text(value))
                    else:
                        values.append(value)
                sheet.append(values)

            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(fill_type="solid", fgColor="334155")

            for column_cells in sheet.columns:
                header = as_text(column_cells[0].value)
                max_len = max(len(as_text(cell.value)) for cell in column_cells)
                width = min(max(max_len + 2, 10), 42)
                letter = get_column_letter(column_cells[0].column)
                sheet.column_dimensions[letter].width = width
                if header in TEXT_COLUMNS:
                    for cell in column_cells[1:]:
                        cell.number_format = "@"
                elif header in NUMERIC_COLUMNS:
                    for cell in column_cells[1:]:
                        cell.number_format = "0"

        try:
            workbook.save(output_path)
        except Exception as exc:
            QMessageBox.critical(self, "Excel出力エラー", f"xlsx出力に失敗しました。\n{exc}")
            return
        QMessageBox.information(self, "Excel出力完了", f"保存しました。\n{output_path}")


def main() -> int:
    set_windows_app_user_model_id()
    app = QApplication(sys.argv)
    icon = app_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)
    window = MainWindow(icon)
    if len(sys.argv) > 1:
        window.load_file(Path(sys.argv[1]))
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
