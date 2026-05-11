from __future__ import annotations

import csv
import sys
from difflib import SequenceMatcher
from pathlib import Path
from time import perf_counter

from PySide6.QtCore import (
    QAbstractTableModel,
    QItemSelectionModel,
    QModelIndex,
    QObject,
    QSize,
    Qt,
    QThread,
    QTimer,
    Signal,
)
from PySide6.QtGui import QColor, QBrush, QFont, QFontDatabase, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStatusBar,
    QTableView,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .alignment import ComparisonOptions, align_sheets, align_sheets_three_way, compare_text_values, normalize_header
from .excel_xml import clone_row_with_values
from .exporter import (
    build_diff_report_rows,
    export_diff_report,
    export_workbook,
    format_diff_report_html,
    format_diff_report_markdown,
)
from .merge_rules import get_merge_rule, list_merge_rules
from .models import AlignedRow, CellData, SheetAlignment, WorkbookData
from .text_diff import compute_inline_diff, render_diff_inline_html

import html as _html


def _escape_tooltip(value: str | None) -> str:
    return _html.escape(value or "", quote=True).replace("\n", "<br>")
from .settings import (
    format_config_label,
    load_settings,
    remember_config,
    remember_root,
    replace_quick_roots,
    save_settings,
)
from .sources import (
    SOURCE_GOOGLE_SHEETS,
    SOURCE_LOCAL,
    SOURCE_SVN,
    SourceBrowseEntry,
    SourceFileEntry,
    WorkbookSource,
    RevisionEntry,
    clear_sources_caches,
    default_google_oauth_token_path,
    describe_google_sheet,
    is_google_sheets_target,
    list_local_table_files,
    list_svn_directory,
    list_svn_revisions,
    list_svn_xml_files,
    load_workbook_from_source,
    infer_source_kind,
    join_source_target,
    parse_google_sheet_target,
    normalize_source_target,
    set_google_auth_settings,
    start_google_oauth_login,
    source_relative_path,
    source_path_name,
)
from .updater import PreparedUpdate, UpdateInfo, check_app_update, compare_versions, launch_prepared_update, prepare_update
from .version import APP_VERSION


STATUS_COLORS = {
    "same": QColor("#ffffff"),
    "left_only": QColor("#e0f2fe"),
    "right_only": QColor("#ffedd5"),
    "conflict": QColor("#fee2e2"),
    "deleted": QColor("#f1f5f9"),
}

SIMPLIFIED_CELL_BACKGROUNDS = {
    "same": QColor("#ffffff"),
    "changed": QColor("#fef9c3"),
    "middle_conflict": QColor("#fef08a"),
    "left_conflict": QColor("#e0f2fe"),
    "right_conflict": QColor("#ffedd5"),
    "deleted": QColor("#f1f5f9"),
}

SIMPLIFIED_CELL_BRUSHES = {key: QBrush(color) for key, color in SIMPLIFIED_CELL_BACKGROUNDS.items()}

_HEADER_FOREGROUND_BRUSH = QBrush(QColor("#64748b"))
_KEY_HEADER_BACKGROUND_BRUSH = QBrush(QColor("#dbeafe"))
_KEY_HEADER_FOREGROUND_BRUSH = QBrush(QColor("#1d4ed8"))
_GROUP_FOREGROUND_BRUSH = QBrush(QColor("#334155"))
_NOTE_FOREGROUND_BRUSH = QBrush(QColor("#475569"))
_DELETED_FOREGROUND_BRUSH = QBrush(QColor("#94a3b8"))
_CONFLICT_FOREGROUND_BRUSH = QBrush(QColor("#ef4444"))

_DIFFERENCE_FOREGROUND_BRUSHES = {
    "left": QBrush(QColor("#0ea5e9")),
    "right": QBrush(QColor("#f59e0b")),
    "middle": QBrush(QColor("#ea580c")),
}

_ROW_HEADER_BACKGROUND_BRUSHES = {
    "conflict": QBrush(QColor("#fecaca")),
    "right_only": QBrush(QColor("#d1fae5")),
    "left_only_or_deleted": QBrush(QColor("#ffe4e6")),
    "changed": QBrush(QColor("#fef3c7")),
    "same": QBrush(QColor("#f8fafc")),
}

_ROW_HEADER_FOREGROUND_BRUSHES = {
    "conflict": QBrush(QColor("#ef4444")),
    "right_only": QBrush(QColor("#10b981")),
    "left_only_or_deleted": QBrush(QColor("#f43f5e")),
    "changed": QBrush(QColor("#d97706")),
    "same": QBrush(QColor("#334155")),
}


def _build_font_variants() -> dict[tuple[bool, bool], QFont]:
    variants: dict[tuple[bool, bool], QFont] = {}
    for bold in (False, True):
        for strike in (False, True):
            font = QFont()
            if bold:
                font.setBold(True)
            if strike:
                font.setStrikeOut(True)
            variants[(bold, strike)] = font
    return variants


_FONT_VARIANTS = _build_font_variants()

PREFERRED_UI_FONTS = [
    "Microsoft YaHei UI",
    "Microsoft YaHei",
    "微软雅黑",
    "DengXian",
    "等线",
    "SimSun",
    "宋体",
]

SOURCE_LABELS = {
    SOURCE_LOCAL: "本地",
    SOURCE_SVN: "SVN",
    SOURCE_GOOGLE_SHEETS: "Google Sheets",
}


class MergeTableModel(QAbstractTableModel):
    def __init__(
        self,
        alignment: SheetAlignment,
        side: str,
        comparison_options: ComparisonOptions | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.alignment = alignment
        self.side = side
        self.comparison_options = comparison_options or ComparisonOptions()
        self.visible_rows = list(range(len(alignment.rows)))
        self._row_diff_cache: dict[int, bool] = {}
        self._cell_diff_cache: dict[tuple[int, int], bool] = {}
        self._row_text_lower_cache: dict[int, str] = {}

    def set_alignment(self, alignment: SheetAlignment, comparison_options: ComparisonOptions | None = None) -> None:
        self.beginResetModel()
        self.alignment = alignment
        if comparison_options is not None:
            self.comparison_options = comparison_options
        self.visible_rows = list(range(len(alignment.rows)))
        self._row_diff_cache.clear()
        self._cell_diff_cache.clear()
        self._row_text_lower_cache.clear()
        self.endResetModel()

    def invalidate_diff_cache(self, row: AlignedRow | None = None) -> None:
        if row is None:
            self._row_diff_cache.clear()
            self._cell_diff_cache.clear()
            self._row_text_lower_cache.clear()
            return
        row_key = id(row)
        self._row_diff_cache.pop(row_key, None)
        self._row_text_lower_cache.pop(row_key, None)
        for key in list(self._cell_diff_cache.keys()):
            if key[0] == row_key:
                del self._cell_diff_cache[key]

    def set_filter(self, view_mode: str, search_text: str) -> None:
        self.set_visible_rows(self.compute_visible_rows(view_mode, search_text))

    def set_visible_rows(self, visible_rows: list[int]) -> None:
        self.beginResetModel()
        self.visible_rows = list(visible_rows)
        self.endResetModel()

    def compute_visible_rows(self, view_mode: str, search_text: str) -> list[int]:
        text = search_text.strip().lower()
        visible_rows: list[int] = []
        for index, row in enumerate(self.alignment.rows):
            if view_mode == "conflict" and row.status != "conflict":
                continue
            if view_mode == "changed" and not self._row_has_difference(row):
                continue
            if text and text not in self._row_text_lower(row):
                continue
            visible_rows.append(index)
        return visible_rows

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.visible_rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.alignment.columns)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if orientation == Qt.Horizontal:
            column = self.alignment.columns[section]
            if role == Qt.DisplayRole:
                return self._header_title(column)
            if role == Qt.ToolTipRole:
                return self._header_tooltip(column)
            if role == Qt.BackgroundRole and normalize_header(column.key) in set(self.alignment.key_fields or []):
                return _KEY_HEADER_BACKGROUND_BRUSH
            if role == Qt.ForegroundRole and normalize_header(column.key) in set(self.alignment.key_fields or []):
                return _KEY_HEADER_FOREGROUND_BRUSH
            if role == Qt.FontRole:
                if normalize_header(column.key) in set(self.alignment.key_fields or []):
                    return _FONT_VARIANTS[(True, False)]
                if column.left_index is None or column.right_index is None:
                    return _FONT_VARIANTS[(True, False)]
                return None
            return None
        if section < 0 or section >= len(self.visible_rows):
            return None
        row = self._aligned_row(section)
        if role == Qt.DisplayRole:
            marker = self._row_header_marker(row)
            return f"{section + 1}{marker}"
        if role == Qt.ToolTipRole:
            return row.reason
        if role == Qt.BackgroundRole:
            return self._row_header_background_brush(row)
        if role == Qt.ForegroundRole:
            return self._row_header_foreground_brush(row)
        if role == Qt.FontRole:
            return _FONT_VARIANTS[(self._row_header_has_difference(row), False)]
        return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None
        row = self._aligned_row(index.row())
        column = self.alignment.columns[index.column()]
        value = self._cell_value(row, index.column(), column)

        if role in {Qt.DisplayRole, Qt.EditRole}:
            return value
        if role == Qt.BackgroundRole:
            return self._background_for_cell(row, index.column() + 1, column)
        if role == Qt.ForegroundRole:
            kind = row.merged_row.kind
            if kind == "header":
                return _HEADER_FOREGROUND_BRUSH
            if kind == "group":
                return _GROUP_FOREGROUND_BRUSH
            if kind == "note":
                return _NOTE_FOREGROUND_BRUSH
            if row.status == "deleted":
                return _DELETED_FOREGROUND_BRUSH
            if self.side != "middle" and index.column() + 1 in row.conflict_columns:
                return _CONFLICT_FOREGROUND_BRUSH
            if self._cell_has_difference(row, index.column() + 1, column):
                return _DIFFERENCE_FOREGROUND_BRUSHES.get(self.side, _DIFFERENCE_FOREGROUND_BRUSHES["middle"])
            return None
        if role == Qt.FontRole:
            kind = row.merged_row.kind
            if kind in {"header", "group"}:
                return _FONT_VARIANTS[(True, False)]
            if row.status == "deleted":
                return _FONT_VARIANTS[(False, True)]
            if self.side != "middle" and index.column() + 1 in row.conflict_columns:
                return _FONT_VARIANTS[(True, False)]
            if self._cell_has_difference(row, index.column() + 1, column):
                return _FONT_VARIANTS[(True, False)]
            return None
        if role == Qt.ToolTipRole:
            return self._cell_tooltip(row, index.column(), column, value)
        return None

    def _cell_tooltip(self, row: AlignedRow, column_index: int, column, value) -> str | None:
        logical_column = column_index + 1
        left_value = row.left_row.value_at(column.left_index or -1) if row.left_row and column.left_index else ""
        right_value = row.right_row.value_at(column.right_index or -1) if row.right_row and column.right_index else ""
        base_value = ""
        base_row = getattr(row, "base_row", None)
        base_index = getattr(column, "base_index", None)
        if base_row is not None and base_index:
            base_value = base_row.value_at(base_index or -1)
        is_conflict = logical_column in row.conflict_columns
        has_diff = not compare_text_values(left_value, right_value, self.comparison_options)
        show_diff = is_conflict or has_diff
        if not show_diff:
            parts: list[str] = []
            if value:
                parts.append(str(value))
            if row.reason:
                parts.append(f"[{row.reason}]")
            return "\n".join(parts) if parts else None

        lines: list[str] = []
        header_label = "冲突" if is_conflict else "差异"
        column_kind = getattr(row, "column_conflict_kinds", {}).get(logical_column, "") if hasattr(row, "column_conflict_kinds") else ""
        kind_label_map = {
            "three_way_diverge": "三方分歧",
            "three_way_left_modified": "左单方改",
            "three_way_right_modified": "右单方改",
            "three_way_same_edit": "双方同改",
        }
        if column_kind in kind_label_map:
            header_label = f"{header_label}（{kind_label_map[column_kind]}）"
        lines.append(f'<b style="color:#b31b1b">{header_label}</b>')
        inline_html = render_diff_inline_html(compute_inline_diff(left_value, right_value))
        if inline_html:
            lines.append(f'<div style="font-family:Consolas,monospace">{inline_html}</div>')
        safe_left = _escape_tooltip(left_value)
        safe_right = _escape_tooltip(right_value)
        lines.append(f'<div style="color:#555"><b>左:</b> {safe_left}</div>')
        lines.append(f'<div style="color:#555"><b>右:</b> {safe_right}</div>')
        if base_value:
            lines.append(f'<div style="color:#777"><i><b>base:</b> {_escape_tooltip(base_value)}</i></div>')
        if row.reason:
            lines.append(f'<div style="color:#777;font-size:11px">[{_escape_tooltip(row.reason)}]</div>')
        return "".join(lines)

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        flags = super().flags(index)
        if self.side == "middle":
            flags |= Qt.ItemIsEditable
        return flags

    def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
        if self.side != "middle" or role != Qt.EditRole or not index.isValid():
            return False
        row = self._aligned_row(index.row())
        cell_index = index.column() + 1
        cell = row.merged_row.cell_at(cell_index)
        if cell is None:
            return False
        cell.value = str(value)
        row.conflict_columns.discard(cell_index)
        if row.status == "conflict" and not row.conflict_columns:
            row.status = "same"
            row.reason = "已手工修正"
        self.invalidate_diff_cache(row)
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def source_row_index(self, visible_row: int) -> int:
        return self.visible_rows[visible_row]

    def _aligned_row(self, visible_row: int) -> AlignedRow:
        return self.alignment.rows[self.visible_rows[visible_row]]

    def _cell_value(self, row: AlignedRow, logical_index: int, column) -> str:
        if self.side == "middle":
            return row.merged_row.value_at(logical_index + 1)
        if self.side == "left":
            return row.left_row.value_at(column.left_index or -1) if row.left_row and column.left_index else ""
        return row.right_row.value_at(column.right_index or -1) if row.right_row and column.right_index else ""

    def _background_for_cell(self, row: AlignedRow, logical_column: int, column) -> QBrush:
        if row.status == "deleted":
            return SIMPLIFIED_CELL_BRUSHES["deleted"]
        if self.side == "middle" and logical_column in row.conflict_columns:
            return SIMPLIFIED_CELL_BRUSHES["middle_conflict"]
        if self.side == "left" and logical_column in row.conflict_columns:
            return SIMPLIFIED_CELL_BRUSHES["left_conflict"]
        if self.side == "right" and logical_column in row.conflict_columns:
            return SIMPLIFIED_CELL_BRUSHES["right_conflict"]
        if row.merged_row.kind in {"header", "group", "note"}:
            return SIMPLIFIED_CELL_BRUSHES["same"]
        if self._cell_has_difference(row, logical_column, column):
            return SIMPLIFIED_CELL_BRUSHES["changed"]
        return SIMPLIFIED_CELL_BRUSHES["same"]

    def _row_text(self, row: AlignedRow) -> str:
        parts: list[str] = []
        seen: set[str] = set()
        for row_data in (row.left_row, row.merged_row, row.right_row, row.base_row):
            if row_data is None:
                continue
            for cell in row_data.cells:
                value = str(cell.value or "")
                if not value or value in seen:
                    continue
                seen.add(value)
                parts.append(value)
        return " ".join(parts)

    def _row_text_lower(self, row: AlignedRow) -> str:
        key = id(row)
        cached = self._row_text_lower_cache.get(key)
        if cached is not None:
            return cached
        result = self._row_text(row).lower()
        self._row_text_lower_cache[key] = result
        return result

    def _row_has_difference(self, row: AlignedRow) -> bool:
        key = id(row)
        cached = self._row_diff_cache.get(key)
        if cached is not None:
            return cached
        result = self._compute_row_difference(row)
        self._row_diff_cache[key] = result
        return result

    def _compute_row_difference(self, row: AlignedRow) -> bool:
        if row.status in {"conflict", "left_only", "right_only", "deleted"}:
            return True
        for logical_index, binding in enumerate(self.alignment.columns, start=1):
            left_value = row.left_row.value_at(binding.left_index or -1) if row.left_row and binding.left_index else ""
            right_value = row.right_row.value_at(binding.right_index or -1) if row.right_row and binding.right_index else ""
            if not compare_text_values(left_value, right_value, self.comparison_options) or logical_index in row.conflict_columns:
                return True
        return False

    def _cell_has_difference(self, row: AlignedRow, logical_column: int, column) -> bool:
        if row.merged_row.kind in {"header", "group", "note"}:
            return False
        key = (id(row), logical_column)
        cached = self._cell_diff_cache.get(key)
        if cached is not None:
            return cached
        if logical_column in row.conflict_columns:
            self._cell_diff_cache[key] = True
            return True
        left_value = row.left_row.value_at(column.left_index or -1) if row.left_row and column.left_index else ""
        right_value = row.right_row.value_at(column.right_index or -1) if row.right_row and column.right_index else ""
        result = not compare_text_values(left_value, right_value, self.comparison_options)
        self._cell_diff_cache[key] = result
        return result

    def _row_header_marker(self, row: AlignedRow) -> str:
        if row.status == "conflict":
            return " !"
        side_state = self._row_header_side_state(row)
        if side_state == "right_only":
            return " +"
        if side_state == "left_only_or_deleted":
            return " -"
        if row.status in {"left_only", "right_only", "deleted"}:
            return ""
        if self._row_has_difference(row):
            return " ~"
        return ""

    def _row_header_background_brush(self, row: AlignedRow) -> QBrush:
        if row.status == "conflict":
            return _ROW_HEADER_BACKGROUND_BRUSHES["conflict"]
        side_state = self._row_header_side_state(row)
        if side_state != "same":
            return _ROW_HEADER_BACKGROUND_BRUSHES[side_state]
        if row.status in {"left_only", "right_only", "deleted"}:
            return _ROW_HEADER_BACKGROUND_BRUSHES["same"]
        if self._row_has_difference(row):
            return _ROW_HEADER_BACKGROUND_BRUSHES["changed"]
        return _ROW_HEADER_BACKGROUND_BRUSHES["same"]

    def _row_header_foreground_brush(self, row: AlignedRow) -> QBrush:
        if row.status == "conflict":
            return _ROW_HEADER_FOREGROUND_BRUSHES["conflict"]
        side_state = self._row_header_side_state(row)
        if side_state != "same":
            return _ROW_HEADER_FOREGROUND_BRUSHES[side_state]
        if row.status in {"left_only", "right_only", "deleted"}:
            return _ROW_HEADER_FOREGROUND_BRUSHES["same"]
        if self._row_has_difference(row):
            return _ROW_HEADER_FOREGROUND_BRUSHES["changed"]
        return _ROW_HEADER_FOREGROUND_BRUSHES["same"]

    def _row_header_side_state(self, row: AlignedRow) -> str:
        if row.status == "right_only":
            return "right_only" if self.side in {"right", "middle"} else "same"
        if row.status in {"left_only", "deleted"}:
            return "left_only_or_deleted" if self.side in {"left", "middle"} else "same"
        return "same"

    def _row_header_has_difference(self, row: AlignedRow) -> bool:
        if row.status in {"left_only", "right_only", "deleted"} and self._row_header_side_state(row) == "same":
            return False
        return self._row_has_difference(row)

    def _header_title(self, column) -> str:
        key_prefix = "ID " if normalize_header(column.key) in set(self.alignment.key_fields or []) else ""
        if column.left_index is None:
            return f"{key_prefix}{column.title} [+右列]"
        if column.right_index is None:
            return f"{key_prefix}{column.title} [-右列]"
        return f"{key_prefix}{column.title}"

    def _header_tooltip(self, column) -> str:
        key_fields = set(self.alignment.key_fields or [])
        key_line = "当前子表组合 ID 字段之一。\n" if normalize_header(column.key) in key_fields else ""
        if column.left_index is None:
            return f"{column.title}\n{key_line}该列仅存在于右侧，可视为右侧新增列。"
        if column.right_index is None:
            return f"{column.title}\n{key_line}该列仅存在于左侧，右侧缺失此列。"
        return f"{column.title}\n{key_line}".strip()


class SvnRepositoryBrowserDialog(QDialog):
    def __init__(self, current_root: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.selected_path = ""
        self._loaded_paths: set[str] = set()
        self._entries_cache: dict[str, list[SourceBrowseEntry]] = {}

        self.setWindowTitle("选择 SVN 目录")
        self.resize(980, 660)

        self.url_input = QLineEdit(current_root)
        self.url_input.setPlaceholderText("输入 svn://... 或 svn:... 路径")
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["目录"])
        self.content = QTreeWidget()
        self.content.setHeaderLabels(["名称", "类型", "Revision", "Author", "Date"])
        self.current_path_label = QLabel("<span style='color: #94a3b8;'>尚未选择任何目录</span>")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("URL"))
        top_row.addWidget(self.url_input, 1)
        refresh_button = QPushButton("读取目录")
        refresh_button.clicked.connect(self.reload_root)
        top_row.addWidget(refresh_button)
        layout.addLayout(top_row)
        layout.addWidget(self.current_path_label)

        splitter = QSplitter()
        splitter.addWidget(self.tree)
        splitter.addWidget(self.content)
        splitter.setSizes([360, 560])
        layout.addWidget(splitter, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.tree.itemExpanded.connect(self._on_tree_expanded)
        self.tree.currentItemChanged.connect(self._on_tree_selection_changed)
        self.content.itemDoubleClicked.connect(self._on_content_double_clicked)

        if current_root.strip():
            self.reload_root()

    def reload_root(self) -> None:
        root = self.root_path()
        if not root:
            QMessageBox.information(self, "提示", "请先输入 SVN 根路径。")
            return
        try:
            normalized = normalize_source_target(root)
            entries = list_svn_directory(normalized)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 目录失败", str(exc))
            return

        self.url_input.setText(normalized)
        self.selected_path = normalized
        self.current_path_label.setText(normalized)
        self._entries_cache = {normalized: entries}
        self._loaded_paths = {normalized}
        self.tree.clear()
        root_item = QTreeWidgetItem([source_path_name(normalized) or normalized])
        root_item.setData(0, Qt.UserRole, normalized)
        self.tree.addTopLevelItem(root_item)
        self._populate_tree_children(root_item, entries)
        root_item.setExpanded(True)
        self.tree.setCurrentItem(root_item)

    def root_path(self) -> str:
        return self.url_input.text().strip()

    def accept(self) -> None:
        item = self.tree.currentItem()
        if item is None:
            QMessageBox.information(self, "提示", "请先选中一个 SVN 目录。")
            return
        path = str(item.data(0, Qt.UserRole) or "").strip()
        if not path:
            QMessageBox.information(self, "提示", "当前目录无效，请重新选择。")
            return
        self.selected_path = path
        super().accept()

    def _populate_tree_children(self, parent_item: QTreeWidgetItem, entries: list[SourceBrowseEntry]) -> None:
        parent_item.takeChildren()
        for entry in entries:
            if entry.kind != "dir":
                continue
            child = QTreeWidgetItem([entry.name])
            child.setData(0, Qt.UserRole, entry.path)
            child.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
            child.addChild(QTreeWidgetItem(["加载中..."]))
            parent_item.addChild(child)

    def _on_tree_expanded(self, item: QTreeWidgetItem) -> None:
        path = str(item.data(0, Qt.UserRole) or "").strip()
        if not path or path in self._loaded_paths:
            return
        try:
            entries = list_svn_directory(path)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 目录失败", str(exc))
            return
        self._entries_cache[path] = entries
        self._loaded_paths.add(path)
        self._populate_tree_children(item, entries)

    def _on_tree_selection_changed(self, current: QTreeWidgetItem | None, previous: QTreeWidgetItem | None) -> None:
        del previous
        if current is None:
            self.content.clear()
            self.current_path_label.setText("<span style='color: #94a3b8;'>尚未选择任何目录</span>")
            return
        path = str(current.data(0, Qt.UserRole) or "").strip()
        if not path:
            return
        self.selected_path = path
        self.current_path_label.setText(path)
        entries = self._entries_cache.get(path)
        if entries is None:
            try:
                entries = list_svn_directory(path)
            except Exception as exc:
                QMessageBox.warning(self, "读取 SVN 目录失败", str(exc))
                return
            self._entries_cache[path] = entries
            self._loaded_paths.add(path)
            self._populate_tree_children(current, entries)
        self._populate_content(entries)

    def _populate_content(self, entries: list[SourceBrowseEntry]) -> None:
        self.content.clear()
        for entry in entries:
            kind_label = "目录" if entry.kind == "dir" else "文件"
            item = QTreeWidgetItem([entry.name, kind_label, entry.revision, entry.author, entry.date])
            item.setData(0, Qt.UserRole, entry.path)
            item.setData(0, Qt.UserRole + 1, entry.kind)
            self.content.addTopLevelItem(item)
        for column in range(self.content.columnCount()):
            self.content.resizeColumnToContents(column)

    def _on_content_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        del column
        if str(item.data(0, Qt.UserRole + 1) or "") != "dir":
            return
        target_path = str(item.data(0, Qt.UserRole) or "").strip()
        if not target_path:
            return
        target_item = self._find_tree_item(target_path)
        if target_item is None:
            return
        self.tree.setCurrentItem(target_item)
        target_item.setExpanded(True)

    def _find_tree_item(self, target_path: str) -> QTreeWidgetItem | None:
        def walk(item: QTreeWidgetItem) -> QTreeWidgetItem | None:
            if str(item.data(0, Qt.UserRole) or "") == target_path:
                return item
            for index in range(item.childCount()):
                found = walk(item.child(index))
                if found is not None:
                    return found
            return None

        for index in range(self.tree.topLevelItemCount()):
            found = walk(self.tree.topLevelItem(index))
            if found is not None:
                return found
        return None


class DocumentPickerDialog(QDialog):
    def __init__(
        self,
        side_label: str,
        source_kind: str,
        current_root: str,
        current_file: str,
        current_revision: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.source_kind = source_kind
        self.selected_entry: SourceFileEntry | None = None
        self.setWindowTitle(f"选择{side_label}文档")
        self.resize(920, 640)

        self.root_input = QLineEdit(current_root)
        self.root_input.setPlaceholderText("本地填目录，SVN 填仓库或工作副本路径")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索文件名")
        self.file_list = QListWidget()
        self.file_list.itemDoubleClicked.connect(lambda _: self.accept())
        self.summary_label = QLabel("<h3 style='color: #94a3b8;'><br/>🌿 请选择一个 XML 文件进行比对</h3>")
        self.revision_list = QListWidget()
        self.revision_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.revision_detail = QTextEdit()
        self.revision_detail.setReadOnly(True)
        self.revision_detail.setMinimumHeight(120)
        self._all_entries: list[SourceFileEntry] = []
        self._revision_entries: list[RevisionEntry] = []
        self._initial_revision = current_revision or ("HEAD" if source_kind == SOURCE_SVN else "WORKING")
        self._preferred_file = current_file
        self._auto_refresh_timer = QTimer(self)
        self._auto_refresh_timer.setSingleShot(True)
        self._auto_refresh_timer.timeout.connect(self._refresh_entries_from_timer)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        head_row = QHBoxLayout()
        head_row.addWidget(QLabel("来源"))
        source_label = QLabel(SOURCE_LABELS.get(source_kind, source_kind.upper()))
        head_row.addWidget(source_label)
        head_row.addSpacing(12)
        head_row.addWidget(QLabel("根路径"))
        head_row.addWidget(self.root_input, 1)
        if source_kind == SOURCE_LOCAL:
            browse_button = QPushButton("选择目录")
            browse_button.clicked.connect(self._choose_root_folder)
            head_row.addWidget(browse_button)
        else:
            browse_button = QPushButton("浏览仓库")
            browse_button.clicked.connect(self._browse_svn_root)
            head_row.addWidget(browse_button)
        refresh_button = QPushButton("读取文件")
        refresh_button.clicked.connect(self.refresh_entries)
        head_row.addWidget(refresh_button)
        layout.addLayout(head_row)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("搜索"))
        filter_row.addWidget(self.search_input, 1)
        if source_kind == SOURCE_SVN:
            history_button = QPushButton("读取历史")
            history_button.clicked.connect(self._load_revision_history)
            filter_row.addWidget(history_button)
            filter_row.addWidget(QLabel("双击左侧文件后自动读取提交记录；可按 Shift 连选一个提交区间"))
        else:
            filter_row.addWidget(QLabel("本地以当前文件为准"))
        layout.addLayout(filter_row)
        layout.addWidget(self.summary_label)
        content_row = QHBoxLayout()
        content_row.setSpacing(8)
        content_row.addWidget(self.file_list, 3)
        if source_kind == SOURCE_SVN:
            revision_panel = QVBoxLayout()
            revision_panel.setSpacing(8)
            revision_panel.addWidget(QLabel("提交记录"))
            revision_panel.addWidget(self.revision_list, 2)
            revision_panel.addWidget(QLabel("提交说明"))
            revision_panel.addWidget(self.revision_detail, 1)
            content_row.addLayout(revision_panel, 2)
        layout.addLayout(content_row, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.search_input.textChanged.connect(self._apply_filter)
        self.file_list.currentItemChanged.connect(self._on_selection_changed)
        self.revision_list.itemSelectionChanged.connect(self._on_revision_selection_changed)
        self.root_input.textChanged.connect(self._schedule_auto_refresh)
        QTimer.singleShot(0, self._schedule_auto_refresh)

    def refresh_entries(self, checked: bool = False, preferred_file: str = "") -> None:
        del checked
        root = self.root().strip()
        if not root:
            self.summary_label.setText("<h3 style='color: #f87171;'>⚠️ 请先输入有效根路径</h3>")
            self._all_entries = []
            self.file_list.clear()
            return
        try:
            if self.source_kind == SOURCE_LOCAL:
                self._all_entries = list_local_table_files(root)
            else:
                self._all_entries = list_svn_xml_files(root)
        except Exception as exc:
            QMessageBox.warning(self, "读取文件失败", str(exc))
            self.summary_label.setText("读取文件失败。")
            self._all_entries = []
            self.file_list.clear()
            return
        kind_label = "表格文件" if self.source_kind == SOURCE_LOCAL else "XML 文件"
        self.summary_label.setText(f"共找到 {len(self._all_entries)} 个{kind_label}")
        self._apply_filter(preferred_name=preferred_file)

    def _schedule_auto_refresh(self) -> None:
        self._auto_refresh_timer.start(120)

    def _refresh_entries_from_timer(self) -> None:
        self.refresh_entries(preferred_file=self._preferred_file)
        self._preferred_file = ""

    def root(self) -> str:
        return self.root_input.text().strip()

    def selected_file_path(self) -> str:
        return self.selected_entry.path if self.selected_entry is not None else ""

    def selected_revision(self) -> str:
        if self.source_kind != SOURCE_SVN:
            return "WORKING"
        revisions = self._selected_revision_values()
        if not revisions:
            return "HEAD"
        value = revisions[0]
        return value or "HEAD"

    def selected_revision_range(self) -> tuple[str, str] | None:
        if self.source_kind != SOURCE_SVN:
            return None
        selected_rows = self._selected_revision_rows()
        if len(selected_rows) <= 1:
            return None
        if selected_rows != list(range(selected_rows[0], selected_rows[-1] + 1)):
            raise RuntimeError("SVN diff 区间必须是连续的提交记录，请用 Shift 选择一个连续区块。")
        selected_revisions = self._selected_revision_values()
        if len(selected_revisions) <= 1:
            return None
        newer = selected_revisions[0] or "HEAD"
        older = selected_revisions[-1] or "HEAD"
        return older, newer

    def accept(self) -> None:
        item = self.file_list.currentItem()
        if item is None:
            QMessageBox.information(self, "提示", "请先选中一个文件。")
            return
        if self.source_kind == SOURCE_SVN:
            try:
                self.selected_revision_range()
            except RuntimeError as exc:
                QMessageBox.information(self, "提示", str(exc))
                return
        entry = item.data(Qt.UserRole)
        if not isinstance(entry, SourceFileEntry):
            QMessageBox.information(self, "提示", "当前选项无效，请重新选择。")
            return
        self.selected_entry = entry
        super().accept()

    def _choose_root_folder(self) -> None:
        start_dir = self.root() or str(Path.cwd())
        folder = QFileDialog.getExistingDirectory(self, "选择本地目录", start_dir)
        if not folder:
            return
        self.root_input.setText(folder)
        self.refresh_entries()

    def _browse_svn_root(self) -> None:
        dialog = SvnRepositoryBrowserDialog(self.root(), self)
        if dialog.exec() != QDialog.Accepted or not dialog.selected_path:
            return
        self.root_input.setText(dialog.selected_path)
        self.refresh_entries()

    def _apply_filter(self, text: str = "", preferred_name: str = "") -> None:
        keyword = (text or self.search_input.text()).strip().lower()
        current_path = self.selected_entry.path if self.selected_entry is not None else ""
        self.file_list.blockSignals(True)
        self.file_list.clear()
        target_row = 0
        visible_entries = []
        for entry in self._all_entries:
            if keyword and keyword not in entry.path.lower() and keyword not in entry.name.lower():
                continue
            visible_entries.append(entry)
        for row_index, entry in enumerate(visible_entries):
            item_text = source_relative_path(self.root(), entry.path) if self.source_kind == SOURCE_SVN else entry.name
            item = QListWidgetItem(item_text)
            item.setToolTip(entry.path)
            item.setData(Qt.UserRole, entry)
            self.file_list.addItem(item)
            if preferred_name and (
                entry.name == preferred_name
                or item_text == preferred_name
                or entry.path.endswith(preferred_name)
            ):
                target_row = row_index
            elif current_path and entry.path == current_path:
                target_row = row_index
        self.file_list.blockSignals(False)
        if self.file_list.count():
            self.file_list.setCurrentRow(target_row)
        else:
            self.selected_entry = None
            self.summary_label.setText("<h3 style='color: #94a3b8;'>🪹 没有匹配的文件</h3>")

    def _on_selection_changed(self, current: QListWidgetItem | None, previous: QListWidgetItem | None) -> None:
        del previous
        if current is None:
            self.selected_entry = None
            return
        entry = current.data(Qt.UserRole)
        if isinstance(entry, SourceFileEntry):
            self.selected_entry = entry
            self.summary_label.setText(f"已选中：{entry.path}")
            if self.source_kind == SOURCE_SVN:
                self._load_revision_history()

    def _load_revision_history(self) -> None:
        if self.source_kind != SOURCE_SVN or self.selected_entry is None:
            return
        try:
            entries = list_svn_revisions(self.selected_entry.path)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 历史失败", str(exc))
            return
        current_revision = self.selected_revision() if self._revision_entries else self._initial_revision
        self._revision_entries = entries
        self.revision_list.blockSignals(True)
        self.revision_list.clear()
        target_row = 0
        for entry in entries:
            label = "HEAD | 最新版本" if entry.revision == "HEAD" else f"r{entry.revision} | {entry.author} | {entry.date}"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, entry.revision)
            item.setToolTip(entry.message)
            self.revision_list.addItem(item)
            if entry.revision == current_revision:
                target_row = self.revision_list.count() - 1
        self.revision_list.blockSignals(False)
        if self.revision_list.count():
            self.revision_list.setCurrentRow(target_row)
        else:
            self.revision_detail.clear()

    def _on_revision_selection_changed(self) -> None:
        selected_rows = self._selected_revision_rows()
        if not selected_rows:
            self.revision_detail.clear()
            return
        if len(selected_rows) > 1:
            if selected_rows != list(range(selected_rows[0], selected_rows[-1] + 1)):
                self.revision_detail.setPlainText("当前选择不是连续区间。\n\n请按住 Shift 选择连续提交记录。")
                return
            selected_revisions = self._selected_revision_values()
            newer = selected_revisions[0] or "HEAD"
            older = selected_revisions[-1] or "HEAD"
            detail_lines = [
                "SVN diff 区间",
                "",
                f"新版本: {newer}",
                f"旧版本: {older}",
                f"提交数量: {len(selected_rows)}",
                "",
                "确定后会自动映射为：",
                f"左侧 = {older}",
                f"右侧 = {newer}",
            ]
            self.revision_detail.setPlainText("\n".join(detail_lines))
            return
        revision = self._selected_revision_values()[0]
        matched = next((item for item in self._revision_entries if item.revision == revision), None)
        if matched is None:
            self.revision_detail.clear()
            return
        if matched.revision == "HEAD":
            self.revision_detail.setPlainText("HEAD\n\n最新版本，不对应单独提交说明。")
            return
        detail_lines = [
            f"Revision: r{matched.revision}",
            f"Author: {matched.author}",
            f"Date: {matched.date}",
            "",
            matched.message or "(无提交说明)",
        ]
        self.revision_detail.setPlainText("\n".join(detail_lines))

    def _selected_revision_rows(self) -> list[int]:
        selected_indexes = self.revision_list.selectionModel().selectedRows() if self.revision_list.selectionModel() else []
        return sorted(index.row() for index in selected_indexes)

    def _selected_revision_values(self) -> list[str]:
        values: list[str] = []
        for row in self._selected_revision_rows():
            item = self.revision_list.item(row)
            if item is None:
                continue
            values.append(str(item.data(Qt.UserRole) or "").strip() or "HEAD")
        return values


class QuickRootManagerDialog(QDialog):
    def __init__(
        self,
        entries: list[dict],
        left_root: str,
        right_root: str,
        google_auth_mode: str,
        google_service_account_path: str,
        google_oauth_client_path: str,
        google_oauth_token_path: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.entries = [{"name": str(item.get("name") or "").strip(), "path": str(item.get("path") or "").strip()} for item in entries]
        self._current_index = -1
        self._initial_google_settings = {
            "google_auth_mode": str(google_auth_mode or "service_account"),
            "google_service_account_path": str(google_service_account_path or "").strip(),
            "google_oauth_client_path": str(google_oauth_client_path or "").strip(),
            "google_oauth_token_path": str(google_oauth_token_path or "").strip(),
        }

        self.setWindowTitle("管理快捷配置")
        self.resize(860, 520)
        self.setMinimumSize(760, 500)

        self.entry_list = QListWidget()
        self.entry_list.setObjectName("quickEntryList")
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("例如：主干 core / 本地 normal")
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("输入本地目录 / SVN 路径 / Google Sheets URL 或 ID")
        self.preview_text = QTextEdit()
        self.preview_text.setObjectName("quickPreviewText")
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(104)
        self.preview_text.setMaximumHeight(148)
        self.preview_text.setPlaceholderText("当前快捷项摘要会显示在这里。")
        self.google_auth_mode_combo = QComboBox()
        self.google_auth_mode_combo.addItem("服务账号", "service_account")
        self.google_auth_mode_combo.addItem("个人登录", "oauth_user")
        self.google_auth_mode_combo.setCurrentIndex(
            max(0, self.google_auth_mode_combo.findData(str(google_auth_mode or "service_account")))
        )
        self.google_account_input = QLineEdit(str(google_service_account_path or "").strip())
        self.google_account_input.setPlaceholderText("可选：google_service_account.json，用于读取私有 Google Sheets")
        self.google_oauth_client_input = QLineEdit(str(google_oauth_client_path or "").strip())
        self.google_oauth_client_input.setPlaceholderText("可选：OAuth 客户端 credentials.json，仅在需要重新登录时使用")
        self.google_oauth_token_input = QLineEdit(str(google_oauth_token_path or "").strip() or default_google_oauth_token_path())
        self.google_oauth_token_input.setPlaceholderText("可直接选择现成的 OAuth token.json，例如 AutoXlsxtoXml 的 token.json")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        content_row = QHBoxLayout()
        content_row.setSpacing(10)

        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("快捷项"))
        left_panel.addWidget(self.entry_list, 1)
        left_actions = QHBoxLayout()
        new_button = QPushButton("新建")
        new_button.clicked.connect(self._new_entry)
        left_actions.addWidget(new_button)
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(self._delete_entry)
        left_actions.addWidget(delete_button)
        left_panel.addLayout(left_actions)
        content_row.addLayout(left_panel, 2)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.NoFrame)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        right_content = QWidget()
        right_content.setObjectName("quickRightContent")
        right_panel = QVBoxLayout(right_content)
        right_panel.setContentsMargins(0, 0, 8, 0)
        right_panel.setSpacing(10)
        form_grid = QGridLayout()
        form_grid.setHorizontalSpacing(8)
        form_grid.setVerticalSpacing(6)
        form_grid.setColumnMinimumWidth(0, 82)
        form_grid.setColumnStretch(1, 1)
        form_grid.setColumnStretch(2, 1)
        form_grid.setColumnMinimumWidth(3, 118)
        form_grid.addWidget(QLabel("名称"), 0, 0)
        form_grid.addWidget(self.name_input, 0, 1, 1, 3)
        form_grid.addWidget(QLabel("路径"), 1, 0)
        form_grid.addWidget(self.path_input, 1, 1, 1, 3)
        browse_local = QPushButton("本地...")
        browse_local.clicked.connect(self._browse_local_root)
        form_grid.addWidget(browse_local, 2, 1)
        browse_svn = QPushButton("SVN...")
        browse_svn.clicked.connect(self._browse_svn_root)
        form_grid.addWidget(browse_svn, 2, 2)
        save_button = QPushButton("保存当前项")
        save_button.clicked.connect(self._save_entry)
        form_grid.addWidget(save_button, 2, 3)
        use_left_button = QPushButton("取左侧当前路径")
        use_left_button.clicked.connect(lambda: self._use_current_root(left_root))
        form_grid.addWidget(use_left_button, 3, 1)
        use_right_button = QPushButton("取右侧当前路径")
        use_right_button.clicked.connect(lambda: self._use_current_root(right_root))
        form_grid.addWidget(use_right_button, 3, 2)
        form_grid.addWidget(QLabel("Google 认证"), 4, 0)
        form_grid.addWidget(self.google_auth_mode_combo, 4, 1)
        self.google_test_button = QPushButton("测试连接")
        self.google_test_button.clicked.connect(self._test_google_connection)
        form_grid.addWidget(self.google_test_button, 4, 2)
        self.google_oauth_login_button = QPushButton("Google 登录")
        self.google_oauth_login_button.clicked.connect(self._login_google_oauth)
        form_grid.addWidget(self.google_oauth_login_button, 4, 3)
        form_grid.addWidget(QLabel("服务账号"), 5, 0)
        form_grid.addWidget(self.google_account_input, 5, 1, 1, 2)
        self.google_service_account_button = QPushButton("选择JSON...")
        self.google_service_account_button.clicked.connect(self._browse_google_service_account)
        form_grid.addWidget(self.google_service_account_button, 5, 3)
        form_grid.addWidget(QLabel("OAuth 客户端"), 6, 0)
        form_grid.addWidget(self.google_oauth_client_input, 6, 1, 1, 2)
        self.google_oauth_client_button = QPushButton("选择JSON...")
        self.google_oauth_client_button.clicked.connect(self._browse_google_oauth_client)
        form_grid.addWidget(self.google_oauth_client_button, 6, 3)
        form_grid.addWidget(QLabel("OAuth Token"), 7, 0)
        form_grid.addWidget(self.google_oauth_token_input, 7, 1, 1, 2)
        self.google_oauth_token_button = QPushButton("保存到...")
        self.google_oauth_token_button.clicked.connect(self._browse_google_oauth_token)
        form_grid.addWidget(self.google_oauth_token_button, 7, 3)
        preview_title = QLabel("当前配置摘要")
        preview_title.setObjectName("dialogSectionTitle")
        right_panel.addLayout(form_grid)
        right_panel.addWidget(preview_title)
        right_panel.addWidget(self.preview_text)
        right_panel.addStretch(1)
        right_scroll.setWidget(right_content)
        content_row.addWidget(right_scroll, 3)

        layout.addLayout(content_row, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._apply_dialog_style()

        self.entry_list.currentRowChanged.connect(self._load_entry)
        self.path_input.textChanged.connect(self._update_preview)
        self.name_input.textChanged.connect(self._update_preview)
        self.google_auth_mode_combo.currentIndexChanged.connect(self._update_auth_inputs)
        self.google_auth_mode_combo.currentIndexChanged.connect(self._update_preview)
        self.google_account_input.textChanged.connect(self._update_preview)
        self.google_oauth_client_input.textChanged.connect(self._update_preview)
        self.google_oauth_token_input.textChanged.connect(self._update_preview)
        self._refresh_entry_list()
        self._update_auth_inputs()
        if self.entry_list.count():
            self.entry_list.setCurrentRow(0)
        else:
            self._new_entry()

    def _apply_dialog_style(self) -> None:
        self.setStyleSheet(
            """
            QDialog {
                background: #f5f7fb;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                color: #111827;
            }
            QLabel {
                color: #111827;
                font-weight: 600;
            }
            QLabel#dialogSectionTitle {
                color: #475569;
                font-size: 12px;
                font-weight: 800;
                padding-top: 2px;
            }
            QListWidget#quickEntryList, QTextEdit#quickPreviewText {
                background: #ffffff;
                border: 1px solid #dfe7f1;
                border-radius: 12px;
                color: #111827;
            }
            QTextEdit#quickPreviewText {
                padding: 6px;
            }
            QListWidget#quickEntryList::item {
                min-height: 24px;
                padding: 4px 8px;
                border-radius: 8px;
                border: 1px solid transparent;
            }
            QListWidget#quickEntryList::item:hover {
                background: #f1f5f9;
            }
            QListWidget#quickEntryList::item:selected {
                background: #dbeafe;
                color: #1d4ed8;
                border: 1px solid #bfdbfe;
            }
            QScrollArea {
                background: transparent;
                border: none;
            }
            QWidget#quickRightContent {
                background: transparent;
            }
            QLineEdit, QComboBox {
                min-height: 30px;
                padding: 0 10px;
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                color: #111827;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid #2563eb;
            }
            QPushButton {
                min-height: 30px;
                padding: 0 14px;
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                color: #111827;
                font-weight: 700;
            }
            QPushButton:hover {
                background: #eff6ff;
                border-color: #93c5fd;
                color: #1d4ed8;
            }
            QPushButton:disabled {
                background: #f1f5f9;
                border-color: #e2e8f0;
                color: #94a3b8;
            }
            """
        )

    def accept(self) -> None:
        if self.name_input.text().strip() or self.path_input.text().strip():
            if not self._save_entry(show_message=False):
                return
        super().accept()

    def result_entries(self) -> list[dict]:
        return [dict(item) for item in self.entries]

    def result_google_service_account_path(self) -> str:
        return self.google_account_input.text().strip()

    def result_google_settings(self) -> dict[str, str]:
        return {
            "google_auth_mode": str(self.google_auth_mode_combo.currentData() or "service_account"),
            "google_service_account_path": self.google_account_input.text().strip(),
            "google_oauth_client_path": self.google_oauth_client_input.text().strip(),
            "google_oauth_token_path": self.google_oauth_token_input.text().strip(),
        }

    def _current_google_settings(self) -> dict[str, str]:
        return {
            "google_auth_mode": str(self.google_auth_mode_combo.currentData() or "service_account"),
            "google_service_account_path": self.google_account_input.text().strip(),
            "google_oauth_client_path": self.google_oauth_client_input.text().strip(),
            "google_oauth_token_path": self.google_oauth_token_input.text().strip(),
        }

    def _refresh_entry_list(self) -> None:
        current_path = ""
        if 0 <= self._current_index < len(self.entries):
            current_path = self.entries[self._current_index]["path"]
        self.entry_list.blockSignals(True)
        self.entry_list.clear()
        current_row = -1
        name_counts: dict[str, int] = {}
        for entry in self.entries:
            name = entry["name"] or "(未命名)"
            name_counts[name] = name_counts.get(name, 0) + 1
        for row, entry in enumerate(self.entries):
            name = entry["name"] or "(未命名)"
            item_label = name
            if name_counts.get(name, 0) > 1:
                item_label = f"{name}  ·  {self._compact_entry_path(entry['path'])}"
            item = QListWidgetItem(item_label)
            item.setToolTip(entry["path"])
            self.entry_list.addItem(item)
            if entry["path"] == current_path:
                current_row = row
        self.entry_list.blockSignals(False)
        if current_row >= 0:
            self.entry_list.setCurrentRow(current_row)

    @staticmethod
    def _compact_entry_path(path: str) -> str:
        normalized = str(path or "").replace("\\", "/").rstrip("/")
        if not normalized:
            return "(空路径)"
        if is_google_sheets_target(normalized):
            try:
                parsed = parse_google_sheet_target(normalized)
            except Exception:  # noqa: BLE001
                return "Google Sheets"
            gid = f"#{parsed['gid']}" if parsed.get("gid") else ""
            return f"Google/{parsed['spreadsheet_id'][:10]}{gid}"
        parts = [part for part in normalized.split("/") if part]
        if len(parts) >= 3:
            return "/".join(parts[-3:])
        return "/".join(parts) or normalized

    def _load_entry(self, row: int) -> None:
        self._current_index = row
        if row < 0 or row >= len(self.entries):
            self.name_input.clear()
            self.path_input.clear()
            self._update_preview()
            return
        entry = self.entries[row]
        self.name_input.setText(entry["name"])
        self.path_input.setText(entry["path"])
        self._update_preview()

    def _new_entry(self) -> None:
        self._current_index = -1
        self.entry_list.clearSelection()
        self.name_input.clear()
        self.path_input.clear()
        self._update_preview()
        self.name_input.setFocus()

    def _save_entry(self, show_message: bool = True) -> bool:
        path = self.path_input.text().strip()
        name = self.name_input.text().strip()
        if not path:
            if show_message:
                QMessageBox.information(self, "提示", "请先填写一个路径。")
            return False
        if not name:
            if is_google_sheets_target(path):
                parsed = parse_google_sheet_target(path)
                gid = f"#{parsed['gid']}" if parsed["gid"] else ""
                name = f"Google/{parsed['spreadsheet_id'][:10]}{gid}"
            else:
                normalized = path.replace("\\", "/").rstrip("/")
                parts = [part for part in normalized.split("/") if part]
                name = "/".join(parts[-2:]) if len(parts) >= 2 else (parts[-1] if parts else path)
        entry = {"name": name, "path": path}
        duplicate_index = next((idx for idx, item in enumerate(self.entries) if item["path"] == path), -1)
        if duplicate_index >= 0 and duplicate_index != self._current_index:
            self.entries[duplicate_index] = entry
            self._current_index = duplicate_index
        elif 0 <= self._current_index < len(self.entries):
            self.entries[self._current_index] = entry
        else:
            self.entries.append(entry)
            self._current_index = len(self.entries) - 1
        self._refresh_entry_list()
        self.entry_list.setCurrentRow(self._current_index)
        if show_message:
            self.preview_text.setPlainText(f"已保存快捷项：{name}")
        return True

    def _delete_entry(self) -> None:
        row = self.entry_list.currentRow()
        if row < 0 or row >= len(self.entries):
            QMessageBox.information(self, "提示", "请先选中一个快捷项。")
            return
        del self.entries[row]
        self._refresh_entry_list()
        if self.entries:
            self.entry_list.setCurrentRow(min(row, len(self.entries) - 1))
        else:
            self._new_entry()

    def _browse_local_root(self) -> None:
        start_dir = self.path_input.text().strip() or str(Path.cwd())
        folder = QFileDialog.getExistingDirectory(self, "选择本地目录", start_dir)
        if not folder:
            return
        self.path_input.setText(folder)

    def _browse_svn_root(self) -> None:
        dialog = SvnRepositoryBrowserDialog(self.path_input.text().strip(), self)
        if dialog.exec() != QDialog.Accepted or not dialog.selected_path:
            return
        self.path_input.setText(dialog.selected_path)

    def _browse_google_service_account(self) -> None:
        start_path = self.google_account_input.text().strip() or str(Path.cwd())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 google_service_account.json",
            start_path,
            "JSON 文件 (*.json);;所有文件 (*)",
        )
        if not file_path:
            return
        self.google_account_input.setText(file_path)

    def _browse_google_oauth_client(self) -> None:
        start_path = self.google_oauth_client_input.text().strip() or str(Path.cwd())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 OAuth credentials.json",
            start_path,
            "JSON 文件 (*.json);;所有文件 (*)",
        )
        if not file_path:
            return
        self.google_oauth_client_input.setText(file_path)

    def _browse_google_oauth_token(self) -> None:
        start_path = self.google_oauth_token_input.text().strip() or default_google_oauth_token_path()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择 OAuth token.json 保存位置",
            start_path,
            "JSON 文件 (*.json);;所有文件 (*)",
        )
        if not file_path:
            return
        self.google_oauth_token_input.setText(file_path)

    def _test_google_connection(self) -> None:
        target = self.path_input.text().strip()
        if not target:
            QMessageBox.information(self, "Google 连接测试", "请先填写 Google Sheets URL 或 Spreadsheet ID。")
            return
        if not is_google_sheets_target(target):
            QMessageBox.information(self, "Google 连接测试", "当前路径不是 Google Sheets URL 或 Spreadsheet ID。")
            return

        current_settings = self._current_google_settings()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            set_google_auth_settings(**current_settings)
            info = describe_google_sheet(target)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Google 连接测试失败", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
            set_google_auth_settings(**self._initial_google_settings)

        summary = "\n".join(
            f"- {item['title']} ({item['row_count']} 行 / {item['column_count']} 列)"
            for item in info["sheets"][:8]
        ) or "- 未读取到工作表信息"
        sheet_count = len(info["sheets"])
        more_hint = "\n..." if sheet_count > 8 else ""
        QMessageBox.information(
            self,
            "Google 连接测试成功",
            f"标题：{info['title']}\n"
            f"Spreadsheet ID：{info['spreadsheet_id']}\n"
            f"工作表数量：{sheet_count}\n\n"
            f"{summary}{more_hint}",
        )

    def _login_google_oauth(self) -> None:
        client_path = self.google_oauth_client_input.text().strip()
        token_path = self.google_oauth_token_input.text().strip() or default_google_oauth_token_path()
        self.google_oauth_token_input.setText(token_path)
        try:
            saved_token_path = start_google_oauth_login(
                client_path,
                token_path,
                manual_prompt_callback=self._prompt_google_oauth_manual,
            )
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Google 登录失败", str(exc))
            return
        self.google_auth_mode_combo.setCurrentIndex(max(0, self.google_auth_mode_combo.findData("oauth_user")))
        self.google_oauth_token_input.setText(saved_token_path)
        QMessageBox.information(
            self,
            "Google 登录成功",
            "已完成 Google 授权，后续可以用当前个人账号读取分享给你的 Google Sheets。",
        )

    def _prompt_google_oauth_manual(self, auth_url: str) -> bool:
        QApplication.clipboard().setText(auth_url)
        box = QMessageBox(self)
        box.setWindowTitle("Google 手动授权")
        box.setIcon(QMessageBox.Information)
        box.setText("无法自动打开浏览器，授权链接已复制到剪贴板。")
        box.setInformativeText(
            "请点击“确定”后，把链接粘贴到浏览器打开并完成 Google 登录。\n"
            "程序会继续等待浏览器回调，无需关闭本工具。"
        )
        box.setDetailedText(auth_url)
        box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        ok_button = box.button(QMessageBox.Ok)
        if ok_button is not None:
            ok_button.setText("确定")
        cancel_button = box.button(QMessageBox.Cancel)
        if cancel_button is not None:
            cancel_button.setText("取消")
        return box.exec() == QMessageBox.Ok

    def _update_auth_inputs(self) -> None:
        auth_mode = str(self.google_auth_mode_combo.currentData() or "service_account")
        service_enabled = auth_mode == "service_account"
        oauth_enabled = auth_mode == "oauth_user"
        self.google_account_input.setEnabled(service_enabled)
        self.google_service_account_button.setEnabled(service_enabled)
        self.google_oauth_client_input.setEnabled(oauth_enabled)
        self.google_oauth_client_button.setEnabled(oauth_enabled)
        self.google_oauth_token_input.setEnabled(oauth_enabled)
        self.google_oauth_token_button.setEnabled(oauth_enabled)
        self.google_oauth_login_button.setEnabled(oauth_enabled)

    def _use_current_root(self, root: str) -> None:
        root = str(root or "").strip()
        if not root:
            QMessageBox.information(self, "提示", "当前主界面还没有可用路径。")
            return
        self.path_input.setText(root)

    def _update_preview(self) -> None:
        name = self.name_input.text().strip() or "(未命名)"
        path = self.path_input.text().strip() or "(未填写路径)"
        auth_mode = str(self.google_auth_mode_combo.currentData() or "service_account")
        auth_mode_label = "个人登录" if auth_mode == "oauth_user" else "服务账号"
        google_credentials = self.google_account_input.text().strip() or "(未配置)"
        google_oauth_client = self.google_oauth_client_input.text().strip() or "(未配置)"
        google_oauth_token = self.google_oauth_token_input.text().strip() or "(未配置)"
        if is_google_sheets_target(path):
            parsed = parse_google_sheet_target(path)
            gid_text = parsed["gid"] or "(未指定)"
            self.preview_text.setPlainText(
                "名称：{name}\n类型：Google Sheets\nSpreadsheet ID：{spreadsheet_id}\nGID：{gid}\n路径：{path}\n"
                "Google 认证：{auth_mode}\n服务账号：{service_account}\nOAuth 客户端：{oauth_client}\nOAuth Token：{oauth_token}".format(
                    name=name,
                    spreadsheet_id=parsed["spreadsheet_id"],
                    gid=gid_text,
                    path=path,
                    auth_mode=auth_mode_label,
                    service_account=google_credentials,
                    oauth_client=google_oauth_client,
                    oauth_token=google_oauth_token,
                )
            )
            return
        self.preview_text.setPlainText(
            "名称：{name}\n路径：{path}\nGoogle 认证：{auth_mode}\n服务账号：{service_account}\nOAuth 客户端：{oauth_client}\nOAuth Token：{oauth_token}".format(
                name=name,
                path=path,
                auth_mode=auth_mode_label,
                service_account=google_credentials,
                oauth_client=google_oauth_client,
                oauth_token=google_oauth_token,
            )
        )


class KeyFieldsDialog(QDialog):
    def __init__(
        self,
        candidates: list[str],
        current_fields: list[str],
        parent: QWidget | None = None,
        *,
        single_mode: bool = False,
    ) -> None:
        super().__init__(parent)
        self.single_mode = single_mode
        self.setWindowTitle("选择ID列" if single_mode else "选择关键字段")
        self.resize(460, 520)

        layout = QVBoxLayout(self)
        hint_text = (
            "选择一个字段作为唯一 ID。该字段在左/右/Base 每一侧都必须非空且唯一。"
            if single_mode
            else "勾选用于判断同一行的字段。多个字段会组成组合键；如果组合键仍重复，会按出现顺序一对一匹配。"
        )
        hint = QLabel(hint_text)
        hint.setWordWrap(True)
        layout.addWidget(hint)

        current_set = set(current_fields)
        ordered_candidates: list[str] = []
        seen: set[str] = set()
        for field in [*current_fields, *candidates]:
            normalized = normalize_header(field)
            if normalized and normalized not in seen:
                seen.add(normalized)
                ordered_candidates.append(normalized)

        self.field_list = QListWidget()
        for field in ordered_candidates:
            item = QListWidgetItem(field)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if field in current_set else Qt.Unchecked)
            self.field_list.addItem(item)
        if single_mode:
            self.field_list.itemChanged.connect(self._on_single_item_changed)
        layout.addWidget(self.field_list, 1)

        self.extra_input = QLineEdit()
        self.extra_input.setPlaceholderText(
            "列表里没有时手动填写一个逻辑字段名" if single_mode else "额外字段，逗号分隔；用于列表里没有但你确定存在的逻辑字段"
        )
        layout.addWidget(self.extra_input)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_single_item_changed(self, changed_item: QListWidgetItem) -> None:
        if changed_item.checkState() != Qt.Checked:
            return
        self.field_list.blockSignals(True)
        for index in range(self.field_list.count()):
            item = self.field_list.item(index)
            if item is not changed_item:
                item.setCheckState(Qt.Unchecked)
        self.field_list.blockSignals(False)

    def selected_fields(self) -> list[str]:
        values: list[str] = []
        seen: set[str] = set()
        for index in range(self.field_list.count()):
            item = self.field_list.item(index)
            if item.checkState() != Qt.Checked:
                continue
            normalized = normalize_header(item.text())
            if normalized and normalized not in seen:
                values.append(normalized)
                seen.add(normalized)
        extra_text = self.extra_input.text().strip()
        for delimiter in ("\n", "；", ";", "，"):
            extra_text = extra_text.replace(delimiter, ",")
        for part in extra_text.split(","):
            if not part.strip():
                continue
            normalized = normalize_header(part)
            if normalized and normalized not in seen:
                values.append(normalized)
                seen.add(normalized)
            if self.single_mode and values:
                return values[:1]
        return values


class SvnRevisionPickerDialog(QDialog):
    def __init__(self, svn_target: str, current_revision: str = "HEAD", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.selected_revision = current_revision or "HEAD"
        self.setWindowTitle("选择 SVN 版本")
        self.resize(820, 520)

        layout = QVBoxLayout(self)
        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("SVN路径"))
        self.target_input = QLineEdit(svn_target)
        top_row.addWidget(self.target_input, 1)
        refresh_button = QPushButton("读取历史")
        refresh_button.clicked.connect(self.reload)
        top_row.addWidget(refresh_button)
        layout.addLayout(top_row)

        self.revision_list = QTreeWidget()
        self.revision_list.setHeaderLabels(["Revision", "Author", "Date", "Message"])
        self.revision_list.setColumnWidth(0, 110)
        self.revision_list.setColumnWidth(1, 120)
        self.revision_list.setColumnWidth(2, 160)
        self.revision_list.itemDoubleClicked.connect(lambda *_: self.accept())
        layout.addWidget(self.revision_list, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if svn_target.strip():
            self.reload()

    def reload(self) -> None:
        target = self.target_input.text().strip()
        if not target:
            QMessageBox.information(self, "提示", "请先填写 SVN 路径。")
            return
        try:
            entries = list_svn_revisions(target)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 历史失败", str(exc))
            return
        self.revision_list.clear()
        selected_item: QTreeWidgetItem | None = None
        for entry in entries:
            item = QTreeWidgetItem([entry.revision, entry.author, entry.date, entry.message])
            item.setData(0, Qt.UserRole, entry.revision)
            self.revision_list.addTopLevelItem(item)
            if entry.revision == self.selected_revision:
                selected_item = item
        if selected_item is not None:
            self.revision_list.setCurrentItem(selected_item)
        elif self.revision_list.topLevelItemCount():
            self.revision_list.setCurrentItem(self.revision_list.topLevelItem(0))

    def accept(self) -> None:
        item = self.revision_list.currentItem()
        if item is None:
            QMessageBox.information(self, "提示", "请先选择一个 SVN 版本。")
            return
        self.selected_revision = str(item.data(0, Qt.UserRole) or item.text(0) or "HEAD")
        super().accept()


class BatchCompareDialog(QDialog):
    def __init__(
        self,
        left_root: str,
        right_root: str,
        output_path: str,
        left_revision: str = "HEAD",
        right_revision: str = "HEAD",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("批量对比")
        self.resize(760, 260)

        layout = QVBoxLayout(self)
        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)

        self.left_root_input = QLineEdit(left_root)
        self.right_root_input = QLineEdit(right_root)
        self.left_revision_input = QLineEdit(left_revision or "HEAD")
        self.right_revision_input = QLineEdit(right_revision or "HEAD")
        self.output_path_input = QLineEdit(output_path)
        self.format_combo = QComboBox()
        self.format_combo.addItem("HTML 报告", "html")
        self.format_combo.addItem("Markdown 报告", "md")
        self.format_combo.addItem("CSV 明细", "csv")

        left_browse = QPushButton("浏览...")
        left_browse.clicked.connect(lambda: self._browse_root(self.left_root_input))
        right_browse = QPushButton("浏览...")
        right_browse.clicked.connect(lambda: self._browse_root(self.right_root_input))
        left_revision_pick = QPushButton("选择...")
        left_revision_pick.clicked.connect(lambda: self._pick_revision(self.left_root_input, self.left_revision_input))
        right_revision_pick = QPushButton("选择...")
        right_revision_pick.clicked.connect(lambda: self._pick_revision(self.right_root_input, self.right_revision_input))
        output_browse = QPushButton("保存到...")
        output_browse.clicked.connect(self._browse_output)

        form.addWidget(QLabel("左侧根路径"), 0, 0)
        form.addWidget(self.left_root_input, 0, 1)
        form.addWidget(left_browse, 0, 2)
        form.addWidget(QLabel("右侧根路径"), 1, 0)
        form.addWidget(self.right_root_input, 1, 1)
        form.addWidget(right_browse, 1, 2)
        form.addWidget(QLabel("左侧SVN版本"), 2, 0)
        form.addWidget(self.left_revision_input, 2, 1)
        form.addWidget(left_revision_pick, 2, 2)
        form.addWidget(QLabel("右侧SVN版本"), 3, 0)
        form.addWidget(self.right_revision_input, 3, 1)
        form.addWidget(right_revision_pick, 3, 2)
        form.addWidget(QLabel("报告文件"), 4, 0)
        form.addWidget(self.output_path_input, 4, 1)
        form.addWidget(output_browse, 4, 2)
        form.addWidget(QLabel("报告格式"), 5, 0)
        form.addWidget(self.format_combo, 5, 1)
        layout.addLayout(form)

        hint = QLabel("按同名表格文件批量对比；SVN 根路径会按填写的版本读取该版本下的完整目录内容。")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> dict[str, str]:
        return {
            "left_root": self.left_root_input.text().strip(),
            "right_root": self.right_root_input.text().strip(),
            "left_revision": self.left_revision_input.text().strip() or "HEAD",
            "right_revision": self.right_revision_input.text().strip() or "HEAD",
            "output_path": self.output_path_input.text().strip(),
            "format": str(self.format_combo.currentData() or "html"),
        }

    def _browse_root(self, target: QLineEdit) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择目录", target.text().strip() or str(Path.cwd()))
        if directory:
            target.setText(directory)

    def _pick_revision(self, root_input: QLineEdit, revision_input: QLineEdit) -> None:
        root = root_input.text().strip()
        if not root or infer_source_kind(root) != SOURCE_SVN:
            QMessageBox.information(self, "提示", "请先填写 SVN 根路径。")
            return
        dialog = SvnRevisionPickerDialog(root, revision_input.text().strip() or "HEAD", self)
        if dialog.exec() == QDialog.Accepted:
            revision_input.setText(dialog.selected_revision)

    def _browse_output(self) -> None:
        output_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "保存批量对比报告",
            self.output_path_input.text().strip() or str(Path.cwd() / "batch_diff_report.html"),
            "HTML 报告 (*.html);;Markdown 报告 (*.md);;CSV 明细 (*.csv)",
        )
        if not output_path:
            return
        output_format = MainWindow._detect_export_mode(output_path, selected_filter)
        if output_format == "xml":
            output_format = "html"
        self.format_combo.setCurrentIndex(max(0, self.format_combo.findData(output_format)))
        self.output_path_input.setText(MainWindow._ensure_extension(output_path, output_format))


class BatchMergeDialog(QDialog):
    def __init__(
        self,
        left_root: str,
        right_root: str,
        output_dir: str,
        template_source: str,
        left_revision: str = "HEAD",
        right_revision: str = "HEAD",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("批量合并")
        self.resize(760, 260)

        layout = QVBoxLayout(self)
        form = QGridLayout()
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(8)

        self.left_root_input = QLineEdit(left_root)
        self.right_root_input = QLineEdit(right_root)
        self.left_revision_input = QLineEdit(left_revision or "HEAD")
        self.right_revision_input = QLineEdit(right_revision or "HEAD")
        self.output_dir_input = QLineEdit(output_dir)
        self.template_source_combo = QComboBox()
        self.template_source_combo.addItem("模板取左侧", "left")
        self.template_source_combo.addItem("模板取右侧", "right")
        self.template_source_combo.setCurrentIndex(max(0, self.template_source_combo.findData(template_source)))

        left_browse = QPushButton("浏览...")
        left_browse.clicked.connect(lambda: self._browse_root(self.left_root_input))
        right_browse = QPushButton("浏览...")
        right_browse.clicked.connect(lambda: self._browse_root(self.right_root_input))
        left_revision_pick = QPushButton("选择...")
        left_revision_pick.clicked.connect(lambda: self._pick_revision(self.left_root_input, self.left_revision_input))
        right_revision_pick = QPushButton("选择...")
        right_revision_pick.clicked.connect(lambda: self._pick_revision(self.right_root_input, self.right_revision_input))
        output_browse = QPushButton("输出到...")
        output_browse.clicked.connect(lambda: self._browse_root(self.output_dir_input))

        form.addWidget(QLabel("左侧根路径"), 0, 0)
        form.addWidget(self.left_root_input, 0, 1)
        form.addWidget(left_browse, 0, 2)
        form.addWidget(QLabel("右侧根路径"), 1, 0)
        form.addWidget(self.right_root_input, 1, 1)
        form.addWidget(right_browse, 1, 2)
        form.addWidget(QLabel("左侧SVN版本"), 2, 0)
        form.addWidget(self.left_revision_input, 2, 1)
        form.addWidget(left_revision_pick, 2, 2)
        form.addWidget(QLabel("右侧SVN版本"), 3, 0)
        form.addWidget(self.right_revision_input, 3, 1)
        form.addWidget(right_revision_pick, 3, 2)
        form.addWidget(QLabel("输出目录"), 4, 0)
        form.addWidget(self.output_dir_input, 4, 1)
        form.addWidget(output_browse, 4, 2)
        form.addWidget(QLabel("模板来源"), 5, 0)
        form.addWidget(self.template_source_combo, 5, 1)
        layout.addLayout(form)

        hint = QLabel("按同名文件批量合并；SVN 根路径会按填写的版本读取该版本下的完整目录内容。")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> dict[str, str]:
        return {
            "left_root": self.left_root_input.text().strip(),
            "right_root": self.right_root_input.text().strip(),
            "left_revision": self.left_revision_input.text().strip() or "HEAD",
            "right_revision": self.right_revision_input.text().strip() or "HEAD",
            "output_dir": self.output_dir_input.text().strip(),
            "template_source": str(self.template_source_combo.currentData() or "left"),
        }

    def _browse_root(self, target: QLineEdit) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择目录", target.text().strip() or str(Path.cwd()))
        if directory:
            target.setText(directory)

    def _pick_revision(self, root_input: QLineEdit, revision_input: QLineEdit) -> None:
        root = root_input.text().strip()
        if not root or infer_source_kind(root) != SOURCE_SVN:
            QMessageBox.information(self, "提示", "请先填写 SVN 根路径。")
            return
        dialog = SvnRevisionPickerDialog(root, revision_input.text().strip() or "HEAD", self)
        if dialog.exec() == QDialog.Accepted:
            revision_input.setText(dialog.selected_revision)


class WorkbookLoadWorker(QObject):
    stage = Signal(str, str)
    finished = Signal(str, object)
    failed = Signal(str, object)

    def __init__(self, side: str, source: WorkbookSource) -> None:
        super().__init__()
        self._side = side
        self._source = source

    def run(self) -> None:
        try:
            self.stage.emit(self._side, f"连接 {self._source.kind} 源…")
            workbook = load_workbook_from_source(self._source)
        except BaseException as exc:
            self.failed.emit(self._side, exc)
            return
        self.finished.emit(self._side, workbook)


class UpdateCheckWorker(QObject):
    finished = Signal(object)
    failed = Signal(object)

    def __init__(self, *, fetch_latest: bool = False) -> None:
        super().__init__()
        self._fetch_latest = fetch_latest

    def run(self) -> None:
        try:
            info = check_app_update("0.0.0" if self._fetch_latest else APP_VERSION)
        except BaseException as exc:
            self.failed.emit(exc)
            return
        self.finished.emit(info)


class UpdatePrepareWorker(QObject):
    progress = Signal(str, int, int)
    finished = Signal(object)
    failed = Signal(object)

    def __init__(self, info: UpdateInfo) -> None:
        super().__init__()
        self._info = info

    def run(self) -> None:
        try:
            prepared = prepare_update(
                self._info,
                progress_callback=lambda stage, current, total: self.progress.emit(stage, current, total),
            )
        except BaseException as exc:
            self.failed.emit(exc)
            return
        self.finished.emit(prepared)


class UpdateDialog(QDialog):
    install_now_requested = Signal(object)
    prepared_for_later = Signal(object)

    def __init__(self, info: UpdateInfo, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._info = info
        self._prepared: PreparedUpdate | None = None
        self._thread: QThread | None = None
        self._worker: UpdatePrepareWorker | None = None
        self._running = False

        self.setWindowTitle("软件更新")
        self.resize(520, 260)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        title = QLabel(f"发现新版本 v{info.latest_version}")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        self.summary_label = QLabel(
            f"当前版本：v{info.current_version}\n"
            f"更新来源：{info.release_root}\n\n"
            "下载和解压会在当前窗口内完成；安装时需要重启程序以覆盖正在运行的文件。"
        )
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.stage_label = QLabel("准备下载更新。")
        layout.addWidget(self.stage_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        if info.notes:
            notes = QTextEdit()
            notes.setReadOnly(True)
            notes.setMaximumHeight(72)
            notes.setPlainText(info.notes)
            layout.addWidget(notes)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        self.download_button = QPushButton("下载更新")
        self.install_button = QPushButton("立即重启并安装")
        self.later_button = QPushButton("稍后安装")
        self.cancel_button = QPushButton("取消")
        self.install_button.setEnabled(False)
        self.later_button.setEnabled(False)
        self.download_button.clicked.connect(self._start_prepare)
        self.install_button.clicked.connect(self._install_now)
        self.later_button.clicked.connect(self._install_later)
        self.cancel_button.clicked.connect(self.reject)
        button_row.addWidget(self.download_button)
        button_row.addWidget(self.install_button)
        button_row.addWidget(self.later_button)
        button_row.addWidget(self.cancel_button)
        layout.addLayout(button_row)

    def reject(self) -> None:
        if self._running:
            QMessageBox.information(self, "更新进行中", "正在下载或解压更新包，请等待完成。")
            return
        super().reject()

    def _start_prepare(self) -> None:
        if self._thread is not None:
            return
        self._running = True
        self.download_button.setEnabled(False)
        self.cancel_button.setEnabled(False)
        self.stage_label.setText("正在准备更新...")
        self.progress_bar.setRange(0, 0)

        thread = QThread(self)
        worker = UpdatePrepareWorker(self._info)
        worker.moveToThread(thread)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_prepare_finished)
        worker.failed.connect(self._on_prepare_failed)
        thread.started.connect(worker.run)
        self._thread = thread
        self._worker = worker
        thread.start()

    def _on_progress(self, stage: str, current: int, total: int) -> None:
        self.stage_label.setText(stage)
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(max(0, min(current, total)))
            return
        self.progress_bar.setRange(0, 0)

    def _on_prepare_finished(self, prepared: object) -> None:
        self._teardown_worker()
        if not isinstance(prepared, PreparedUpdate):
            self._on_prepare_failed(RuntimeError("更新准备结果无效。"))
            return
        self._prepared = prepared
        self._running = False
        self.stage_label.setText("更新包已下载并解压完成。")
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(1)
        self.install_button.setEnabled(True)
        self.later_button.setEnabled(True)
        self.cancel_button.setText("关闭")
        self.cancel_button.setEnabled(True)

    def _on_prepare_failed(self, exc: object) -> None:
        self._teardown_worker()
        self._running = False
        self.download_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.stage_label.setText("更新准备失败。")
        QMessageBox.warning(self, "更新失败", str(exc))

    def _teardown_worker(self) -> None:
        thread = self._thread
        worker = self._worker
        self._thread = None
        self._worker = None
        if thread is not None:
            thread.quit()
            thread.wait(2000)
            thread.deleteLater()
        if worker is not None:
            worker.deleteLater()

    def _install_now(self) -> None:
        if self._prepared is None:
            return
        self.install_now_requested.emit(self._prepared)
        self.accept()

    def _install_later(self) -> None:
        if self._prepared is None:
            return
        self.prepared_for_later.emit(self._prepared)
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"通用表格合并工具 v{APP_VERSION}")
        screen = QApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
            width = max(1100, min(int(geo.width() * 0.88), 1680))
            height = max(720, min(int(geo.height() * 0.86), 960))
            self.resize(width, height)
        else:
            self.resize(1400, 880)

        self.settings = load_settings()
        self.left_workbook: WorkbookData | None = None
        self.right_workbook: WorkbookData | None = None
        self.base_workbook: WorkbookData | None = None
        self.alignments: dict[str, SheetAlignment] = {}
        self.current_sheet_name: str | None = None
        self.left_folder_path: str | None = str(self.settings.get("left_root") or Path.cwd())
        self.right_folder_path: str | None = str(self.settings.get("right_root") or Path.cwd())
        self.base_folder_path: str | None = str(self.settings.get("base_root") or Path.cwd())
        self.base_revision: str = str(self.settings.get("base_revision", "HEAD") or "HEAD")
        self.pending_left_path: str | None = None
        self.pending_right_path: str | None = None
        self.pending_base_path: str | None = None
        self._sheet_state_warmup_queue: list[str] = []
        self._sheet_state_warmup_index = 0
        self._sheet_state_warmup_token = 0
        self._warmup_started_at = 0.0
        self._last_compare_metrics: dict[str, float] = {}
        self._resize_token = 0
        self._load_threads: dict[str, QThread] = {}
        self._load_workers: dict[str, WorkbookLoadWorker] = {}
        self._pending_workbooks: dict[str, WorkbookData | None] = {}
        self._pending_errors: dict[str, BaseException] = {}
        self._load_progress: QProgressDialog | None = None
        self._load_cancelled = False
        self._compare_started_at = 0.0
        self._update_thread: QThread | None = None
        self._update_worker: UpdateCheckWorker | None = None
        self._update_check_started = False
        self._manual_update_check_running = False
        self._pending_prepared_update: PreparedUpdate | None = None
        self._column_width_cache: dict[tuple[str, str, int, int], list[int]] = {}
        self._duplicate_id_rows: list[dict[str, object]] = []
        self._last_side_dock_width = int(self.settings.get("side_dock_width", 360) or 360)
        self._large_diff_confirmed_sheets: set[str] = set()

        self.sheet_list = QListWidget()
        self.sheet_list.setObjectName("sheetList")
        self.sheet_list.currentItemChanged.connect(self._on_sheet_changed)
        self.sheet_summary_label = QLabel("请选择左右两个 XML 文件后开始比对。")
        self.sheet_summary_label.setObjectName("sheetSummary")
        self.left_source_combo = QComboBox()
        self.right_source_combo = QComboBox()
        for combo in (self.left_source_combo, self.right_source_combo):
            combo.addItem("本地", SOURCE_LOCAL)
            combo.addItem("SVN", SOURCE_SVN)
            combo.addItem("Google Sheets", SOURCE_GOOGLE_SHEETS)
        self.left_file_combo = QComboBox()
        self.right_file_combo = QComboBox()
        self.left_revision_combo = QComboBox()
        self.right_revision_combo = QComboBox()
        for combo in (self.left_revision_combo, self.right_revision_combo):
            combo.setEditable(True)
            combo.setInsertPolicy(QComboBox.NoInsert)
        self.left_root_combo = QComboBox()
        self.right_root_combo = QComboBox()
        self.base_root_combo = QComboBox()
        for combo in (self.left_root_combo, self.right_root_combo, self.base_root_combo):
            combo.setEditable(True)
            combo.setInsertPolicy(QComboBox.NoInsert)
            combo.setMinimumContentsLength(40)
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            combo.setToolTip("输入本地目录或 SVN 根路径，再点击“选择文档”进入文档选择窗口")
        self.left_quick_combo = QComboBox()
        self.right_quick_combo = QComboBox()
        self.base_quick_combo = QComboBox()
        for combo in (self.left_quick_combo, self.right_quick_combo, self.base_quick_combo):
            combo.setMinimumContentsLength(16)
            combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.left_file_label = QLabel("左侧快照: 未选择")
        self.right_file_label = QLabel("右侧快照: 未选择")
        self.history_combo = QComboBox()
        self.history_combo.setMinimumContentsLength(28)
        self.rule_combo = QComboBox()
        for rule in list_merge_rules():
            self.rule_combo.addItem(rule.title, rule.rule_id)
        self.rule_combo.currentIndexChanged.connect(self._on_rule_changed)
        self.template_source_combo = QComboBox()
        self.template_source_combo.addItem("模板取左侧", "left")
        self.template_source_combo.addItem("模板取右侧", "right")
        self.rule_summary_label = QLabel()
        self.rule_summary_label.setWordWrap(True)
        self.three_way_checkbox = QCheckBox("启用三方合并")
        self.three_way_checkbox.setToolTip(
            "开启后需选择 Base（共同祖先）文件。\n"
            "单方修改将自动合并；左右均改时才标记为冲突。"
        )
        self.base_file_label = QLabel("Base: 未选择")
        self.compare_only_mode = QCheckBox("纯比对模式")
        self.compare_only_mode.setChecked(True)
        self.compare_only_mode.toggled.connect(self._update_compare_mode)
        self.ignore_trim_whitespace_diff = QCheckBox("忽略首尾空白差异")
        self.ignore_trim_whitespace_diff.setToolTip("开启后，比较时会忽略单元格内容前后的空格、换行和制表符差异。")
        self.ignore_all_whitespace_diff = QCheckBox("忽略所有空白")
        self.ignore_all_whitespace_diff.setToolTip("开启后比较时会把所有空格/换行/制表符全部去除。")
        self.ignore_case_diff = QCheckBox("忽略大小写")
        self.normalize_fullwidth_diff = QCheckBox("全/半角等价")
        self.normalize_fullwidth_diff.setToolTip("开启后把全角 ABC，（）等视为与半角 ABC,() 相同。")
        self.numeric_tolerance_input = QLineEdit()
        self.numeric_tolerance_input.setPlaceholderText("数值容差 (默认 0)")
        self.numeric_tolerance_input.setFixedWidth(120)
        self.numeric_tolerance_input.setToolTip("两侧都是数字时允许的绝对误差，例如 0.001。")
        self.key_fields_input = QLineEdit()
        self.key_fields_input.setPlaceholderText("手动关键字段，逗号分隔，例如 skillid,level")
        self.key_fields_input.setToolTip("留空时自动推断；填写后优先按这些逻辑字段对齐。")
        self.key_fields_button = QPushButton("选择字段")
        self.changed_sheets_only = QCheckBox("只看变更")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索当前表中的字段或内容")
        self._search_debounce_timer = QTimer(self)
        self._search_debounce_timer.setSingleShot(True)
        self._search_debounce_timer.setInterval(180)
        self._search_debounce_timer.timeout.connect(self._apply_filters)
        self.search_input.textChanged.connect(lambda _: self._search_debounce_timer.start())
        self.search_input.returnPressed.connect(lambda: self._jump_to_search(1, restart=False))
        self.view_mode_combo = QComboBox()
        self.view_mode_combo.addItem("看全部", "all")
        self.view_mode_combo.addItem("看差异", "changed")
        self.view_mode_combo.addItem("看冲突", "conflict")
        self.view_mode_combo.currentIndexChanged.connect(self._apply_filters)
        self.advanced_toggle = QPushButton("更多选项")
        self.advanced_toggle.setObjectName("ghostButton")
        self.advanced_toggle.setCheckable(True)
        self.detail_panel = QTextEdit()
        self.detail_panel.setObjectName("detailText")
        self.detail_panel.setReadOnly(True)
        self.detail_panel.setMinimumHeight(48)
        self.duplicate_id_list = QListWidget()
        self.duplicate_id_list.setObjectName("duplicateIdList")
        self.duplicate_id_list.setMaximumHeight(92)
        self.duplicate_id_list.itemDoubleClicked.connect(self._jump_to_duplicate_item)

        self.left_table = QTableView()
        self.middle_table = QTableView()
        self.right_table = QTableView()
        self.middle_panel: QWidget | None = None
        self.merge_action_bar: QWidget | None = None
        self.left_model: MergeTableModel | None = None
        self.middle_model: MergeTableModel | None = None
        self.right_model: MergeTableModel | None = None
        self._revision_entries: dict[str, list[RevisionEntry]] = {"left": [], "right": []}
        self._history_entries: list[dict] = []
        self._sheet_state_cache: dict[str, dict[str, object]] = {}
        self._suspend_auto_selection = False
        self._syncing_selection = False
        self._connected_selection_models: set[int] = set()
        self._last_search_text = ""
        self._last_search_match: tuple[int, int] | None = None

        self._build_ui()
        self._wire_tables()
        self._install_shortcuts()
        self._apply_window_style()
        self._autoload_defaults()

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("appRoot")
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        nav_rail = QFrame()
        nav_rail.setObjectName("navRail")
        nav_rail.setFixedWidth(64)
        nav_layout = QVBoxLayout(nav_rail)
        nav_layout.setContentsMargins(8, 12, 8, 10)
        nav_layout.setSpacing(8)
        app_badge = QLabel("合")
        app_badge.setObjectName("navBadge")
        app_badge.setAlignment(Qt.AlignCenter)
        nav_layout.addWidget(app_badge)
        nav_layout.addSpacing(8)

        self.source_panel_toggle = self._make_command_button(
            "来源",
            "配置",
            lambda: self._toggle_top_panel("source"),
            checkable=True,
            tooltip="展开或收起左右文档来源配置",
        )
        self.advanced_toggle.setObjectName("commandButton")
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.setText("合并设置")
        self.advanced_toggle.setMinimumSize(72, 36)
        self.advanced_toggle.setMaximumHeight(36)
        self.advanced_toggle.setToolTip("展开或收起合并规则、文本比较和显示设置")
        self.sheet_panel_toggle = self._make_nav_button("子表", lambda: self._toggle_side_panel("sheet"), checkable=True)
        nav_layout.addWidget(self.sheet_panel_toggle)
        nav_layout.addStretch(1)
        self.update_check_button = QPushButton(f"v{APP_VERSION}\n更新")
        self.update_check_button.setObjectName("updateCheckButton")
        self.update_check_button.setToolTip(f"当前版本 v{APP_VERSION}\n点击检查最新版本")
        self.update_check_button.clicked.connect(self._manual_check_update)
        nav_layout.addWidget(self.update_check_button)

        workspace = QWidget()
        workspace.setObjectName("workspaceRoot")
        workspace_layout = QVBoxLayout(workspace)
        workspace_layout.setContentsMargins(10, 10, 10, 10)
        workspace_layout.setSpacing(8)

        command_bar = QFrame()
        command_bar.setObjectName("commandBar")
        command_layout = QHBoxLayout(command_bar)
        command_layout.setContentsMargins(14, 9, 14, 9)
        command_layout.setSpacing(7)
        title_box = QFrame()
        title_box.setObjectName("appTitleBox")
        title_layout = QVBoxLayout(title_box)
        title_layout.setContentsMargins(0, 0, 18, 0)
        title_layout.setSpacing(0)
        title = QLabel("分久必合")
        title.setObjectName("appTitle")
        subtitle = QLabel("文档对比 / 合并工作台")
        subtitle.setObjectName("appSubtitle")
        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        command_layout.addWidget(title_box)
        command_layout.addSpacing(6)

        self.prev_diff_button = self._make_command_button(
            "上一",
            "差异",
            lambda: self._jump_to_diff(-1),
            tooltip="在当前视图内跳到上一处差异/冲突 (Shift+F3)",
        )
        self.next_diff_button = self._make_command_button(
            "下一",
            "差异",
            lambda: self._jump_to_diff(1),
            tooltip="在当前视图内跳到下一处差异/冲突 (F3)",
        )
        for button in (
            self._make_command_button("开始", "比对", self._start_compare, primary=True),
            self._make_command_button("导出", "结果", self._export),
            self._make_command_button("差异", "报告", self._export_diff_report),
        ):
            command_layout.addWidget(button)
        self._add_command_separator(command_layout)
        for button in (
            self._make_command_button("批量", "对比", self._start_batch_compare),
            self._make_command_button("批量", "合并", self._start_batch_merge),
        ):
            command_layout.addWidget(button)
        self._add_command_separator(command_layout)
        command_layout.addWidget(self.source_panel_toggle)
        command_layout.addWidget(self.advanced_toggle)
        self._add_command_separator(command_layout)
        for button in (
            self.prev_diff_button,
            self.next_diff_button,
        ):
            command_layout.addWidget(button)
        command_layout.addStretch(1)
        command_layout.addWidget(QLabel("搜索"))
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(280)
        command_layout.addWidget(self.search_input, 1)
        workspace_layout.addWidget(command_bar)

        self.source_summary_bar = QFrame()
        self.source_summary_bar.setObjectName("sourceSummaryBar")
        source_summary_layout = QHBoxLayout(self.source_summary_bar)
        source_summary_layout.setContentsMargins(12, 5, 12, 5)
        source_summary_layout.setSpacing(8)
        source_summary_layout.addWidget(QLabel("当前来源"))
        self.source_summary_left_label = QLabel("左: 未选择")
        self.source_summary_left_label.setObjectName("sourceSummaryText")
        self.source_summary_right_label = QLabel("右: 未选择")
        self.source_summary_right_label.setObjectName("sourceSummaryText")
        self.source_summary_rule_label = QLabel("规则: 未选择")
        self.source_summary_rule_label.setObjectName("sourceSummaryText")
        source_summary_layout.addWidget(self.source_summary_left_label, 2)
        source_summary_layout.addWidget(self.source_summary_right_label, 2)
        source_summary_layout.addWidget(self.source_summary_rule_label, 1)
        source_summary_layout.addStretch(1)
        choose_left_summary_button = QPushButton("换左")
        choose_left_summary_button.setObjectName("miniButton")
        choose_left_summary_button.clicked.connect(lambda: self._choose_file("left"))
        source_summary_layout.addWidget(choose_left_summary_button)
        choose_right_summary_button = QPushButton("换右")
        choose_right_summary_button.setObjectName("miniButton")
        choose_right_summary_button.clicked.connect(lambda: self._choose_file("right"))
        source_summary_layout.addWidget(choose_right_summary_button)
        change_source_button = QPushButton("来源配置")
        change_source_button.setObjectName("miniButton")
        change_source_button.clicked.connect(lambda: self._toggle_top_panel("source"))
        source_summary_layout.addWidget(change_source_button)
        self.source_summary_bar.setVisible(False)
        workspace_layout.addWidget(self.source_summary_bar)

        self.top_option_panel = QFrame()
        self.top_option_panel.setObjectName("topOptionPanel")
        self.top_option_layout = QVBoxLayout(self.top_option_panel)
        self.top_option_layout.setContentsMargins(10, 8, 10, 10)
        self.top_option_layout.setSpacing(8)
        self.top_option_panel.setVisible(False)
        workspace_layout.addWidget(self.top_option_panel)

        filter_strip = QFrame()
        filter_strip.setObjectName("filterStrip")
        self.filter_strip = filter_strip
        filter_layout = QHBoxLayout(filter_strip)
        filter_layout.setContentsMargins(12, 7, 12, 7)
        filter_layout.setSpacing(8)
        filter_layout.addWidget(self.compare_only_mode)
        filter_layout.addWidget(self.three_way_checkbox)
        filter_layout.addSpacing(8)
        filter_layout.addWidget(QLabel("查看"))
        filter_layout.addWidget(self.view_mode_combo)
        filter_layout.addWidget(self.ignore_trim_whitespace_diff)
        filter_layout.addSpacing(10)
        legend_row = self._build_legend_row()
        filter_layout.addLayout(legend_row)
        filter_layout.addStretch(1)
        workspace_layout.addWidget(filter_strip)

        body_splitter = QSplitter(Qt.Horizontal)
        body_splitter.setChildrenCollapsible(False)
        self.content_splitter = body_splitter
        body_splitter.splitterMoved.connect(self._on_side_dock_splitter_moved)

        sheet_shell = QWidget()
        sheet_shell.setObjectName("sheetShell")
        sheet_shell_layout = QVBoxLayout(sheet_shell)
        sheet_shell_layout.setContentsMargins(0, 0, 0, 0)
        sheet_shell_layout.setSpacing(8)
        self.sheet_shell = sheet_shell

        self.side_drawer_header = QFrame()
        self.side_drawer_header.setObjectName("sideDrawerHeader")
        side_header_layout = QHBoxLayout(self.side_drawer_header)
        side_header_layout.setContentsMargins(12, 10, 10, 8)
        side_header_layout.setSpacing(8)
        side_title_box = QVBoxLayout()
        side_title_box.setContentsMargins(0, 0, 0, 0)
        side_title_box.setSpacing(0)
        self.side_drawer_title = QLabel("子表")
        self.side_drawer_title.setObjectName("sideDrawerTitle")
        self.side_drawer_subtitle = QLabel("工作表导航")
        self.side_drawer_subtitle.setObjectName("sideDrawerSubtitle")
        side_title_box.addWidget(self.side_drawer_title)
        side_title_box.addWidget(self.side_drawer_subtitle)
        side_header_layout.addLayout(side_title_box, 1)
        self.side_drawer_close_button = QPushButton("×")
        self.side_drawer_close_button.setObjectName("drawerCloseButton")
        self.side_drawer_close_button.setFixedSize(28, 28)
        self.side_drawer_close_button.setToolTip("收起左侧面板")
        self.side_drawer_close_button.clicked.connect(lambda: self._activate_side_panel(""))
        side_header_layout.addWidget(self.side_drawer_close_button)
        sheet_shell_layout.addWidget(self.side_drawer_header)

        self.source_panel = QFrame()
        self.source_panel.setObjectName("sourcePanel")
        source_panel_layout = QVBoxLayout(self.source_panel)
        source_panel_layout.setContentsMargins(10, 0, 10, 10)
        source_panel_layout.setSpacing(8)

        source_grid = QGridLayout()
        source_grid.setHorizontalSpacing(6)
        source_grid.setVerticalSpacing(5)

        left_source_card = QFrame()
        left_source_card.setObjectName("compactSourceCard")
        left_source_layout = QGridLayout(left_source_card)
        left_source_layout.setContentsMargins(8, 8, 8, 8)
        left_source_layout.setHorizontalSpacing(6)
        left_source_layout.setVerticalSpacing(5)
        left_source_layout.addWidget(QLabel("左侧"), 0, 0)
        left_source_layout.addWidget(self.left_quick_combo, 0, 1)
        left_source_layout.addWidget(self.left_root_combo, 1, 0, 1, 2)
        left_pick_button = QPushButton("选择文档")
        left_pick_button.setToolTip("打开文档选择窗口；窗口内可继续切换根路径或浏览仓库。")
        left_pick_button.clicked.connect(lambda: self._choose_file("left"))
        left_source_layout.addWidget(left_pick_button, 2, 0)
        left_source_layout.addWidget(self.left_file_label, 2, 1)

        right_source_card = QFrame()
        right_source_card.setObjectName("compactSourceCard")
        right_source_layout = QGridLayout(right_source_card)
        right_source_layout.setContentsMargins(8, 8, 8, 8)
        right_source_layout.setHorizontalSpacing(6)
        right_source_layout.setVerticalSpacing(5)
        right_source_layout.addWidget(QLabel("右侧"), 0, 0)
        right_source_layout.addWidget(self.right_quick_combo, 0, 1)
        right_source_layout.addWidget(self.right_root_combo, 1, 0, 1, 2)
        right_pick_button = QPushButton("选择文档")
        right_pick_button.setToolTip("打开文档选择窗口；窗口内可继续切换根路径或浏览仓库。")
        right_pick_button.clicked.connect(lambda: self._choose_file("right"))
        right_source_layout.addWidget(right_pick_button, 2, 0)
        right_source_layout.addWidget(self.right_file_label, 2, 1)

        source_grid.addWidget(left_source_card, 0, 0)
        source_grid.addWidget(right_source_card, 0, 1)
        source_grid.setColumnStretch(0, 1)
        source_grid.setColumnStretch(1, 1)
        source_panel_layout.addLayout(source_grid)

        same_name_button = QPushButton("右侧同名")
        same_name_button.clicked.connect(self._match_right_file_by_left_name)
        manage_quick_button = QPushButton("管理快捷配置")
        manage_quick_button.clicked.connect(self._manage_quick_roots)
        refresh_cache_button = QPushButton("刷新缓存")
        refresh_cache_button.setToolTip("清空工作簿 / SVN 版本 / Google 元数据缓存，强制下次对比重新拉取。")
        refresh_cache_button.clicked.connect(self._refresh_sources_caches)
        quick_actions = QHBoxLayout()
        quick_actions.setSpacing(6)
        quick_actions.addWidget(manage_quick_button)
        quick_actions.addWidget(same_name_button)
        quick_actions.addWidget(refresh_cache_button)
        quick_actions.addStretch(1)
        source_panel_layout.addLayout(quick_actions)

        # Base (three-way) row is intentionally separate so the pick button never stretches across the path area.
        self._base_row_container = QWidget()
        base_row = QGridLayout(self._base_row_container)
        base_row.setContentsMargins(0, 4, 0, 0)
        base_row.setHorizontalSpacing(6)
        base_row.setVerticalSpacing(6)
        self._base_row_label = QLabel("Base（三方）")
        self._base_pick_button = QPushButton("选择文档")
        self._base_pick_button.clicked.connect(self._choose_base_file)
        self.base_file_label.setToolTip("Base 是三方合并的共同祖先版本。")
        base_row.addWidget(self._base_row_label, 0, 0)
        base_row.addWidget(self.base_quick_combo, 0, 1)
        base_row.addWidget(self.base_root_combo, 1, 0, 1, 2)
        base_row.addWidget(self._base_pick_button, 2, 0)
        base_row.addWidget(self.base_file_label, 2, 1)
        source_panel_layout.addWidget(self._base_row_container)

        # Hide base row by default; shown when three_way_checkbox is toggled.
        self._base_row_label.setVisible(False)
        self.base_quick_combo.setVisible(False)
        self.base_root_combo.setVisible(False)
        self._base_pick_button.setVisible(False)
        self.base_file_label.setVisible(False)
        self._base_row_container.setVisible(False)

        self.source_panel.setVisible(False)
        self.top_option_layout.addWidget(self.source_panel)

        self.advanced_panel = QWidget()
        self.advanced_panel.setObjectName("optionTray")
        advanced_grid = QGridLayout(self.advanced_panel)
        advanced_grid.setContentsMargins(8, 0, 8, 8)
        advanced_grid.setHorizontalSpacing(8)
        advanced_grid.setVerticalSpacing(6)
        advanced_grid.addWidget(QLabel("默认规则"), 0, 0)
        advanced_grid.addWidget(self.rule_combo, 0, 1)
        advanced_grid.addWidget(QLabel("模板来源"), 0, 2)
        advanced_grid.addWidget(self.template_source_combo, 0, 3)
        advanced_grid.addWidget(QLabel("关键字段"), 0, 4)
        advanced_grid.addWidget(self.key_fields_input, 0, 5)
        advanced_grid.addWidget(self.key_fields_button, 0, 6)

        self.font_size_combo = QComboBox()
        for size in (8, 9, 10, 11, 12, 13, 14):
            self.font_size_combo.addItem(f"{size} pt", size)
        self.row_height_combo = QComboBox()
        for height in (20, 22, 24, 28, 32, 36, 40, 48, 56):
            self.row_height_combo.addItem(f"行高 {height}", height)
        self.header_height_combo = QComboBox()
        for height in (24, 28, 32, 40, 56, 72, 96, 128):
            self.header_height_combo.addItem(f"表头 {height}", height)
        advanced_grid.addWidget(self.ignore_all_whitespace_diff, 1, 0, 1, 2)
        advanced_grid.addWidget(self.ignore_case_diff, 1, 2)
        advanced_grid.addWidget(self.normalize_fullwidth_diff, 1, 3)
        advanced_grid.addWidget(QLabel("容差"), 1, 4)
        advanced_grid.addWidget(self.numeric_tolerance_input, 1, 5)
        advanced_grid.addWidget(self.font_size_combo, 1, 6)
        advanced_grid.addWidget(self.row_height_combo, 1, 7)
        advanced_grid.addWidget(self.header_height_combo, 1, 8)
        advanced_grid.setColumnStretch(1, 2)
        advanced_grid.setColumnStretch(3, 1)
        advanced_grid.setColumnStretch(5, 2)
        self.advanced_panel.setVisible(False)
        self.top_option_layout.addWidget(self.advanced_panel)

        left_panel = self._make_card()
        left_panel.setObjectName("sidePanel")
        self.sheet_panel = left_panel
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 0, 8, 8)
        left_layout.setSpacing(6)
        sheet_header_row = QHBoxLayout()
        sheet_header_row.setContentsMargins(2, 0, 2, 0)
        sheet_header_row.addWidget(QLabel("工作表"))
        sheet_header_row.addStretch(1)
        sheet_header_row.addWidget(self.changed_sheets_only)
        left_layout.addLayout(sheet_header_row)
        left_layout.addWidget(self.sheet_summary_label)
        self.sheet_filter_input = QLineEdit()
        self.sheet_filter_input.setPlaceholderText("搜索工作表名")
        self.sheet_filter_input.setClearButtonEnabled(True)
        self.sheet_filter_input.textChanged.connect(self._filter_sheet_list)
        left_layout.addWidget(self.sheet_filter_input)
        left_layout.addWidget(self.sheet_list, 1)
        sheet_shell_layout.addWidget(left_panel, 1)
        body_splitter.addWidget(sheet_shell)

        main_panel = self._make_card()
        main_panel.setObjectName("workspacePanel")
        main_layout = QVBoxLayout(main_panel)
        main_layout.setContentsMargins(7, 7, 7, 7)
        main_layout.setSpacing(6)

        self.prepare_panel = self._build_prepare_panel()
        main_layout.addWidget(self.prepare_panel, 1)

        self.merge_action_bar = QWidget()
        self.merge_action_bar.setObjectName("mergeActionBar")
        action_layout = QHBoxLayout(self.merge_action_bar)
        action_layout.setContentsMargins(8, 4, 8, 4)
        action_layout.setSpacing(6)
        action_layout.addWidget(QLabel("人工处理"))
        take_left_button = QPushButton("取左")
        take_left_button.clicked.connect(lambda: self._apply_row_choice("left"))
        action_layout.addWidget(take_left_button)
        take_right_button = QPushButton("取右")
        take_right_button.clicked.connect(lambda: self._apply_row_choice("right"))
        action_layout.addWidget(take_right_button)
        reset_button = QPushButton("重置自动")
        reset_button.clicked.connect(lambda: self._apply_row_choice("auto"))
        action_layout.addWidget(reset_button)
        check_merged_id_button = QPushButton("检测合并ID")
        check_merged_id_button.setToolTip("按当前子表右键指定的ID列检查预合并结果是否有重复ID。")
        check_merged_id_button.clicked.connect(self._check_merged_ids)
        action_layout.addWidget(check_merged_id_button)
        action_layout.addSpacing(16)
        action_layout.addWidget(QLabel("备注"))
        self.row_note_input = QLineEdit()
        self.row_note_input.setPlaceholderText("给当前行加一个备注，便于导出/沟通")
        self.row_note_input.editingFinished.connect(self._apply_row_note)
        action_layout.addWidget(self.row_note_input, 1)
        action_layout.addStretch(0)
        main_layout.addWidget(self.merge_action_bar)

        self.sheet_overview_bar = QWidget()
        self.sheet_overview_bar.setObjectName("sheetOverviewBar")
        overview_layout = QHBoxLayout(self.sheet_overview_bar)
        overview_layout.setContentsMargins(8, 3, 8, 3)
        overview_layout.setSpacing(6)
        overview_title = QLabel("当前表")
        overview_title.setObjectName("overviewTitle")
        overview_layout.addWidget(overview_title)
        self.overview_added_button = QPushButton("新增 0")
        self.overview_deleted_button = QPushButton("删除 0")
        self.overview_conflict_button = QPushButton("冲突 0")
        self.overview_changed_columns_button = QPushButton("改动列 0")
        self.overview_duplicate_id_button = QPushButton("重复ID 0")
        self.overview_added_button.setObjectName("pillAdded")
        self.overview_deleted_button.setObjectName("pillDeleted")
        self.overview_conflict_button.setObjectName("pillConflict")
        self.overview_changed_columns_button.setObjectName("pillChanged")
        self.overview_duplicate_id_button.setObjectName("pillDuplicate")
        self.overview_added_button.clicked.connect(
            lambda: self._jump_to_row_status({"right_only"}, "新增")
        )
        self.overview_deleted_button.clicked.connect(
            lambda: self._jump_to_row_status({"left_only", "deleted"}, "删除")
        )
        self.overview_conflict_button.clicked.connect(
            lambda: self._jump_to_row_status({"conflict"}, "冲突")
        )
        self.overview_changed_columns_button.clicked.connect(self._jump_to_changed_row)
        self.overview_duplicate_id_button.clicked.connect(self._jump_to_next_duplicate_id)
        for button in (
            self.overview_added_button,
            self.overview_deleted_button,
            self.overview_conflict_button,
            self.overview_changed_columns_button,
            self.overview_duplicate_id_button,
        ):
            button.setMaximumHeight(26)
            overview_layout.addWidget(button)
        overview_layout.addStretch(1)
        main_layout.addWidget(self.sheet_overview_bar)

        splitter = QSplitter()
        splitter.addWidget(self._wrap_table("左侧文件", self.left_table))
        self.middle_panel = self._wrap_table("预合并结果", self.middle_table)
        splitter.addWidget(self.middle_panel)
        splitter.addWidget(self._wrap_table("右侧文件", self.right_table))
        splitter.setSizes([500, 620, 500])
        self.tables_splitter = splitter
        detail_container = self._make_card()
        detail_container.setObjectName("detailPanel")
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(7, 5, 7, 7)
        detail_layout.setSpacing(4)
        detail_header = QFrame()
        detail_header.setObjectName("detailHeaderBar")
        detail_header_layout = QHBoxLayout(detail_header)
        detail_header_layout.setContentsMargins(8, 2, 8, 2)
        detail_header_layout.setSpacing(6)
        detail_title = QLabel("当前行差异详情")
        detail_title.setObjectName("detailTitle")
        detail_header_layout.addWidget(detail_title)
        detail_header_layout.addStretch(1)
        detail_tip = QLabel("选择一行查看字段级差异")
        detail_tip.setObjectName("detailTip")
        detail_header_layout.addWidget(detail_tip)
        detail_layout.addWidget(detail_header)
        detail_layout.addWidget(self.detail_panel)
        self.duplicate_id_label = QLabel("重复 ID 定位")
        self.duplicate_id_label.setObjectName("detailTitle")
        detail_layout.addWidget(self.duplicate_id_label)
        detail_layout.addWidget(self.duplicate_id_list)
        self.duplicate_id_label.setVisible(False)
        self.duplicate_id_list.setVisible(False)
        vertical_splitter = QSplitter(Qt.Vertical)
        vertical_splitter.addWidget(splitter)
        vertical_splitter.addWidget(detail_container)
        vertical_splitter.setCollapsible(0, False)
        vertical_splitter.setCollapsible(1, True)
        vertical_splitter.setSizes([760, 104])
        self.table_area_splitter = vertical_splitter
        main_layout.addWidget(vertical_splitter, 1)

        body_splitter.addWidget(main_panel)
        body_splitter.setStretchFactor(0, 0)
        body_splitter.setStretchFactor(1, 1)
        initial_side_width = self._clamp_side_dock_width(self._last_side_dock_width)
        body_splitter.setSizes([initial_side_width, 1280])
        active_side_panel = str(self.settings.get("active_side_panel", "sheet") or "sheet")
        active_top_panel = str(self.settings.get("active_top_panel", "") or "")
        if active_side_panel in {"source", "settings"} and not active_top_panel:
            active_top_panel = active_side_panel
            active_side_panel = ""
        self._activate_side_panel(active_side_panel if active_side_panel == "sheet" else "", save=False)
        self._activate_top_panel(active_top_panel, save=False)
        workspace_layout.addWidget(body_splitter, 1)
        root_layout.addWidget(nav_rail)
        root_layout.addWidget(workspace, 1)
        self.setCentralWidget(central)
        self.setStatusBar(QStatusBar())
        self._status_permanent_label = QLabel("")
        self._status_permanent_label.setStyleSheet("color: #667085; padding: 0 8px;")
        self.statusBar().addPermanentWidget(self._status_permanent_label)
        self._pending_update_button = QPushButton("重启安装更新")
        self._pending_update_button.setObjectName("pendingUpdateButton")
        self._pending_update_button.setMaximumHeight(24)
        self._pending_update_button.setVisible(False)
        self._pending_update_button.clicked.connect(self._install_pending_update)
        self.statusBar().addPermanentWidget(self._pending_update_button)

    def _build_prepare_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("preparePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        hero = QHBoxLayout()
        hero.setSpacing(14)
        title_box = QVBoxLayout()
        title_box.setSpacing(4)
        title = QLabel("准备开始比对")
        title.setObjectName("prepareTitle")
        subtitle = QLabel("选择左右文档后直接开始；复杂路径、SVN、Google 配置可从“来源”展开。")
        subtitle.setObjectName("prepareSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        hero.addLayout(title_box, 1)
        open_source_button = QPushButton("打开来源配置")
        open_source_button.setObjectName("prepareSecondaryButton")
        open_source_button.clicked.connect(lambda: self._toggle_top_panel("source"))
        hero.addWidget(open_source_button)
        layout.addLayout(hero)

        cards = QHBoxLayout()
        cards.setSpacing(14)
        cards.addWidget(self._build_prepare_source_card("left", "左侧文档"), 1)
        cards.addWidget(self._build_prepare_source_card("right", "右侧文档"), 1)
        layout.addLayout(cards)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.prepare_match_button = QPushButton("右侧同名")
        self.prepare_match_button.setObjectName("prepareSecondaryButton")
        self.prepare_match_button.clicked.connect(self._match_right_file_by_left_name)
        action_row.addWidget(self.prepare_match_button)
        action_row.addStretch(1)
        self.prepare_hint_label = QLabel("未选择左右文档")
        self.prepare_hint_label.setObjectName("prepareHint")
        action_row.addWidget(self.prepare_hint_label)
        self.prepare_start_button = QPushButton("开始比对")
        self.prepare_start_button.setObjectName("preparePrimaryButton")
        self.prepare_start_button.clicked.connect(self._start_compare)
        action_row.addWidget(self.prepare_start_button)
        layout.addLayout(action_row)
        layout.addStretch(1)
        return panel

    def _build_prepare_source_card(self, side: str, title: str) -> QFrame:
        card = QFrame()
        card.setObjectName("prepareSourceCard")
        card.setMaximumHeight(260)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setObjectName("prepareCardTitle")
        header.addWidget(title_label)
        header.addStretch(1)
        type_label = QLabel("未配置")
        type_label.setObjectName("prepareBadge")
        header.addWidget(type_label)
        layout.addLayout(header)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(8)
        quick_label = QLabel("快捷")
        quick_label.setObjectName("preparePath")
        quick_combo = QComboBox()
        quick_combo.setObjectName("prepareQuickCombo")
        quick_combo.setMinimumContentsLength(18)
        quick_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        quick_combo.currentIndexChanged.connect(lambda index, s=side: self._apply_prepare_quick_root(s, index))
        quick_row.addWidget(quick_label)
        quick_row.addWidget(quick_combo, 1)
        layout.addLayout(quick_row)

        root_row = QHBoxLayout()
        root_row.setSpacing(8)
        root_label = QLabel("路径")
        root_label.setObjectName("preparePath")
        root_combo = QComboBox()
        root_combo.setObjectName("prepareRootCombo")
        root_combo.setEditable(True)
        root_combo.setInsertPolicy(QComboBox.NoInsert)
        root_combo.setMinimumContentsLength(32)
        root_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        root_combo.activated.connect(lambda _index, s=side: self._apply_prepare_root_text(s))
        if root_combo.lineEdit() is not None:
            root_combo.lineEdit().editingFinished.connect(lambda s=side: self._apply_prepare_root_text(s))
        root_row.addWidget(root_label)
        root_row.addWidget(root_combo, 1)
        file_label = QLabel("快照: 未选择")
        file_label.setObjectName("prepareFile")
        file_label.setWordWrap(True)
        file_label.setMinimumHeight(64)
        file_label.setMaximumHeight(88)
        layout.addLayout(root_row)
        layout.addWidget(file_label)

        choose_button = QPushButton(f"选择{title}")
        choose_button.setObjectName("prepareChooseButton")
        choose_button.clicked.connect(lambda: self._choose_file(side))
        layout.addWidget(choose_button)

        if side == "left":
            self.prepare_left_quick_combo = quick_combo
            self.prepare_left_type_label = type_label
            self.prepare_left_root_combo = root_combo
            self.prepare_left_file_label = file_label
        else:
            self.prepare_right_quick_combo = quick_combo
            self.prepare_right_type_label = type_label
            self.prepare_right_root_combo = root_combo
            self.prepare_right_file_label = file_label
        return card

    def _wrap_table(self, title: str, table: QTableView) -> QWidget:
        widget = QFrame()
        widget.setObjectName("tablePanel")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(8, 7, 8, 8)
        layout.setSpacing(4)
        header = QFrame()
        header.setObjectName("tableHeaderBar")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(8, 2, 8, 2)
        header_layout.setSpacing(6)
        title_label = QLabel(title)
        title_label.setObjectName("tableTitle")
        header_layout.addWidget(title_label)
        header_layout.addStretch(1)
        layout.addWidget(header)
        table.verticalHeader().setDefaultSectionSize(int(self.settings.get("table_row_height", 24) or 24))
        table.horizontalHeader().setDefaultSectionSize(int(self.settings.get("table_header_height", 32) or 32))
        table.horizontalHeader().setMinimumHeight(int(self.settings.get("table_header_height", 32) or 32))
        table.horizontalHeader().setMaximumHeight(int(self.settings.get("table_header_height", 32) or 32))
        table.setSelectionBehavior(QTableView.SelectRows)
        table.setSelectionMode(QTableView.ExtendedSelection)
        table.setWordWrap(False)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)
        layout.addWidget(table)
        return widget

    def _make_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        return card

    def _make_toolbar_group(self, title: str) -> tuple[QFrame, QVBoxLayout]:
        group = QFrame()
        group.setObjectName("toolbarGroup")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 7, 10, 8)
        layout.setSpacing(5)
        label = QLabel(title)
        label.setObjectName("toolbarGroupTitle")
        layout.addWidget(label)
        return group, layout

    def _add_command_separator(self, layout: QHBoxLayout) -> None:
        separator = QFrame()
        separator.setObjectName("commandSeparator")
        separator.setFixedWidth(1)
        separator.setMinimumHeight(28)
        layout.addWidget(separator)

    def _make_nav_button(self, text: str, handler=None, *, checkable: bool = False) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("navButton")
        button.setCheckable(checkable)
        button.setMinimumHeight(40)
        if handler is not None:
            button.clicked.connect(handler)
        return button

    def _make_command_button(
        self,
        title: str,
        subtitle: str,
        handler,
        *,
        primary: bool = False,
        checkable: bool = False,
        tooltip: str = "",
    ) -> QPushButton:
        button = QPushButton(f"{title}{subtitle}")
        button.setObjectName("commandPrimaryButton" if primary else "commandButton")
        button.setMinimumSize(72, 36)
        button.setMaximumHeight(36)
        button.setCheckable(checkable)
        if tooltip:
            button.setToolTip(tooltip)
        button.clicked.connect(handler)
        return button

    def _apply_window_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: #f5f7fb;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                color: #111827;
            }
            QWidget#appRoot {
                background: #f5f7fb;
            }
            QWidget#workspaceRoot {
                background: #f5f7fb;
            }
            QFrame#navRail {
                background: #fbfcfe;
                border: none;
                border-right: 1px solid #e5eaf2;
            }
            QLabel#navBadge {
                min-width: 40px;
                max-width: 40px;
                min-height: 40px;
                max-height: 40px;
                border-radius: 12px;
                background: #2563eb;
                color: #ffffff;
                font-size: 20px;
                font-weight: 900;
            }
            QPushButton#navButton {
                min-height: 40px;
                padding: 0;
                border-radius: 12px;
                border: 1px solid transparent;
                background: transparent;
                color: #64748b;
                font-size: 12px;
                font-weight: 800;
            }
            QPushButton#navButton:hover {
                background: #eef4ff;
                border-color: #dbeafe;
                color: #1d4ed8;
            }
            QPushButton#navButton:checked {
                background: #2563eb;
                border-color: #2563eb;
                color: #ffffff;
            }
            QPushButton#updateCheckButton {
                min-height: 38px;
                padding: 4px 2px;
                border-radius: 11px;
                border: 1px solid #e2e8f0;
                background: #f8fafc;
                color: #64748b;
                font-size: 10px;
                font-weight: 700;
            }
            QPushButton#updateCheckButton:hover {
                background: #eff6ff;
                color: #1d4ed8;
                border-color: #bfdbfe;
            }
            QPushButton#updateCheckButton:disabled {
                color: #94a3b8;
                background: #f1f5f9;
                border-color: #e2e8f0;
            }
            QFrame#card, QFrame#topBar, QFrame#workspacePanel, QFrame#detailPanel, QFrame#tablePanel {
                background: #ffffff;
                border: 1px solid #e1e8f2;
                border-radius: 16px;
            }
            QFrame#topBar {
                background: #ffffff;
                border: 1px solid #dce5f0;
            }
            QFrame#workspacePanel {
                background: #ffffff;
                border: 1px solid #dfe7f1;
            }
            QFrame#tablePanel {
                background: #ffffff;
                border: 1px solid #e4ecf5;
                border-radius: 14px;
            }
            QFrame#quickBox {
                background: #f7f7fa;
                border: 1px solid #e3e3e8;
                border-radius: 10px;
            }
            QFrame#toolbarGroup {
                background: #f8fafc;
                border: 1px solid #e4eaf2;
                border-radius: 14px;
            }
            QFrame#commandBar {
                background: #ffffff;
                border: 1px solid #dfe7f1;
                border-radius: 16px;
            }
            QFrame#filterStrip {
                background: #ffffff;
                border: 1px solid #e4ecf5;
                border-radius: 14px;
            }
            QFrame#sourceSummaryBar {
                background: #ffffff;
                border: 1px solid #e6edf5;
                border-radius: 14px;
            }
            QFrame#topOptionPanel {
                background: #ffffff;
                border: 1px solid #dfe7f1;
                border-radius: 16px;
            }
            QLabel#sourceSummaryText {
                color: #334155;
                font-size: 12px;
                font-weight: 700;
                padding: 2px 8px;
                background: #f8fafc;
                border: 1px solid #edf2f7;
                border-radius: 9px;
            }
            QPushButton#miniButton {
                min-height: 26px;
                max-height: 26px;
                padding: 0 12px;
                border-radius: 9px;
                border: 1px solid #dbe5f0;
                background: #ffffff;
                color: #334155;
                font-size: 12px;
                font-weight: 800;
            }
            QPushButton#miniButton:hover {
                background: #eff6ff;
                color: #1d4ed8;
                border-color: #bfdbfe;
            }
            QFrame#appTitleBox {
                background: transparent;
                border: none;
            }
            QLabel#appTitle {
                color: #111827;
                font-size: 20px;
                font-weight: 800;
                letter-spacing: 1px;
            }
            QLabel#dialogTitle {
                color: #111827;
                font-size: 18px;
                font-weight: 800;
            }
            QLabel#appSubtitle {
                color: #64748b;
                font-size: 11px;
                font-weight: 500;
            }
            QFrame#commandSeparator {
                background: #d8e0ea;
                border: none;
                margin-left: 4px;
                margin-right: 4px;
            }
            QLabel#toolbarGroupTitle {
                color: #64748b;
                font-size: 11px;
                font-weight: 700;
                padding-left: 2px;
            }
            QLabel#panelTitle {
                color: #111827;
                font-size: 13px;
                font-weight: 800;
                padding: 0 0 4px 2px;
            }
            QFrame#sideDrawerHeader {
                background: transparent;
                border: none;
                border-bottom: 1px solid #e8eef6;
                border-radius: 0;
            }
            QLabel#sideDrawerTitle {
                color: #111827;
                font-size: 15px;
                font-weight: 900;
            }
            QLabel#sideDrawerSubtitle {
                color: #64748b;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton#drawerCloseButton {
                min-width: 28px;
                max-width: 28px;
                min-height: 28px;
                max-height: 28px;
                padding: 0;
                border-radius: 9px;
                background: #f8fafc;
                color: #64748b;
                border: 1px solid #e2e8f0;
                font-size: 15px;
                font-weight: 800;
            }
            QPushButton#drawerCloseButton:hover {
                background: #eef2ff;
                color: #1d4ed8;
                border-color: #c7d2fe;
            }
            QLabel#sheetSummary {
                color: #64748b;
                font-size: 11px;
                font-weight: 600;
                padding: 2px 4px 5px 4px;
            }
            QFrame#compactSourceCard {
                background: #fbfcfe;
                border: 1px solid #e7edf5;
                border-radius: 12px;
            }
            QWidget#sheetShell {
                background: #ffffff;
                border: 1px solid #dfe7f1;
                border-radius: 18px;
            }
            QFrame#sourcePanel, QFrame#sidePanel, QWidget#optionTray {
                background: transparent;
                border: none;
                border-radius: 0;
            }
            QWidget#sheetRail {
                background: #172033;
                border: 1px solid #172033;
                border-radius: 16px;
            }
            QPushButton#sheetRailButton {
                min-width: 30px;
                max-width: 30px;
                min-height: 54px;
                padding: 4px 2px;
                border-radius: 12px;
                background: #243047;
                border: 1px solid #34425b;
                color: #dbeafe;
                font-weight: 700;
            }
            QPushButton#sheetRailButton:checked {
                background: #38bdf8;
                border-color: #38bdf8;
                color: #ffffff;
            }
            QPushButton#sheetRailButton:hover {
                background: #2f3e5a;
            }
            QPushButton#sheetRailButton:checked:hover {
                background: #0ea5e9;
            }
            QLabel {
                color: #172033;
            }
            QLabel#windowTitle {
                font-size: 20px;
                font-weight: 700;
                color: #1d1d1f;
            }
            QLabel#windowSubtitle {
                color: #6e6e73;
            }
            QFrame#tableHeaderBar, QFrame#detailHeaderBar {
                background: #fbfcfe;
                border: 1px solid #edf3f8;
                border-radius: 10px;
            }
            QFrame#preparePanel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ffffff, stop:1 #f3f7ff);
                border: 1px solid #dfe7f1;
                border-radius: 18px;
            }
            QFrame#prepareSourceCard {
                background: rgba(255, 255, 255, 230);
                border: 1px solid #dfe8f3;
                border-radius: 16px;
            }
            QLabel#prepareTitle {
                color: #0f172a;
                font-size: 24px;
                font-weight: 900;
            }
            QLabel#prepareSubtitle {
                color: #64748b;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#prepareCardTitle {
                color: #0f172a;
                font-size: 15px;
                font-weight: 900;
            }
            QLabel#prepareBadge {
                color: #1d4ed8;
                background: #dbeafe;
                border: 1px solid #bfdbfe;
                border-radius: 10px;
                padding: 2px 10px;
                font-size: 11px;
                font-weight: 900;
            }
            QLabel#preparePath {
                color: #64748b;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#prepareFile {
                color: #172033;
                font-size: 13px;
                font-weight: 800;
                background: #f8fafc;
                border: 1px solid #edf2f7;
                border-radius: 12px;
                padding: 10px;
            }
            QLabel#prepareHint {
                color: #64748b;
                font-size: 12px;
                font-weight: 700;
            }
            QPushButton#preparePrimaryButton {
                min-height: 42px;
                padding: 0 28px;
                border-radius: 14px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 14px;
                font-weight: 900;
            }
            QPushButton#preparePrimaryButton:hover {
                background: #1d4ed8;
                border-color: #1d4ed8;
            }
            QPushButton#preparePrimaryButton:disabled {
                background: #cbd5e1;
                border-color: #cbd5e1;
                color: #f8fafc;
            }
            QPushButton#prepareSecondaryButton, QPushButton#prepareChooseButton {
                min-height: 34px;
                padding: 0 18px;
                border-radius: 12px;
                background: #ffffff;
                border: 1px solid #cbd5e1;
                color: #172033;
                font-weight: 800;
            }
            QPushButton#prepareSecondaryButton:hover, QPushButton#prepareChooseButton:hover {
                background: #eff6ff;
                border-color: #93c5fd;
                color: #1d4ed8;
            }
            QComboBox#prepareQuickCombo, QComboBox#prepareRootCombo {
                min-height: 32px;
                background: #ffffff;
                border: 1px solid #d6e0ec;
                border-radius: 11px;
                color: #172033;
                font-weight: 700;
            }
            QLabel#tableTitle, QLabel#detailTitle, QLabel#overviewTitle {
                font-size: 12px;
                font-weight: 800;
                color: #172033;
                padding: 0;
            }
            QLabel#detailTip {
                color: #94a3b8;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton {
                min-height: 30px;
                padding: 0 14px;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                background: #ffffff;
                color: #172033;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #f8fafc;
                border-color: #94a3b8;
            }
            QPushButton:pressed {
                background: #e2e8f0;
            }
            QPushButton#primaryButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-weight: 600;
            }
            QPushButton#primaryButton:hover {
                background: #1d4ed8;
                border-color: #1d4ed8;
            }
            QPushButton#primaryButton:pressed {
                background: #1e40af;
            }
            QPushButton#pendingUpdateButton {
                min-height: 22px;
                max-height: 24px;
                padding: 0 10px;
                border-radius: 11px;
                background: #fff7ed;
                color: #9a3412;
                border: 1px solid #fed7aa;
                font-size: 12px;
                font-weight: 800;
            }
            QPushButton#pendingUpdateButton:hover {
                background: #ffedd5;
                border-color: #fdba74;
            }
            QPushButton#secondaryButton {
                background: #f2f2f7;
                color: #1d1d1f;
                border-color: #d8d8de;
            }
            QPushButton#ghostButton {
                background: rgba(255, 255, 255, 120);
                border: 1px solid #e1e1e6;
                color: #515154;
            }
            QPushButton#ghostButton:hover {
                background: #ffffff;
            }
            QPushButton#commandButton, QPushButton#commandPrimaryButton {
                min-height: 36px;
                max-height: 36px;
                min-width: 72px;
                padding: 0 13px;
                border-radius: 12px;
                font-size: 12px;
                font-weight: 700;
                background: #ffffff;
                border: 1px solid #d6dee9;
                color: #172033;
            }
            QPushButton#commandButton:hover {
                background: #eef6ff;
                border-color: #bfdbfe;
            }
            QPushButton#commandButton:checked {
                background: #dbeafe;
                border-color: #93c5fd;
                color: #1d4ed8;
            }
            QPushButton#commandPrimaryButton {
                background: #2563eb;
                border-color: #2563eb;
                color: #ffffff;
            }
            QPushButton#commandPrimaryButton:hover {
                background: #1d4ed8;
                border-color: #1d4ed8;
            }
            QPushButton#commandPrimaryButton:pressed {
                background: #1e40af;
                border-color: #1e40af;
            }
            QPushButton#pillAdded, QPushButton#pillDeleted, QPushButton#pillConflict,
            QPushButton#pillChanged, QPushButton#pillDuplicate {
                min-height: 22px;
                max-height: 22px;
                padding: 0 9px;
                border-radius: 11px;
                font-size: 12px;
                font-weight: 800;
            }
            QPushButton#pillAdded {
                background: #e9f8ee;
                color: #188038;
                border: 1px solid #c9ecd3;
            }
            QPushButton#pillDeleted {
                background: #fff1f1;
                color: #c62828;
                border: 1px solid #ffd1d1;
            }
            QPushButton#pillConflict {
                background: #fff5e5;
                color: #a15c07;
                border: 1px solid #ffdaa3;
            }
            QPushButton#pillChanged {
                background: #eef5ff;
                color: #006edb;
                border: 1px solid #cfe1ff;
            }
            QPushButton#pillDuplicate {
                background: #f4f0ff;
                color: #5e35b1;
                border: 1px solid #ded4ff;
            }
            QPushButton#pillAdded:disabled, QPushButton#pillDeleted:disabled, QPushButton#pillConflict:disabled,
            QPushButton#pillChanged:disabled, QPushButton#pillDuplicate:disabled {
                background: #f5f5f7;
                color: #a1a1a6;
                border: 1px solid #ebebef;
            }
            QComboBox, QLineEdit {
                min-height: 30px;
                padding: 0 10px;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                background: #ffffff;
                color: #172033;
            }
            QComboBox:hover, QLineEdit:hover {
                border-color: #94a3b8;
            }
            QComboBox:focus, QLineEdit:focus {
                border: 1px solid #2563eb;
            }
            QProgressBar {
                min-height: 14px;
                border: 1px solid #dbe3ee;
                border-radius: 7px;
                background: #f8fafc;
                text-align: center;
                color: #334155;
                font-size: 11px;
                font-weight: 700;
            }
            QProgressBar::chunk {
                border-radius: 6px;
                background: #2563eb;
            }
            QListWidget, QTableView, QTextEdit {
                border: 1px solid #e1e7f0;
                border-radius: 14px;
                background: #ffffff;
                alternate-background-color: #f9fbfd;
                gridline-color: #eef2f7;
                color: #172033;
                selection-background-color: #dbeafe;
                selection-color: #1e3a8a;
            }
            QListWidget::item {
                min-height: 25px;
                padding: 3px 8px;
                border-radius: 9px;
                border: 1px solid transparent;
            }
            QListWidget::item:hover {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
            }
            QListWidget::item:selected {
                background: #dbeafe;
                border: 1px solid #bfdbfe;
                color: #1d4ed8;
            }
            QListWidget#sheetList {
                border: 1px solid #e7edf5;
                border-radius: 14px;
                background: #ffffff;
            }
            QListWidget#sheetList::item {
                margin: 2px 3px;
                padding: 4px 8px;
                border-radius: 8px;
            }
            QWidget#mergeActionBar, QWidget#sheetOverviewBar {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 14px;
            }
            QWidget#sheetOverviewBar {
                background: #ffffff;
                border: 1px solid #e5ebf3;
                border-radius: 14px;
            }
            QFrame#detailPanel {
                background: #ffffff;
                border: 1px solid #e1e7f0;
                border-radius: 16px;
            }
            QTextEdit#detailText {
                border: 1px solid #edf2f7;
                border-radius: 12px;
                background: #ffffff;
                padding: 6px;
                color: #172033;
            }
            QListWidget#duplicateIdList {
                border: 1px solid #edf2f7;
                border-radius: 12px;
                background: #ffffff;
            }
            QHeaderView::section {
                background: #f6f9fc;
                color: #334155;
                padding: 1px 4px;
                border: none;
                border-right: 1px solid #edf2f7;
                border-bottom: 1px solid #dfe7f1;
                font-weight: 700;
            }
            QCheckBox {
                color: #334155;
                spacing: 6px;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
                border-radius: 4px;
                border: 1px solid #c7c7cc;
                background: #ffffff;
            }
            QCheckBox::indicator:checked {
                background: #007aff;
                border: 1px solid #007aff;
            }
            QScrollBar:vertical {
                border: none;
                background: transparent;
                width: 9px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #c7c7cc;
                border-radius: 4px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #8e8e93;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar:horizontal {
                border: none;
                background: transparent;
                height: 9px;
                border-radius: 4px;
            }
            QScrollBar::handle:horizontal {
                background: #c7c7cc;
                border-radius: 4px;
                min-width: 20px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #8e8e93;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            QSplitter::handle {
                background: transparent;
            }
            QSplitter::handle:horizontal {
                width: 7px;
            }
            QSplitter::handle:vertical {
                height: 7px;
            }
            QSplitter::handle:hover {
                background: #e5e5ea;
            }
            """
        )

    def _wire_tables(self) -> None:
        self.left_table.verticalScrollBar().valueChanged.connect(self.middle_table.verticalScrollBar().setValue)
        self.middle_table.verticalScrollBar().valueChanged.connect(self.left_table.verticalScrollBar().setValue)
        self.middle_table.verticalScrollBar().valueChanged.connect(self.right_table.verticalScrollBar().setValue)
        self.right_table.verticalScrollBar().valueChanged.connect(self.middle_table.verticalScrollBar().setValue)
        self.left_table.verticalScrollBar().valueChanged.connect(self.right_table.verticalScrollBar().setValue)
        self.right_table.verticalScrollBar().valueChanged.connect(self.left_table.verticalScrollBar().setValue)
        self.left_table.horizontalScrollBar().valueChanged.connect(self.middle_table.horizontalScrollBar().setValue)
        self.middle_table.horizontalScrollBar().valueChanged.connect(self.left_table.horizontalScrollBar().setValue)
        self.middle_table.horizontalScrollBar().valueChanged.connect(self.right_table.horizontalScrollBar().setValue)
        self.right_table.horizontalScrollBar().valueChanged.connect(self.middle_table.horizontalScrollBar().setValue)
        self.left_table.horizontalScrollBar().valueChanged.connect(self.right_table.horizontalScrollBar().setValue)
        self.right_table.horizontalScrollBar().valueChanged.connect(self.left_table.horizontalScrollBar().setValue)
        self.left_table.clicked.connect(lambda index: self._sync_selection(index.row()))
        self.right_table.clicked.connect(lambda index: self._sync_selection(index.row()))
        for table, side in ((self.left_table, "left"), (self.middle_table, "middle"), (self.right_table, "right")):
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(
                lambda pos, t=table, s=side: self._show_cell_context_menu(t, s, pos)
            )
            header = table.horizontalHeader()
            header.setContextMenuPolicy(Qt.CustomContextMenu)
            header.customContextMenuRequested.connect(
                lambda pos, t=table, s=side: self._show_header_context_menu(t, s, pos)
            )

    def _autoload_defaults(self) -> None:
        set_google_auth_settings(
            str(self.settings.get("google_auth_mode", "service_account")),
            service_account_path=str(self.settings.get("google_service_account_path", "")),
            oauth_client_path=str(self.settings.get("google_oauth_client_path", "")),
            oauth_token_path=str(self.settings.get("google_oauth_token_path", "")),
        )
        self.left_source_combo.currentIndexChanged.connect(lambda: self._on_source_type_changed("left"))
        self.right_source_combo.currentIndexChanged.connect(lambda: self._on_source_type_changed("right"))
        self.left_root_combo.currentTextChanged.connect(lambda text: self._on_root_changed("left", text))
        self.right_root_combo.currentTextChanged.connect(lambda text: self._on_root_changed("right", text))
        self.base_root_combo.currentTextChanged.connect(lambda text: self._on_root_changed("base", text))
        self.left_quick_combo.currentIndexChanged.connect(lambda index: self._apply_side_quick_root("left", index))
        self.right_quick_combo.currentIndexChanged.connect(lambda index: self._apply_side_quick_root("right", index))
        self.base_quick_combo.currentIndexChanged.connect(lambda index: self._apply_side_quick_root("base", index))
        self.left_file_combo.currentTextChanged.connect(lambda text: self._on_file_combo_changed("left", text))
        self.right_file_combo.currentTextChanged.connect(lambda text: self._on_file_combo_changed("right", text))
        self.left_revision_combo.currentIndexChanged.connect(lambda: self._on_revision_changed("left"))
        self.right_revision_combo.currentIndexChanged.connect(lambda: self._on_revision_changed("right"))
        self.template_source_combo.currentIndexChanged.connect(lambda: self._remember_current_settings())
        self.ignore_trim_whitespace_diff.toggled.connect(self._on_compare_text_option_changed)
        self.ignore_all_whitespace_diff.toggled.connect(self._on_compare_text_option_changed)
        self.ignore_case_diff.toggled.connect(self._on_compare_text_option_changed)
        self.normalize_fullwidth_diff.toggled.connect(self._on_compare_text_option_changed)
        self.numeric_tolerance_input.editingFinished.connect(lambda: self._on_compare_text_option_changed(True))
        self.key_fields_input.editingFinished.connect(self._on_manual_key_fields_changed)
        self.key_fields_button.clicked.connect(self._edit_manual_key_fields)
        self.font_size_combo.currentIndexChanged.connect(self._on_font_size_changed)
        self.row_height_combo.currentIndexChanged.connect(self._on_table_spacing_changed)
        self.header_height_combo.currentIndexChanged.connect(self._on_table_spacing_changed)
        self.history_combo.currentIndexChanged.connect(self._apply_history_selection)
        self.advanced_toggle.clicked.connect(lambda: self._toggle_top_panel("settings"))
        self.changed_sheets_only.toggled.connect(lambda: self._refresh_sheet_list(self.current_sheet_name))
        self.three_way_checkbox.toggled.connect(self._on_three_way_toggled)
        self._populate_root_combos()
        self._populate_history_combo()
        self.left_source_combo.setCurrentIndex(max(0, self.left_source_combo.findData(self.settings.get("left_source", SOURCE_LOCAL))))
        self.right_source_combo.setCurrentIndex(max(0, self.right_source_combo.findData(self.settings.get("right_source", SOURCE_LOCAL))))
        self.template_source_combo.setCurrentIndex(max(0, self.template_source_combo.findData(self.settings.get("template_source", "left"))))
        self.ignore_trim_whitespace_diff.blockSignals(True)
        self.ignore_trim_whitespace_diff.setChecked(bool(self.settings.get("ignore_trim_whitespace_diff", False)))
        self.ignore_trim_whitespace_diff.blockSignals(False)
        for checkbox, key in (
            (self.ignore_all_whitespace_diff, "ignore_all_whitespace_diff"),
            (self.ignore_case_diff, "ignore_case_diff"),
            (self.normalize_fullwidth_diff, "normalize_fullwidth_diff"),
        ):
            checkbox.blockSignals(True)
            checkbox.setChecked(bool(self.settings.get(key, False)))
            checkbox.blockSignals(False)
        saved_tolerance = str(self.settings.get("numeric_tolerance_diff", "")).strip()
        self.numeric_tolerance_input.blockSignals(True)
        self.numeric_tolerance_input.setText(saved_tolerance)
        self.numeric_tolerance_input.blockSignals(False)
        saved_key_fields = str(self.settings.get("manual_key_fields", "")).strip()
        self.key_fields_input.blockSignals(True)
        self.key_fields_input.setText(saved_key_fields)
        self.key_fields_input.blockSignals(False)
        self._update_key_field_ui()
        saved_font = int(self.settings.get("ui_font_size", 9) or 9)
        index = max(0, self.font_size_combo.findData(saved_font))
        self.font_size_combo.blockSignals(True)
        self.font_size_combo.setCurrentIndex(index)
        self.font_size_combo.blockSignals(False)
        self._apply_ui_font_size(saved_font)
        saved_row_height = int(self.settings.get("table_row_height", 24) or 24)
        saved_header_height = int(self.settings.get("table_header_height", 32) or 32)
        row_index = max(0, self.row_height_combo.findData(saved_row_height))
        header_index = max(0, self.header_height_combo.findData(saved_header_height))
        self.row_height_combo.blockSignals(True)
        self.header_height_combo.blockSignals(True)
        self.row_height_combo.setCurrentIndex(row_index)
        self.header_height_combo.setCurrentIndex(header_index)
        self.row_height_combo.blockSignals(False)
        self.header_height_combo.blockSignals(False)
        self._apply_table_spacing()
        self._refresh_file_combos()
        self._restore_saved_file_selections()
        self._update_folder_labels()
        self._update_rule_summary()
        self._on_source_type_changed("left")
        self._on_source_type_changed("right")
        self._update_compare_mode()
        # Restore three-way state
        saved_base_path = str(self.settings.get("base_file_path", "")).strip()
        if saved_base_path:
            self.pending_base_path = saved_base_path
            self._update_snapshot_label("base")
        self.three_way_checkbox.blockSignals(True)
        self.three_way_checkbox.setChecked(bool(self.settings.get("three_way_enabled", False)))
        self.three_way_checkbox.blockSignals(False)
        self._on_three_way_toggled(self.three_way_checkbox.isChecked())
        self._update_workspace_mode()
        self.statusBar().showMessage("请选择左右两个 XML 文件，然后点击“开始比对”。", 8000)
        QTimer.singleShot(900, self._check_update_on_startup)

    def _toggle_advanced_panel(self, checked: bool) -> None:
        self._activate_top_panel("settings" if checked else "")

    def _activate_top_panel(self, panel_name: str, *, save: bool = True) -> None:
        if not all(hasattr(self, name) for name in ("top_option_panel", "source_panel", "advanced_panel")):
            return
        active = panel_name if panel_name in {"source", "settings"} else ""
        self.source_panel.setVisible(active == "source")
        self.advanced_panel.setVisible(active == "settings")
        self.top_option_panel.setVisible(bool(active))
        for button, name in (
            (self.source_panel_toggle, "source"),
            (self.advanced_toggle, "settings"),
        ):
            button.blockSignals(True)
            button.setChecked(active == name)
            button.blockSignals(False)
        if save:
            self.settings["active_top_panel"] = active
            self.settings["source_panel_collapsed"] = active != "source"
            save_settings(self.settings)

    def _current_top_panel(self) -> str:
        if getattr(self, "source_panel", None) is not None and not self.source_panel.isHidden():
            return "source"
        if getattr(self, "advanced_panel", None) is not None and not self.advanced_panel.isHidden():
            return "settings"
        return ""

    def _toggle_top_panel(self, panel_name: str) -> None:
        if self._current_top_panel() == panel_name:
            self._activate_top_panel("")
            return
        self._activate_top_panel(panel_name)

    def _activate_side_panel(self, panel_name: str, *, save: bool = True) -> None:
        if not hasattr(self, "sheet_panel"):
            return
        active = "sheet" if panel_name == "sheet" else ""
        self.sheet_panel.setVisible(active == "sheet")
        if hasattr(self, "side_drawer_header"):
            self.side_drawer_header.setVisible(bool(active))
            self.side_drawer_title.setText("子表")
            self.side_drawer_subtitle.setText("工作表导航与变更状态")
        self.sheet_panel_toggle.blockSignals(True)
        self.sheet_panel_toggle.setChecked(active == "sheet")
        self.sheet_panel_toggle.blockSignals(False)
        self._update_side_dock_width()
        if save:
            self.settings["active_side_panel"] = active
            self.settings["sheet_panel_collapsed"] = active != "sheet"
            self.settings["side_dock_width"] = self._last_side_dock_width
            save_settings(self.settings)

    def _current_side_panel(self) -> str:
        if getattr(self, "sheet_panel", None) is not None and not self.sheet_panel.isHidden():
            return "sheet"
        return ""

    def _toggle_side_panel(self, panel_name: str) -> None:
        if self._current_side_panel() == panel_name:
            self._activate_side_panel("")
            return
        self._activate_side_panel(panel_name)

    def _toggle_source_panel(self) -> None:
        self._toggle_top_panel("source")

    def _apply_source_panel_state(self, collapsed: bool, *, save: bool = True) -> None:
        self._activate_top_panel("" if collapsed else "source", save=save)

    def _toggle_sheet_panel(self) -> None:
        self._toggle_side_panel("sheet")

    def _apply_sheet_panel_state(self, collapsed: bool, *, save: bool = True) -> None:
        if not hasattr(self, "sheet_panel") or not hasattr(self, "sheet_shell"):
            return
        self.sheet_panel.setVisible(not collapsed)
        self.sheet_panel_toggle.setChecked(not collapsed)
        self.sheet_panel_toggle.setText("子表")
        self.sheet_panel_toggle.setToolTip("展开左侧子表视图" if collapsed else "隐藏左侧子表视图")
        self._update_side_dock_width()
        if save:
            self.settings["sheet_panel_collapsed"] = collapsed
            save_settings(self.settings)

    def _update_side_dock_width(self) -> None:
        if not hasattr(self, "sheet_shell"):
            return
        panels = [
            getattr(self, "sheet_panel", None),
        ]
        visible = any(panel is not None and not panel.isHidden() for panel in panels)
        self.sheet_shell.setVisible(visible)
        if visible:
            width = self._clamp_side_dock_width(self._last_side_dock_width)
            self.sheet_shell.setMinimumWidth(260)
            self.sheet_shell.setMaximumWidth(520)
            if hasattr(self, "content_splitter"):
                total = max(1, sum(self.content_splitter.sizes()))
                self.content_splitter.setSizes([width, max(700, total - width)])
            return
        self._remember_side_dock_width()
        self.sheet_shell.setMinimumWidth(0)
        self.sheet_shell.setMaximumWidth(0)
        if hasattr(self, "content_splitter"):
            total = max(1, sum(self.content_splitter.sizes()))
            self.content_splitter.setSizes([0, max(700, total)])

    def _clamp_side_dock_width(self, width: int) -> int:
        return max(260, min(520, int(width or 320)))

    def _remember_side_dock_width(self) -> None:
        if not hasattr(self, "content_splitter"):
            return
        sizes = self.content_splitter.sizes()
        if not sizes or sizes[0] <= 0:
            return
        self._last_side_dock_width = self._clamp_side_dock_width(sizes[0])

    def _on_side_dock_splitter_moved(self, pos: int, index: int) -> None:
        del pos, index
        if not getattr(self, "sheet_shell", None) or self.sheet_shell.isHidden():
            return
        self._remember_side_dock_width()
        self.settings["side_dock_width"] = self._last_side_dock_width
        save_settings(self.settings)

    # ── Three-way merge helpers ──────────────────────────────────────────────

    def _on_three_way_toggled(self, checked: bool) -> None:
        self._base_row_container.setVisible(checked)
        self._base_row_label.setVisible(checked)
        self.base_quick_combo.setVisible(checked)
        self.base_root_combo.setVisible(checked)
        self._base_pick_button.setVisible(checked)
        self.base_file_label.setVisible(checked)
        if not checked:
            self.base_workbook = None
        self.settings["three_way_enabled"] = checked
        save_settings(self.settings)

    def _choose_base_file(self) -> None:
        current_root = str(self.base_folder_path or "").strip()
        if current_root:
            self._sync_source_kind_from_root("base", current_root)
        source_kind = self._selected_source_kind("base")
        if source_kind == SOURCE_GOOGLE_SHEETS:
            target = str(self.base_folder_path or "").strip()
            if not target:
                QMessageBox.information(self, "Google Sheets", "请先输入 Base 的 Google Sheets URL 或 Spreadsheet ID。")
                return
            try:
                info = describe_google_sheet(target)
            except Exception as exc:
                QMessageBox.warning(self, "读取 Google Sheets 失败", str(exc))
                return
            self.pending_base_path = target
            self._update_snapshot_label("base")
            self._remember_current_settings()
            QMessageBox.information(
                self,
                "Base Google Sheets 已识别",
                f"标题：{info['title']}\nSpreadsheet ID：{info['spreadsheet_id']}\n工作表数量：{len(info['sheets'])}",
            )
            return

        dialog = DocumentPickerDialog(
            "Base（三方）",
            source_kind,
            current_root,
            self._display_file_name_for_side("base", self.pending_base_path or ""),
            self._selected_revision("base"),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        selected_path = dialog.selected_file_path()
        if not selected_path:
            return
        self.base_folder_path = dialog.root()
        if source_kind == SOURCE_SVN:
            interval = dialog.selected_revision_range()
            if interval is not None:
                QMessageBox.information(self, "提示", "Base 是共同祖先版本，只能选择单个 SVN revision。")
                return
            self.base_revision = dialog.selected_revision()
        else:
            self.base_revision = "WORKING"
        self._update_folder_labels()
        self._set_pending_file("base", selected_path, remember=False, load_revisions=False)
        self._remember_current_settings()

    def _populate_root_combos(self) -> None:
        root_choices = self._combined_root_choices()
        combo_entries = [
            ("left", self.left_root_combo, str(self.left_folder_path or "")),
            ("right", self.right_root_combo, str(self.right_folder_path or "")),
            ("base", self.base_root_combo, str(self.base_folder_path or "")),
        ]
        if hasattr(self, "prepare_left_root_combo"):
            combo_entries.append(("left", self.prepare_left_root_combo, str(self.left_folder_path or "")))
        if hasattr(self, "prepare_right_root_combo"):
            combo_entries.append(("right", self.prepare_right_root_combo, str(self.right_folder_path or "")))
        for side, combo, current in combo_entries:
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(root_choices)
            if current and combo.findText(current) < 0:
                combo.addItem(current)
            combo.setCurrentText(current)
            combo.blockSignals(False)
        self._populate_quick_root_combos()

    def _populate_quick_root_combos(self) -> None:
        quick_roots = list(self.settings.get("quick_roots", []))
        combo_pairs = [
            (self.left_quick_combo, str(self.left_folder_path or "")),
            (self.right_quick_combo, str(self.right_folder_path or "")),
            (self.base_quick_combo, str(self.base_folder_path or "")),
        ]
        if hasattr(self, "prepare_left_quick_combo"):
            combo_pairs.append((self.prepare_left_quick_combo, str(self.left_folder_path or "")))
        if hasattr(self, "prepare_right_quick_combo"):
            combo_pairs.append((self.prepare_right_quick_combo, str(self.right_folder_path or "")))
        for combo, active_root in combo_pairs:
            current_path = active_root or str(combo.currentData() or "")
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("快捷路径", "")
            for entry in quick_roots:
                combo.addItem(str(entry.get("name") or ""), str(entry.get("path") or ""))
            target_index = max(0, combo.findData(current_path))
            combo.setCurrentIndex(target_index)
            combo.setToolTip(str(combo.currentData() or ""))
            combo.blockSignals(False)

    def _apply_prepare_quick_root(self, side: str, index: int) -> None:
        if index <= 0:
            return
        combo = self.prepare_left_quick_combo if side == "left" else self.prepare_right_quick_combo
        root = str(combo.itemData(index) or "").strip()
        if not root:
            return
        self._set_root_value(side, root)
        combo.setToolTip(root)
        side_label = "左" if side == "left" else "右"
        self.statusBar().showMessage(f"已套用{side_label}侧快捷路径，下一步选择文档。", 4000)

    def _apply_prepare_root_text(self, side: str) -> None:
        combo = self.prepare_left_root_combo if side == "left" else self.prepare_right_root_combo
        root = combo.currentText().strip()
        current = str((self.left_folder_path if side == "left" else self.right_folder_path) or "").strip()
        if not root or root == current:
            return
        self._set_root_value(side, root)
        side_label = "左" if side == "left" else "右"
        self.statusBar().showMessage(f"已更新{side_label}侧路径，下一步选择文档。", 4000)

    def _combined_root_choices(self) -> list[str]:
        choices: list[str] = []
        seen: set[str] = set()
        quick_paths = [str(item.get("path") or "").strip() for item in self.settings.get("quick_roots", []) if isinstance(item, dict)]
        for root in [*quick_paths, *self.settings.get("recent_roots", [])]:
            value = str(root or "").strip()
            if not value or value in seen:
                continue
            seen.add(value)
            choices.append(value)
        return choices

    def _populate_history_combo(self) -> None:
        self._history_entries = list(self.settings.get("recent_configs", []))
        self.history_combo.blockSignals(True)
        self.history_combo.clear()
        self.history_combo.addItem("最近配置", None)
        for entry in self._history_entries:
            self.history_combo.addItem(format_config_label(entry), entry)
        self.history_combo.setCurrentIndex(0)
        self.history_combo.blockSignals(False)

    def _restore_saved_file_selections(self) -> None:
        self._restore_saved_file_selection("left")
        self._restore_saved_file_selection("right")

    def _restore_saved_file_selection(self, side: str) -> None:
        saved_path = str(self.settings.get(f"{side}_file_path", "")).strip()
        file_name = str(self.settings.get(f"{side}_file", "")).strip() or source_path_name(saved_path)
        if not file_name and not saved_path:
            return
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        if folder is None:
            return
        combo = self.left_file_combo if side == "left" else self.right_file_combo
        if self._selected_source_kind(side) == SOURCE_SVN and combo.findText(file_name) < 0:
            combo.addItem(file_name)
        if combo.findText(file_name) < 0:
            return
        combo.blockSignals(True)
        combo.setCurrentText(file_name)
        combo.blockSignals(False)
        file_path = saved_path or self._join_selected_target(side, folder, file_name)
        self._set_pending_file(side, file_path, remember=False, load_revisions=False)
        if self._selected_source_kind(side) == SOURCE_SVN:
            self._load_revisions_for_side(side, preferred_revision=str(self.settings.get(f"{side}_revision", "HEAD")))

    def _apply_side_quick_root(self, side: str, index: int) -> None:
        if index <= 0:
            return
        if side == "left":
            combo = self.left_quick_combo
        elif side == "right":
            combo = self.right_quick_combo
        else:
            combo = self.base_quick_combo
        root = str(combo.itemData(index) or "").strip()
        if not root:
            return
        self._set_root_value(side, root)
        combo.setToolTip(root)
        side_label = {"left": "左", "right": "右", "base": "Base"}.get(side, side)
        self.statusBar().showMessage(f"已套用{side_label}侧快捷路径，点击“选择文档”可继续在窗口内细化位置。", 4000)

    def _refresh_sources_caches(self) -> None:
        clear_sources_caches()
        self.statusBar().showMessage("已刷新缓存：下次比对将重新读取工作簿、SVN 版本和 Google 元数据。", 4000)

    def _check_update_on_startup(self) -> None:
        if self._update_check_started:
            return
        self._start_update_check(manual=False)

    def _manual_check_update(self) -> None:
        self._start_update_check(manual=True)

    def _start_update_check(self, *, manual: bool = False) -> None:
        if self._update_thread is not None:
            if manual:
                self.statusBar().showMessage("正在检测更新，请稍候。", 3000)
            return
        if manual:
            self._manual_update_check_running = True
            self.update_check_button.setEnabled(False)
            self.update_check_button.setText("检查中")
            self.statusBar().showMessage("正在检查软件更新。", 3000)
        else:
            self._update_check_started = True
        thread = QThread(self)
        worker = UpdateCheckWorker(fetch_latest=manual)
        worker.moveToThread(thread)
        worker.finished.connect(self._on_update_check_finished)
        worker.failed.connect(self._on_update_check_failed)
        thread.started.connect(worker.run)
        self._update_thread = thread
        self._update_worker = worker
        thread.start()

    def _on_update_check_finished(self, info: object) -> None:
        was_manual = self._manual_update_check_running
        self._teardown_update_worker()
        if hasattr(self, "update_check_button"):
            self.update_check_button.setEnabled(True)
        if not isinstance(info, UpdateInfo):
            if was_manual:
                self._set_update_button_version(APP_VERSION, latest_version="")
                QMessageBox.information(
                    self,
                    "检查更新",
                    f"当前版本：v{APP_VERSION}\n最新版本：未获取到\n\n当前没有可用更新，或更新源未启用。",
                )
            return
        latest_version = info.latest_version
        self._set_update_button_version(APP_VERSION, latest_version=latest_version)
        has_update = compare_versions(latest_version, APP_VERSION) > 0
        if was_manual:
            message = f"当前版本：v{APP_VERSION}\n最新版本：v{latest_version}"
            if not has_update:
                QMessageBox.information(self, "检查更新", message + "\n\n当前已是最新版本。")
                return
            choice = QMessageBox.question(
                self,
                "发现新版本",
                message + "\n\n是否现在下载更新？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if choice != QMessageBox.Yes:
                self.statusBar().showMessage(f"发现新版本 v{latest_version}，已暂不更新。", 5000)
                return
        if has_update:
            update_info = UpdateInfo(
                current_version=APP_VERSION,
                latest_version=info.latest_version,
                release_root=info.release_root,
                version_target=info.version_target,
                source=info.source,
                package_url=info.package_url,
                notes=info.notes,
            )
            self._prompt_apply_update(update_info)

    def _on_update_check_failed(self, exc: object) -> None:
        was_manual = self._manual_update_check_running
        self._teardown_update_worker()
        if hasattr(self, "update_check_button"):
            self.update_check_button.setEnabled(True)
            self.update_check_button.setText(f"v{APP_VERSION}\n更新")
            self.update_check_button.setToolTip(f"当前版本 v{APP_VERSION}\n点击检查最新版本")
        if was_manual:
            QMessageBox.warning(self, "检查更新失败", f"当前版本：v{APP_VERSION}\n\n错误：{exc}")
        self.statusBar().showMessage(f"检测更新失败：{exc}", 6000)

    def _teardown_update_worker(self) -> None:
        thread = self._update_thread
        worker = self._update_worker
        self._update_thread = None
        self._update_worker = None
        self._manual_update_check_running = False
        if thread is not None:
            thread.quit()
            thread.wait(2000)
            thread.deleteLater()
        if worker is not None:
            worker.deleteLater()

    def _set_update_button_version(self, current_version: str, *, latest_version: str = "") -> None:
        if not hasattr(self, "update_check_button"):
            return
        if latest_version:
            self.update_check_button.setText(f"v{current_version}\n最新 {latest_version}")
            self.update_check_button.setToolTip(
                f"当前版本 v{current_version}\n最新版本 v{latest_version}\n点击重新检查"
            )
            return
        self.update_check_button.setText(f"v{current_version}\n更新")
        self.update_check_button.setToolTip(f"当前版本 v{current_version}\n点击检查最新版本")

    def _prompt_apply_update(self, info: UpdateInfo) -> None:
        dialog = UpdateDialog(info, self)
        dialog.install_now_requested.connect(self._launch_prepared_update)
        dialog.prepared_for_later.connect(self._set_pending_update)
        dialog.exec()

    def _set_pending_update(self, prepared: object) -> None:
        if not isinstance(prepared, PreparedUpdate):
            return
        self._pending_prepared_update = prepared
        self._pending_update_button.setVisible(True)
        self.statusBar().showMessage(
            f"v{prepared.info.latest_version} 更新已准备好，可点击右下角“重启安装更新”。",
            8000,
        )

    def _install_pending_update(self) -> None:
        if self._pending_prepared_update is None:
            self._pending_update_button.setVisible(False)
            return
        self._launch_prepared_update(self._pending_prepared_update)

    def _launch_prepared_update(self, prepared: object) -> None:
        if not isinstance(prepared, PreparedUpdate):
            return
        try:
            launch_prepared_update(prepared)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "更新失败", str(exc))
            return
        QMessageBox.information(self, "开始更新", "更新脚本已启动，程序将退出并完成覆盖。")
        QApplication.quit()

    def _manage_quick_roots(self) -> None:
        dialog = QuickRootManagerDialog(
            list(self.settings.get("quick_roots", [])),
            str(self.left_folder_path or ""),
            str(self.right_folder_path or ""),
            str(self.settings.get("google_auth_mode", "service_account")),
            str(self.settings.get("google_service_account_path", "")),
            str(self.settings.get("google_oauth_client_path", "")),
            str(self.settings.get("google_oauth_token_path", "")),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        replace_quick_roots(self.settings, dialog.result_entries())
        google_settings = dialog.result_google_settings()
        self.settings.update(google_settings)
        set_google_auth_settings(
            str(self.settings.get("google_auth_mode", "service_account")),
            service_account_path=str(self.settings.get("google_service_account_path", "")),
            oauth_client_path=str(self.settings.get("google_oauth_client_path", "")),
            oauth_token_path=str(self.settings.get("google_oauth_token_path", "")),
        )
        save_settings(self.settings)
        self._populate_root_combos()
        self.statusBar().showMessage("快捷配置与 Google 认证已更新。", 4000)

    def _on_root_changed(self, side: str, text: str) -> None:
        value = text.strip()
        self._sync_source_kind_from_root(side, value)
        is_google = self._selected_source_kind(side) == SOURCE_GOOGLE_SHEETS
        if side == "left":
            self.left_folder_path = value or None
            self.pending_left_path = value or None if is_google else None
            if not is_google:
                self.left_file_label.setText("左侧快照: 未选择")
                self.left_file_label.setToolTip("")
        elif side == "right":
            self.right_folder_path = value or None
            self.pending_right_path = value or None if is_google else None
            if not is_google:
                self.right_file_label.setText("右侧快照: 未选择")
                self.right_file_label.setToolTip("")
        else:
            self.base_folder_path = value or None
            self.pending_base_path = value or None if is_google else None
            if not is_google:
                self.base_file_label.setText("Base: 未选择")
                self.base_file_label.setToolTip("")
        if side in {"left", "right"}:
            self._refresh_file_combo(side)
        if is_google and value:
            self._update_snapshot_label(side)
        self._remember_current_settings()
        self._update_prepare_panel()

    def _on_file_combo_changed(self, side: str, text: str) -> None:
        if self._suspend_auto_selection:
            return
        file_name = text.strip()
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        if not file_name or folder is None:
            if side == "left":
                self.pending_left_path = None
                self.left_file_label.setText("左侧快照: 未选择")
                self.left_file_label.setToolTip("")
            else:
                self.pending_right_path = None
                self.right_file_label.setText("右侧快照: 未选择")
                self.right_file_label.setToolTip("")
            self._update_prepare_panel()
            return
        file_path = self._join_selected_target(side, folder, file_name)
        if self._selected_source_kind(side) == SOURCE_LOCAL and not Path(file_path).exists():
            return
        self._set_pending_file(side, file_path)

    def _on_revision_changed(self, side: str) -> None:
        self._update_snapshot_label(side)
        self._remember_current_settings()

    def _update_snapshot_label(self, side: str) -> None:
        file_path = self._current_selected_file_path(side)
        if file_path is None:
            self._update_prepare_panel()
            return
        source_kind = self._selected_source_kind(side)
        file_name = self._display_file_name_for_side(side, file_path)
        if source_kind in {SOURCE_LOCAL, SOURCE_GOOGLE_SHEETS}:
            label = file_name
        else:
            label = f"{file_name}@{self._selected_revision(side)}"
        if side == "left":
            self.left_file_label.setText(f"左侧快照: {label}")
            self.left_file_label.setToolTip(file_path)
        elif side == "right":
            self.right_file_label.setText(f"右侧快照: {label}")
            self.right_file_label.setToolTip(file_path)
        else:
            self.base_file_label.setText(f"Base: {label}")
            self.base_file_label.setToolTip(file_path)
        self._update_prepare_panel()

    def _has_active_compare(self) -> bool:
        return self.left_workbook is not None and self.right_workbook is not None

    def _side_snapshot_summary(self, side: str) -> tuple[str, str]:
        file_path = self._current_selected_file_path(side)
        if file_path is None:
            return "未选择文档", ""
        source_kind = self._selected_source_kind(side)
        file_name = self._display_file_name_for_side(side, file_path)
        if source_kind == SOURCE_SVN:
            file_name = f"{file_name}@{self._selected_revision(side)}"
        return file_name, file_path

    @staticmethod
    def _compact_display_path(value: str, *, limit: int = 92) -> str:
        text = str(value or "").strip()
        if not text:
            return "未配置"
        if len(text) <= limit:
            return text
        normalized = text.replace("\\", "/")
        parts = [part for part in normalized.split("/") if part]
        suffix = "/".join(parts[-4:]) if len(parts) >= 4 else text[-limit + 3 :]
        return f".../{suffix}" if len(suffix) + 4 <= limit else f"...{text[-limit + 3:]}"

    def _update_prepare_panel(self) -> None:
        if not hasattr(self, "prepare_panel"):
            return
        for side in ("left", "right"):
            root = self.left_folder_path if side == "left" else self.right_folder_path
            type_label = self.prepare_left_type_label if side == "left" else self.prepare_right_type_label
            file_label = self.prepare_left_file_label if side == "left" else self.prepare_right_file_label
            source_kind = self._selected_source_kind(side)
            snapshot, tooltip = self._side_snapshot_summary(side)
            type_label.setText(SOURCE_LABELS.get(source_kind, source_kind.upper()))
            root_combo = self.prepare_left_root_combo if side == "left" else self.prepare_right_root_combo
            current_root = str(root or "")
            if root_combo.currentText() != current_root:
                root_combo.blockSignals(True)
                if current_root and root_combo.findText(current_root) < 0:
                    root_combo.addItem(current_root)
                root_combo.setCurrentText(current_root)
                root_combo.blockSignals(False)
            root_combo.setToolTip(current_root)
            file_label.setText(f"快照: {snapshot}")
            file_label.setToolTip(tooltip)

        left_snapshot, left_tooltip = self._side_snapshot_summary("left")
        right_snapshot, right_tooltip = self._side_snapshot_summary("right")
        self.source_summary_left_label.setText(f"左: {left_snapshot}")
        self.source_summary_right_label.setText(f"右: {right_snapshot}")
        self.source_summary_left_label.setToolTip(left_tooltip)
        self.source_summary_right_label.setToolTip(right_tooltip)
        self.source_summary_rule_label.setText(f"规则: {self._active_merge_rule().title}")

        left_ready = self._current_selected_file_path("left") is not None
        right_ready = self._current_selected_file_path("right") is not None
        ready = left_ready and right_ready
        if hasattr(self, "prepare_start_button"):
            self.prepare_start_button.setEnabled(ready)
        if hasattr(self, "prepare_match_button"):
            self.prepare_match_button.setEnabled(left_ready)
        if hasattr(self, "prepare_hint_label"):
            self.prepare_hint_label.setText("已选择左右文档，可以开始比对" if ready else "先选择左右两个文档")

    def _update_workspace_mode(self, *, prefer_sheet: bool = False) -> None:
        has_compare = self._has_active_compare()
        if hasattr(self, "prepare_panel"):
            self.prepare_panel.setVisible(not has_compare)
        if hasattr(self, "table_area_splitter"):
            self.table_area_splitter.setVisible(has_compare)
        if hasattr(self, "sheet_overview_bar"):
            self.sheet_overview_bar.setVisible(has_compare)
        if hasattr(self, "source_summary_bar"):
            self.source_summary_bar.setVisible(has_compare)
        if hasattr(self, "filter_strip"):
            self.filter_strip.setVisible(has_compare)
        if hasattr(self, "merge_action_bar"):
            self.merge_action_bar.setVisible(has_compare and not self.compare_only_mode.isChecked())
            self.merge_action_bar.setMaximumHeight(0 if (not has_compare or self.compare_only_mode.isChecked()) else 16777215)
        if not has_compare and self._current_side_panel() == "sheet":
            self._activate_side_panel("", save=False)
        elif has_compare and prefer_sheet and self._current_side_panel() == "":
            self._activate_side_panel("sheet", save=False)
        self._update_prepare_panel()

    def _apply_history_selection(self, index: int) -> None:
        if index <= 0:
            return
        config = self.history_combo.itemData(index)
        if not isinstance(config, dict):
            return
        self._suspend_auto_selection = True
        try:
            self._apply_saved_config(config)
        finally:
            self._suspend_auto_selection = False
        self.history_combo.blockSignals(True)
        self.history_combo.setCurrentIndex(0)
        self.history_combo.blockSignals(False)

    def _apply_saved_config(self, config: dict) -> None:
        self.left_folder_path = str(config.get("left_root") or "").strip() or None
        self.right_folder_path = str(config.get("right_root") or "").strip() or None
        self.base_folder_path = str(config.get("base_root") or self.base_folder_path or "").strip() or None
        self.pending_base_path = str(config.get("base_file_path") or "").strip() or None
        self.base_revision = str(config.get("base_revision") or "HEAD")
        self._populate_root_combos()
        self.left_source_combo.setCurrentIndex(max(0, self.left_source_combo.findData(config.get("left_source", SOURCE_LOCAL))))
        self.right_source_combo.setCurrentIndex(max(0, self.right_source_combo.findData(config.get("right_source", SOURCE_LOCAL))))
        self.template_source_combo.setCurrentIndex(max(0, self.template_source_combo.findData(config.get("template_source", "left"))))
        self.ignore_trim_whitespace_diff.blockSignals(True)
        self.ignore_trim_whitespace_diff.setChecked(bool(config.get("ignore_trim_whitespace_diff", False)))
        self.ignore_trim_whitespace_diff.blockSignals(False)
        manual_key_fields = str(config.get("manual_key_fields", "")).strip()
        self.key_fields_input.blockSignals(True)
        self.key_fields_input.setText(manual_key_fields)
        self.key_fields_input.blockSignals(False)
        sheet_key_fields = config.get("sheet_key_fields")
        if isinstance(sheet_key_fields, dict):
            normalized_sheet_keys: dict[str, list[str]] = {}
            for sheet, field_value in sheet_key_fields.items():
                sheet_name = str(sheet).strip()
                if not sheet_name:
                    continue
                raw_fields = field_value if isinstance(field_value, list) else [field_value]
                fields: list[str] = []
                seen: set[str] = set()
                for raw_field in raw_fields:
                    field = normalize_header(raw_field)
                    if field and field not in seen:
                        fields.append(field)
                        seen.add(field)
                if fields:
                    normalized_sheet_keys[sheet_name] = fields
            self.settings["sheet_key_fields"] = normalized_sheet_keys
        self._update_key_field_ui()
        self._refresh_file_combos()
        self._apply_saved_file_selection(
            "left",
            str(config.get("left_file", "")),
            str(config.get("left_revision", "HEAD")),
            str(config.get("left_file_path", "")),
        )
        self._apply_saved_file_selection(
            "right",
            str(config.get("right_file", "")),
            str(config.get("right_revision", "HEAD")),
            str(config.get("right_file_path", "")),
        )
        if self.pending_base_path:
            self._update_snapshot_label("base")
        self._remember_current_settings()
        self.statusBar().showMessage("最近配置已恢复，可直接开始比对。", 4000)

    def _apply_saved_file_selection(self, side: str, file_name: str, revision: str, file_path_hint: str = "") -> None:
        file_name = file_name or source_path_name(file_path_hint)
        if not file_name:
            return
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        combo = self.left_file_combo if side == "left" else self.right_file_combo
        if self._selected_source_kind(side) == SOURCE_SVN and combo.findText(file_name) < 0:
            combo.addItem(file_name)
        if folder is None or combo.findText(file_name) < 0:
            return
        combo.blockSignals(True)
        combo.setCurrentText(file_name)
        combo.blockSignals(False)
        file_path = file_path_hint or self._join_selected_target(side, folder, file_name)
        self._set_pending_file(side, file_path, remember=False, load_revisions=False)
        if self._selected_source_kind(side) == SOURCE_SVN:
            self._load_revisions_for_side(side, preferred_revision=revision)

    def _refresh_file_combos(self) -> None:
        self._refresh_file_combo("left")
        self._refresh_file_combo("right")

    def _refresh_file_combo(self, side: str) -> None:
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        combo = self.left_file_combo if side == "left" else self.right_file_combo
        current_text = combo.currentText()
        table_files: list[str] = []
        if folder is not None and self._selected_source_kind(side) == SOURCE_LOCAL:
            table_files = [entry.name for entry in list_local_table_files(folder)]

        combo.blockSignals(True)
        combo.clear()
        combo.addItems(table_files)
        if current_text in table_files:
            combo.setCurrentText(current_text)
        combo.blockSignals(False)
        if current_text and current_text not in table_files:
            if side == "left":
                self.pending_left_path = None
                self.left_file_label.setText("左侧快照: 未选择")
                self.left_file_label.setToolTip("")
            else:
                self.pending_right_path = None
                self.right_file_label.setText("右侧快照: 未选择")
                self.right_file_label.setToolTip("")

    def _set_root_value(self, side: str, value: str) -> None:
        if side == "left":
            combo = self.left_root_combo
        elif side == "right":
            combo = self.right_root_combo
        else:
            combo = self.base_root_combo
        combo.setCurrentText(value.strip())

    def _sync_source_kind_from_root(self, side: str, root_value: str) -> None:
        if side == "base":
            return
        source_kind = infer_source_kind(root_value)
        combo = self.left_source_combo if side == "left" else self.right_source_combo
        target_index = combo.findData(source_kind)
        if target_index < 0 or combo.currentIndex() == target_index:
            return
        combo.blockSignals(True)
        combo.setCurrentIndex(target_index)
        combo.blockSignals(False)
        self._on_source_type_changed(side)

    def _selected_source_kind(self, side: str) -> str:
        if side == "base":
            return infer_source_kind(str(self.base_folder_path or self.pending_base_path or ""))
        combo = self.left_source_combo if side == "left" else self.right_source_combo
        return str(combo.currentData())

    def _selected_revision(self, side: str) -> str:
        if side == "base":
            if self._selected_source_kind("base") != SOURCE_SVN:
                return "WORKING"
            return self.base_revision or "HEAD"
        combo = self.left_revision_combo if side == "left" else self.right_revision_combo
        value = str(combo.currentData() or combo.currentText()).strip()
        return value or "HEAD"

    def _revision_combo(self, side: str) -> QComboBox:
        return self.left_revision_combo if side == "left" else self.right_revision_combo

    def _join_selected_target(self, side: str, root: str, file_name: str) -> str:
        if self._selected_source_kind(side) == SOURCE_LOCAL:
            root_path = Path(root)
            if root_path.is_file():
                return str(root_path)
            return str(root_path / file_name)
        if self._selected_source_kind(side) == SOURCE_GOOGLE_SHEETS:
            return root
        return join_source_target(root, file_name)

    def _on_source_type_changed(self, side: str) -> None:
        revision_combo = self._revision_combo(side)
        is_svn = self._selected_source_kind(side) == SOURCE_SVN
        revision_combo.setEnabled(is_svn)
        if is_svn and revision_combo.count() == 0:
            revision_combo.addItem("HEAD")
        if not is_svn:
            revision_combo.clear()
            revision_combo.addItem("WORKING")
        file_path = self._current_selected_file_path(side)
        if is_svn and file_path is not None:
            self._load_revisions_for_side(side)
        else:
            self._update_snapshot_label(side)
        self._remember_current_settings()

    def _current_selected_file_path(self, side: str) -> str | None:
        if side == "base":
            if self.pending_base_path is not None:
                return self.pending_base_path
            if self._selected_source_kind("base") == SOURCE_GOOGLE_SHEETS and self.base_folder_path is not None:
                return str(self.base_folder_path)
            return None
        pending = self.pending_left_path if side == "left" else self.pending_right_path
        if pending is not None:
            return pending
        file_name = self.left_file_combo.currentText() if side == "left" else self.right_file_combo.currentText()
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        if self._selected_source_kind(side) == SOURCE_GOOGLE_SHEETS and folder is not None:
            return str(folder)
        if file_name and folder is not None:
            return self._join_selected_target(side, folder, file_name)
        return None

    def _load_revisions_for_side(self, side: str, preferred_revision: str | None = None) -> None:
        if self._selected_source_kind(side) != SOURCE_SVN:
            return
        file_path = self._current_selected_file_path(side)
        if file_path is None:
            QMessageBox.information(self, "提示", f"请先为{side}侧选择一个本地工作副本中的 XML 文件。")
            return
        try:
            entries = list_svn_revisions(file_path)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 历史失败", str(exc))
            return
        self._revision_entries[side] = entries
        combo = self._revision_combo(side)
        combo.blockSignals(True)
        combo.clear()
        for entry in entries:
            label = entry.revision if entry.revision == "HEAD" else f"r{entry.revision} | {entry.author} | {entry.date}"
            combo.addItem(label, entry.revision)
        if combo.count():
            selected_revision = preferred_revision or str(self.settings.get(f"{side}_revision", "HEAD"))
            target_index = max(0, combo.findData(selected_revision))
            combo.setCurrentIndex(target_index)
        combo.blockSignals(False)
        self._on_revision_changed(side)
        self.statusBar().showMessage(f"{SOURCE_LABELS[SOURCE_SVN]} 历史已载入：{source_path_name(file_path)}", 5000)

    def _update_folder_labels(self) -> None:
        self._set_root_combo_value("left", str(self.left_folder_path) if self.left_folder_path else "")
        self._set_root_combo_value("right", str(self.right_folder_path) if self.right_folder_path else "")
        self._set_root_combo_value("base", str(self.base_folder_path) if self.base_folder_path else "")

    def _set_root_combo_value(self, side: str, value: str) -> None:
        if side == "left":
            combo = self.left_root_combo
        elif side == "right":
            combo = self.right_root_combo
        else:
            combo = self.base_root_combo
        combo.blockSignals(True)
        if value and combo.findText(value) < 0:
            combo.addItem(value)
        combo.setCurrentText(value)
        combo.blockSignals(False)

    def _remember_current_settings(self) -> None:
        left_root = str(self.left_folder_path) if self.left_folder_path else ""
        right_root = str(self.right_folder_path) if self.right_folder_path else ""
        base_root = str(self.base_folder_path) if self.base_folder_path else ""
        self.settings["left_root"] = left_root
        self.settings["right_root"] = right_root
        self.settings["base_root"] = base_root
        self.settings["left_source"] = self._selected_source_kind("left")
        self.settings["right_source"] = self._selected_source_kind("right")
        self.settings["left_file"] = self.left_file_combo.currentText().strip()
        self.settings["right_file"] = self.right_file_combo.currentText().strip()
        self.settings["left_file_path"] = self.pending_left_path or ""
        self.settings["right_file_path"] = self.pending_right_path or ""
        self.settings["left_revision"] = self._selected_revision("left")
        self.settings["right_revision"] = self._selected_revision("right")
        self.settings["template_source"] = str(self.template_source_combo.currentData())
        self.settings["ignore_trim_whitespace_diff"] = self.ignore_trim_whitespace_diff.isChecked()
        self.settings["ignore_all_whitespace_diff"] = self.ignore_all_whitespace_diff.isChecked()
        self.settings["ignore_case_diff"] = self.ignore_case_diff.isChecked()
        self.settings["normalize_fullwidth_diff"] = self.normalize_fullwidth_diff.isChecked()
        self.settings["numeric_tolerance_diff"] = self.numeric_tolerance_input.text().strip()
        self.settings["manual_key_fields"] = ",".join(self._manual_key_fields())
        self.settings["strict_single_id_mode"] = False
        self.settings["sheet_key_fields"] = self._sheet_key_fields_map()
        self.settings["ui_font_size"] = int(self.font_size_combo.currentData() or 9)
        self.settings["table_row_height"] = int(self.row_height_combo.currentData() or 24)
        self.settings["table_header_height"] = int(self.header_height_combo.currentData() or 32)
        if hasattr(self, "source_panel") and hasattr(self, "sheet_panel") and hasattr(self, "advanced_panel"):
            self.settings["active_side_panel"] = "sheet" if self.sheet_panel.isVisible() else ""
            self.settings["active_top_panel"] = self._current_top_panel()
            self.settings["sheet_panel_collapsed"] = not self.sheet_panel.isVisible()
            self.settings["source_panel_collapsed"] = not self.source_panel.isVisible()
            self.settings["side_dock_width"] = self._last_side_dock_width
        self.settings["three_way_enabled"] = self.three_way_checkbox.isChecked()
        self.settings["base_file_path"] = self.pending_base_path or ""
        self.settings["base_revision"] = self._selected_revision("base")
        if left_root:
            remember_root(self.settings, left_root)
        if right_root:
            remember_root(self.settings, right_root)
        if base_root:
            remember_root(self.settings, base_root)
        if self.settings["left_file"] and self.settings["right_file"]:
            remember_config(self.settings, self._current_config_snapshot())
        save_settings(self.settings)
        self._populate_root_combos()
        self._populate_history_combo()

    def _current_config_snapshot(self) -> dict:
        return {
            "left_root": str(self.left_folder_path) if self.left_folder_path else "",
            "right_root": str(self.right_folder_path) if self.right_folder_path else "",
            "base_root": str(self.base_folder_path) if self.base_folder_path else "",
            "left_source": self._selected_source_kind("left"),
            "right_source": self._selected_source_kind("right"),
            "left_file": self.left_file_combo.currentText().strip(),
            "right_file": self.right_file_combo.currentText().strip(),
            "left_file_path": self.pending_left_path or "",
            "right_file_path": self.pending_right_path or "",
            "base_file_path": self.pending_base_path or "",
            "left_revision": self._selected_revision("left"),
            "right_revision": self._selected_revision("right"),
            "base_revision": self._selected_revision("base"),
            "template_source": str(self.template_source_combo.currentData()),
            "ignore_trim_whitespace_diff": self.ignore_trim_whitespace_diff.isChecked(),
            "manual_key_fields": ",".join(self._manual_key_fields()),
            "strict_single_id_mode": False,
            "sheet_key_fields": self._sheet_key_fields_map(),
            "table_row_height": int(self.row_height_combo.currentData() or 24),
            "table_header_height": int(self.header_height_combo.currentData() or 32),
        }

    def _active_merge_rule(self):
        return get_merge_rule(self.rule_combo.currentData())

    def _manual_key_fields(self) -> list[str]:
        raw_text = self.key_fields_input.text().strip() if hasattr(self, "key_fields_input") else ""
        if not raw_text:
            return []
        for delimiter in ("\n", "；", ";", "，"):
            raw_text = raw_text.replace(delimiter, ",")
        values: list[str] = []
        seen: set[str] = set()
        for part in raw_text.split(","):
            if not part.strip():
                continue
            normalized = normalize_header(part)
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            values.append(normalized)
        return values

    def _strict_single_key_enabled(self) -> bool:
        return False

    def _sheet_key_fields_map(self) -> dict[str, list[str]]:
        value = self.settings.get("sheet_key_fields", {}) if hasattr(self, "settings") else {}
        if not isinstance(value, dict):
            return {}
        result: dict[str, list[str]] = {}
        for sheet_name, field_value in value.items():
            sheet = str(sheet_name or "").strip()
            fields: list[str] = []
            raw_fields = field_value if isinstance(field_value, list) else [field_value]
            seen: set[str] = set()
            for raw_field in raw_fields:
                field = normalize_header(raw_field)
                if field and field not in seen:
                    fields.append(field)
                    seen.add(field)
            if sheet and fields:
                result[sheet] = fields
        return result

    def _key_fields_for_sheet(self, sheet_name: str | None) -> list[str]:
        if sheet_name:
            sheet_keys = self._sheet_key_fields_map().get(str(sheet_name))
            if sheet_keys:
                return sheet_keys
        return self._manual_key_fields()

    def _strict_single_key_enabled_for_sheet(self, sheet_name: str | None) -> bool:
        del sheet_name
        return False

    def _update_key_field_ui(self) -> None:
        if not hasattr(self, "key_fields_button"):
            return
        self.key_fields_button.setText("选择字段")
        self.key_fields_input.setPlaceholderText("手动关键字段，逗号分隔，例如 skillid,level")
        self.key_fields_input.setToolTip("留空时自动推断；填写后优先按这些逻辑字段对齐。当前子表可右键表头标记 ID 列，该设置优先于这里的通用关键字段。")

    def _candidate_key_fields(self) -> list[str]:
        if self.current_sheet_name and (self.left_workbook or self.right_workbook):
            left_sheet = self.left_workbook.get_sheet(self.current_sheet_name) if self.left_workbook else None
            right_sheet = self.right_workbook.get_sheet(self.current_sheet_name) if self.right_workbook else None
            if left_sheet is not None and right_sheet is not None:
                right_headers = {
                    normalize_header(header)
                    for header in right_sheet.logical_headers
                    if str(header).strip()
                }
                candidates: list[str] = []
                seen: set[str] = set()
                for header in left_sheet.logical_headers:
                    if not str(header).strip():
                        continue
                    normalized = normalize_header(header)
                    if normalized and normalized in right_headers and normalized not in seen:
                        candidates.append(normalized)
                        seen.add(normalized)
                return candidates
        return []

    def _edit_manual_key_fields(self) -> None:
        current_fields = self._manual_key_fields()
        candidates = self._candidate_key_fields()
        if candidates:
            dialog = KeyFieldsDialog(
                candidates,
                current_fields,
                self,
                single_mode=False,
            )
            if dialog.exec() != QDialog.Accepted:
                return
            self.key_fields_input.setText(",".join(dialog.selected_fields()))
            self._on_manual_key_fields_changed()
            return

        current_value = ",".join(current_fields)
        text, ok = QInputDialog.getText(
            self,
            "手动关键字段",
            "当前还没有可读取的左右表字段。按逗号分隔填写逻辑字段名，留空表示自动推断。",
            text=current_value,
        )
        if not ok:
            return
        self.key_fields_input.setText(text)
        self._on_manual_key_fields_changed()

    def _on_manual_key_fields_changed(self) -> None:
        fields = self._manual_key_fields()
        normalized = ",".join(fields)
        if self.key_fields_input.text().strip() != normalized:
            self.key_fields_input.blockSignals(True)
            self.key_fields_input.setText(normalized)
            self.key_fields_input.blockSignals(False)
        self._remember_current_settings()
        if self.left_workbook is None and self.right_workbook is None:
            return
        self.statusBar().showMessage("关键字段设置已更新，下一次点击“开始比对”会按新字段重新对齐。", 5000)

    def _comparison_options(self) -> ComparisonOptions:
        tolerance_text = self.numeric_tolerance_input.text().strip() if hasattr(self, "numeric_tolerance_input") else ""
        try:
            tolerance = float(tolerance_text) if tolerance_text else 0.0
        except ValueError:
            tolerance = 0.0
        return ComparisonOptions(
            ignore_trim_whitespace=self.ignore_trim_whitespace_diff.isChecked(),
            ignore_all_whitespace=getattr(self, "ignore_all_whitespace_diff", None) is not None and self.ignore_all_whitespace_diff.isChecked(),
            ignore_case=getattr(self, "ignore_case_diff", None) is not None and self.ignore_case_diff.isChecked(),
            normalize_fullwidth=getattr(self, "normalize_fullwidth_diff", None) is not None and self.normalize_fullwidth_diff.isChecked(),
            numeric_tolerance=max(0.0, tolerance),
        )

    def _update_rule_summary(self) -> None:
        rule = self._active_merge_rule()
        summary = f"{rule.title}：{rule.summary}"
        self.rule_summary_label.setText(summary)
        self.rule_combo.setToolTip(summary)

    def _on_rule_changed(self) -> None:
        self._update_rule_summary()
        self._update_prepare_panel()
        if self.left_workbook is None and self.right_workbook is None:
            return
        current_sheet = self.current_sheet_name
        self.alignments.clear()
        self._sheet_state_cache.clear()
        self._refresh_sheet_list(current_sheet)
        self.statusBar().showMessage(f"默认规则已切换为“{self._active_merge_rule().title}”。", 5000)

    def _on_compare_text_option_changed(self, checked: bool) -> None:
        self._remember_current_settings()
        if self.left_workbook is None and self.right_workbook is None:
            return
        current_sheet = self.current_sheet_name
        self.alignments.clear()
        self._sheet_state_cache.clear()
        self._refresh_sheet_list(current_sheet)
        status = "已开启忽略首尾空白差异" if checked else "已恢复严格文本比较"
        self.statusBar().showMessage(status, 5000)

    def _update_compare_mode(self) -> None:
        compare_only = self.compare_only_mode.isChecked()
        has_compare = self._has_active_compare()
        if has_compare and not compare_only and self.middle_model is None and self.current_sheet_name:
            alignment = self.alignments.get(self.current_sheet_name)
            if alignment is not None:
                self._ensure_middle_model(alignment)
                self._apply_filters()
                self._schedule_resize_columns()
                self._update_detail_panel()
        if self.middle_panel is not None:
            self.middle_panel.setVisible(has_compare and not compare_only)
        if self.merge_action_bar is not None:
            self.merge_action_bar.setVisible(has_compare and not compare_only)
            self.merge_action_bar.setMaximumHeight(0 if (not has_compare or compare_only) else 16777215)
        if has_compare and getattr(self, "tables_splitter", None) is not None:
            sizes = self.tables_splitter.sizes()
            total = sum(sizes) or 1200
            if compare_only:
                self.tables_splitter.setSizes([total // 2, 0, total - total // 2])
            else:
                self.tables_splitter.setSizes([int(total * 0.31), int(total * 0.38), total - int(total * 0.31) - int(total * 0.38)])
        self._update_workspace_mode()
        mode_text = "纯比对模式：默认只加载左右对比，切到合并模式时再加载中间合并表。" if compare_only else "合并模式：已加载预合并结果，可直接人工处理。"
        self.statusBar().showMessage(mode_text, 4000)

    def _choose_file(self, side: str) -> None:
        current_root = str(self.left_folder_path if side == "left" else self.right_folder_path or "").strip()
        if current_root:
            self._sync_source_kind_from_root(side, current_root)
        source_kind = self._selected_source_kind(side)
        if source_kind == SOURCE_GOOGLE_SHEETS:
            target = str(self.left_folder_path if side == "left" else self.right_folder_path or "").strip()
            if not target:
                QMessageBox.information(self, "Google Sheets", "请先输入 Google Sheets URL 或 Spreadsheet ID。")
                return
            try:
                info = describe_google_sheet(target)
            except Exception as exc:
                QMessageBox.warning(self, "读取 Google Sheets 失败", str(exc))
                return
            if side == "left":
                self.pending_left_path = target
            else:
                self.pending_right_path = target
            self._update_snapshot_label(side)
            self._remember_current_settings()
            summary = "\n".join(
                f"- {item['title']} ({item['row_count']} 行 / {item['column_count']} 列)"
                for item in info["sheets"][:8]
            )
            QMessageBox.information(
                self,
                "Google Sheets 已识别",
                f"标题：{info['title']}\nSpreadsheet ID：{info['spreadsheet_id']}\n"
                f"工作表数量：{len(info['sheets'])}\n\n{summary}",
            )
            return
        side_label = "左侧" if side == "left" else "右侧"
        dialog = DocumentPickerDialog(
            side_label,
            source_kind,
            str(self.left_folder_path if side == "left" else self.right_folder_path or ""),
            self.left_file_combo.currentText() if side == "left" else self.right_file_combo.currentText(),
            self._selected_revision(side),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        selected_path = dialog.selected_file_path()
        if not selected_path:
            return
        root = dialog.root()
        interval = dialog.selected_revision_range() if source_kind == SOURCE_SVN else None
        if interval is not None:
            self._apply_svn_revision_interval(selected_path, root, interval[0], interval[1])
            return
        if side == "left":
            self.left_folder_path = root
        else:
            self.right_folder_path = root
        self._update_folder_labels()
        self._refresh_file_combo(side)
        self._set_pending_file(side, selected_path, remember=False, load_revisions=False)
        if source_kind == SOURCE_SVN:
            self._load_revisions_for_side(side, preferred_revision=dialog.selected_revision())
        else:
            self._update_snapshot_label(side)
        self._remember_current_settings()

    def _apply_svn_revision_interval(self, file_path: str, root: str, older_revision: str, newer_revision: str) -> None:
        for side, combo in (("left", self.left_source_combo), ("right", self.right_source_combo)):
            target_index = combo.findData(SOURCE_SVN)
            if target_index >= 0:
                combo.blockSignals(True)
                combo.setCurrentIndex(target_index)
                combo.blockSignals(False)
                self._on_source_type_changed(side)
        self.left_folder_path = root
        self.right_folder_path = root
        self._update_folder_labels()
        self._refresh_file_combo("left")
        self._refresh_file_combo("right")
        self._set_pending_file("left", file_path, remember=False, load_revisions=False)
        self._set_pending_file("right", file_path, remember=False, load_revisions=False)
        self._load_revisions_for_side("left", preferred_revision=older_revision)
        self._load_revisions_for_side("right", preferred_revision=newer_revision)
        self._remember_current_settings()
        self.statusBar().showMessage(
            f"已设置 SVN diff 区间：{source_path_name(file_path)} | 左侧={older_revision} -> 右侧={newer_revision}",
            8000,
        )

    def _select_from_combo(self, side: str) -> None:
        file_name = self.left_file_combo.currentText() if side == "left" else self.right_file_combo.currentText()
        folder = self.left_folder_path if side == "left" else self.right_folder_path
        if not file_name:
            QMessageBox.information(self, "提示", f"{side}侧当前路径里没有可选择的文件。")
            return
        if folder is None:
            QMessageBox.information(self, "提示", f"请先锁定{side}侧文件夹。")
            return
        self._set_pending_file(side, self._join_selected_target(side, folder, file_name))

    def _match_right_file_by_left_name(self) -> None:
        left_name = self.left_file_combo.currentText().strip() or source_path_name(self.pending_left_path or "")
        if not left_name:
            QMessageBox.information(self, "提示", "请先在左侧选择一个文件名。")
            return
        if self._selected_source_kind("right") == SOURCE_LOCAL:
            index = self.right_file_combo.findText(left_name)
            if index < 0:
                QMessageBox.information(self, "提示", f"右侧文件夹中不存在同名文件：{left_name}")
                return
            self.right_file_combo.setCurrentIndex(index)
            return
        root = self.right_folder_path
        if not root:
            QMessageBox.information(self, "提示", "请先填写右侧 SVN 根路径。")
            return
        try:
            candidates = list_svn_xml_files(root)
        except Exception as exc:
            QMessageBox.warning(self, "读取 SVN 文件失败", str(exc))
            return
        matched = next((item for item in candidates if source_relative_path(root, item.path) == left_name), None)
        if matched is None:
            left_basename = source_path_name(left_name)
            matched = next((item for item in candidates if item.name == left_basename), None)
        if matched is None:
            QMessageBox.information(self, "提示", f"右侧 SVN 路径下不存在同名文件：{left_name}")
            return
        self._set_pending_file("right", matched.path)

    def _set_pending_file(self, side: str, file_path: str, *, remember: bool = True, load_revisions: bool = True) -> None:
        source_kind = self._selected_source_kind(side)
        if side == "left":
            self.pending_left_path = file_path
        elif side == "right":
            self.pending_right_path = file_path
        else:
            self.pending_base_path = file_path
        file_name = self._display_file_name_for_side(side, file_path)
        if side in {"left", "right"}:
            combo = self.left_file_combo if side == "left" else self.right_file_combo
            if combo.findText(file_name) < 0:
                combo.addItem(file_name)
            combo.blockSignals(True)
            combo.setCurrentText(file_name)
            combo.blockSignals(False)
        if source_kind == SOURCE_SVN and load_revisions:
            self._load_revisions_for_side(side)
        else:
            self._update_snapshot_label(side)
        if remember:
            self._remember_current_settings()
        side_label = {"left": "左", "right": "右", "base": "Base"}.get(side, side)
        self.statusBar().showMessage(
            f"已选择{side_label}侧{SOURCE_LABELS[source_kind]}文件: {file_name}。点击“开始比对”后才会真正加载。",
            5000,
        )

    def _display_file_name_for_side(self, side: str, file_path: str) -> str:
        source_kind = self._selected_source_kind(side)
        if source_kind == SOURCE_SVN:
            if side == "left":
                root = self.left_folder_path
            elif side == "right":
                root = self.right_folder_path
            else:
                root = self.base_folder_path
            if root:
                return source_relative_path(root, file_path)
        return source_path_name(file_path)

    def _start_compare(self) -> None:
        if self._current_selected_file_path("left") is None or self._current_selected_file_path("right") is None:
            QMessageBox.information(self, "提示", "请先把左右两个文件都选好，再开始比对。")
            return
        if self._load_threads:
            QMessageBox.information(self, "提示", "上一次加载还在进行中，请稍候或取消。")
            return

        three_way = self.three_way_checkbox.isChecked()
        base_path = self.pending_base_path if three_way else None
        if three_way and not base_path:
            QMessageBox.information(self, "提示", "启用三方合并时需先选择 Base 文件。")
            return

        try:
            left_source = self._build_workbook_source("left")
            right_source = self._build_workbook_source("right")
            base_source = self._build_base_workbook_source(base_path) if base_path else None
        except Exception as exc:
            self._show_friendly_error("准备加载失败", exc)
            return

        self._compare_started_at = perf_counter()
        self._pending_workbooks = {"left": None, "right": None}
        if base_source is not None:
            self._pending_workbooks["base"] = None
        self._pending_errors = {}
        self._load_cancelled = False

        progress = QProgressDialog("正在加载…", "取消", 0, 0, self)
        progress.setWindowTitle("加载中")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.reset()
        progress.hide()
        progress.canceled.connect(self._on_load_cancelled)
        self._load_progress = progress
        QTimer.singleShot(200, self._maybe_show_load_progress)

        sources = [("left", left_source), ("right", right_source)]
        if base_source is not None:
            sources.append(("base", base_source))
        for side, source in sources:
            thread = QThread(self)
            worker = WorkbookLoadWorker(side, source)
            worker.moveToThread(thread)
            worker.stage.connect(self._on_load_stage)
            worker.finished.connect(self._on_load_finished)
            worker.failed.connect(self._on_load_failed)
            thread.started.connect(worker.run)
            self._load_threads[side] = thread
            self._load_workers[side] = worker
            thread.start()

    def _maybe_show_load_progress(self) -> None:
        if self._load_cancelled or not self._load_threads:
            return
        if self._load_progress is not None:
            self._load_progress.show()

    def _on_load_stage(self, side: str, text: str) -> None:
        if self._load_progress is None:
            return
        label = {"left": "左", "right": "右", "base": "Base"}.get(side, side)
        self._load_progress.setLabelText(f"{label}侧：{text}")

    def _on_load_finished(self, side: str, workbook: object) -> None:
        if isinstance(workbook, WorkbookData):
            self._pending_workbooks[side] = workbook
        self._teardown_worker(side)
        self._maybe_finalize_compare()

    def _on_load_failed(self, side: str, exc: object) -> None:
        if isinstance(exc, BaseException):
            self._pending_errors[side] = exc
        self._teardown_worker(side)
        self._maybe_finalize_compare()

    def _on_load_cancelled(self) -> None:
        self._load_cancelled = True
        self.statusBar().showMessage("已请求取消加载，后台任务会跑完后自动释放。", 4000)

    def _teardown_worker(self, side: str) -> None:
        thread = self._load_threads.pop(side, None)
        worker = self._load_workers.pop(side, None)
        if thread is not None:
            thread.quit()
            thread.wait(2000)
            thread.deleteLater()
        if worker is not None:
            worker.deleteLater()

    def _maybe_finalize_compare(self) -> None:
        if self._load_threads:
            return
        was_cancelled = self._load_cancelled
        progress = self._load_progress
        self._load_progress = None
        if progress is not None:
            try:
                progress.canceled.disconnect(self._on_load_cancelled)
            except (RuntimeError, TypeError):
                pass
            progress.close()
            progress.deleteLater()

        if was_cancelled:
            self._pending_workbooks = {}
            self._pending_errors = {}
            self.statusBar().showMessage("已取消加载。", 4000)
            return

        if self._pending_errors:
            side, exc = next(iter(self._pending_errors.items()))
            hint = "常见原因：路径错误、SVN 鉴权失败、Google 认证凭据未配置、文件格式不受支持。"
            side_label = "左侧" if side == "left" else "右侧"
            self._show_friendly_error(f"{side_label}加载失败", exc, hint=hint)
            self._pending_workbooks = {}
            self._pending_errors = {}
            return

        left_wb = self._pending_workbooks.get("left")
        right_wb = self._pending_workbooks.get("right")
        base_wb = self._pending_workbooks.get("base")
        self._pending_workbooks = {}
        if left_wb is None or right_wb is None:
            return
        self._finalize_compare_with_workbooks(left_wb, right_wb, base_wb)

    def _finalize_compare_with_workbooks(
        self, left_workbook: WorkbookData, right_workbook: WorkbookData, base_workbook: WorkbookData | None = None
    ) -> None:
        source_loaded_at = perf_counter()
        self.left_workbook = left_workbook
        self.right_workbook = right_workbook
        self.base_workbook = base_workbook
        self.alignments.clear()
        self._sheet_state_cache.clear()
        self._column_width_cache.clear()
        self._duplicate_id_rows = []
        self._large_diff_confirmed_sheets.clear()
        self._cancel_sheet_state_warmup()
        self._remember_current_settings()
        self._refresh_sheet_list(schedule_warmup=False)
        self._update_workspace_mode(prefer_sheet=True)
        first_sheet_ready_at = perf_counter()
        self._last_compare_metrics = {
            "load_ms": (source_loaded_at - self._compare_started_at) * 1000.0,
            "first_ready_ms": (first_sheet_ready_at - self._compare_started_at) * 1000.0,
        }
        self._schedule_resize_columns()
        QTimer.singleShot(120, lambda: self._schedule_sheet_state_warmup())
        mode_suffix = " [三方合并]" if base_workbook is not None else ""
        self.statusBar().showMessage(
            "比对已开始：{left} <-> {right}{mode}，规则 {rule} | 源加载 {load_ms} | 首屏 {first_ready_ms}".format(
                left=left_workbook.source_label,
                right=right_workbook.source_label,
                mode=mode_suffix,
                rule=self._active_merge_rule().title,
                load_ms=self._format_duration_ms(self._last_compare_metrics["load_ms"]),
                first_ready_ms=self._format_duration_ms(self._last_compare_metrics["first_ready_ms"]),
            ),
            8000,
        )

    def _build_workbook_source(self, side: str) -> WorkbookSource:
        file_path = self._current_selected_file_path(side)
        if file_path is None:
            raise ValueError(f"{side}侧还没有选中文件。")
        source_kind = self._selected_source_kind(side)
        if source_kind == SOURCE_LOCAL:
            local_path = Path(file_path)
            return WorkbookSource(
                kind=SOURCE_LOCAL,
                file_path=str(file_path),
                display_name=local_path.name,
                source_root=str(local_path.parent),
            )
        if source_kind == SOURCE_SVN:
            revision = self._selected_revision(side)
            if side == "left":
                source_root = self.left_folder_path
            elif side == "right":
                source_root = self.right_folder_path
            else:
                source_root = self.base_folder_path
            return WorkbookSource(
                kind=SOURCE_SVN,
                file_path=str(file_path),
                revision=revision,
                display_name=f"{source_path_name(file_path)}@{revision}",
                source_root=str(source_root or ""),
            )
        if source_kind == SOURCE_GOOGLE_SHEETS:
            parsed = parse_google_sheet_target(file_path)
            display_name = parsed["spreadsheet_id"]
            if parsed["gid"]:
                display_name = f"{display_name}#gid={parsed['gid']}"
            metadata = (("gid", parsed["gid"]),) if parsed["gid"] else ()
            return WorkbookSource(
                kind=SOURCE_GOOGLE_SHEETS,
                file_path=str(file_path),
                display_name=display_name,
                source_root=str(file_path),
                metadata=metadata,
            )
        raise ValueError(f"暂不支持的来源类型: {source_kind}")

    def _build_base_workbook_source(self, file_path: str) -> WorkbookSource:
        del file_path
        return self._build_workbook_source("base")

    def _refresh_sheet_list(self, preferred_sheet: str | None = None, *, schedule_warmup: bool = True) -> None:
        self._cancel_sheet_state_warmup()
        self.sheet_list.blockSignals(True)
        self.sheet_list.clear()
        sheet_names = self._available_sheet_names()
        if self.changed_sheets_only.isChecked():
            self._ensure_sheet_states(sheet_names)
        for sheet_name in sheet_names:
            state = self._sheet_state_cache.get(sheet_name)
            if self.changed_sheets_only.isChecked():
                state = self._sheet_state(sheet_name)
            if self.changed_sheets_only.isChecked() and state is not None and not bool(state["changed"]):
                continue
            title = sheet_name
            item = QListWidgetItem(title)
            item.setData(Qt.UserRole, sheet_name)
            self._apply_sheet_item_state(item, sheet_name, state)
            self.sheet_list.addItem(item)
        self._filter_sheet_list(self.sheet_filter_input.text() if hasattr(self, "sheet_filter_input") else "")
        self._update_sheet_summary()
        target_row = self._visible_sheet_row(preferred_sheet)
        self.sheet_list.blockSignals(False)
        if target_row is not None:
            self.sheet_list.setCurrentRow(target_row)
            current = self.sheet_list.currentItem()
            if current is not None:
                self._on_sheet_changed(current, None)
        else:
            self.sheet_list.setCurrentRow(-1)
            self.current_sheet_name = None
            message = "当前筛选下没有子表。" if sheet_names else "请先选择左右文件并开始比对。"
            self.detail_panel.setHtml(f"<b>{message}</b>")
        if schedule_warmup and not self.changed_sheets_only.isChecked():
            self._schedule_sheet_state_warmup(sheet_names)

    def _show_cell_context_menu(self, table: QTableView, side: str, pos) -> None:
        if self.middle_model is None or self.current_sheet_name is None:
            return
        index = table.indexAt(pos)
        if not index.isValid():
            return
        if self.compare_only_mode.isChecked():
            self.compare_only_mode.setChecked(False)
            if self.middle_model is None:
                return
        visible_row = index.row()
        logical_column = index.column() + 1
        source_index = self.middle_model.source_row_index(visible_row)
        if source_index >= len(self.middle_model.alignment.rows):
            return
        row = self.middle_model.alignment.rows[source_index]
        binding = self.middle_model.alignment.columns[index.column()]
        menu = QMenu(table)
        left_value = row.left_row.value_at(binding.left_index or -1) if row.left_row and binding.left_index else ""
        right_value = row.right_row.value_at(binding.right_index or -1) if row.right_row and binding.right_index else ""
        current_value = row.merged_row.value_at(logical_column)

        def _truncate(text: str) -> str:
            snippet = (text or "").replace("\n", " ")
            return (snippet[:20] + "…") if len(snippet) > 20 else snippet

        left_action = menu.addAction(f"此单元格取左值：{_truncate(left_value) or '(空)'}")
        left_action.setEnabled(bool(row.left_row) and left_value != current_value)
        right_action = menu.addAction(f"此单元格取右值：{_truncate(right_value) or '(空)'}")
        right_action.setEnabled(bool(row.right_row) and right_value != current_value)
        clear_action = menu.addAction("此单元格清空")
        menu.addSeparator()
        take_left_row = menu.addAction("整行取左")
        take_right_row = menu.addAction("整行取右")
        reset_row = menu.addAction("整行重置自动")
        delete_merged_row = menu.addAction("删除合并行")
        delete_merged_row.setEnabled(side == "middle")
        menu.addSeparator()
        copy_action = menu.addAction("复制该单元格")

        chosen = menu.exec(table.viewport().mapToGlobal(pos))
        if chosen is None:
            return
        if chosen is left_action:
            self._set_cell_value(row, logical_column, left_value, f"取左列 {binding.title}")
        elif chosen is right_action:
            self._set_cell_value(row, logical_column, right_value, f"取右列 {binding.title}")
        elif chosen is clear_action:
            self._set_cell_value(row, logical_column, "", f"清空列 {binding.title}")
        elif chosen is take_left_row:
            self._select_rows([visible_row], source=None)
            self._apply_row_choice("left")
        elif chosen is take_right_row:
            self._select_rows([visible_row], source=None)
            self._apply_row_choice("right")
        elif chosen is reset_row:
            self._select_rows([visible_row], source=None)
            self._apply_row_choice("auto")
        elif chosen is delete_merged_row:
            self._delete_merged_rows([visible_row])
        elif chosen is copy_action:
            clipboard = QApplication.clipboard()
            if clipboard is not None:
                source_value = left_value if side == "left" else right_value if side == "right" else current_value
                clipboard.setText(source_value)
                self.statusBar().showMessage("已复制单元格内容到剪贴板。", 3000)

    def _show_header_context_menu(self, table: QTableView, side: str, pos) -> None:
        del side
        if self.current_sheet_name is None:
            return
        model = table.model()
        if not isinstance(model, MergeTableModel):
            return
        logical_section = table.horizontalHeader().logicalIndexAt(pos)
        if logical_section < 0 or logical_section >= len(model.alignment.columns):
            return
        binding = model.alignment.columns[logical_section]
        field_key = normalize_header(binding.key)
        if not field_key:
            return
        sheet_name = str(self.current_sheet_name)
        current_sheet_keys = self._sheet_key_fields_map().get(sheet_name, [])
        is_selected_key = field_key in current_sheet_keys
        menu = QMenu(table)
        set_action = menu.addAction(
            f"{'取消' if is_selected_key else '加入'}组合 ID 字段：{binding.title}"
        )
        clear_action = menu.addAction("清空当前子表组合 ID")
        clear_action.setEnabled(bool(current_sheet_keys))
        chosen = menu.exec(table.horizontalHeader().mapToGlobal(pos))
        if chosen is set_action:
            if is_selected_key:
                self._remove_sheet_id_check_field(sheet_name, field_key)
            else:
                self._add_sheet_id_check_field(sheet_name, field_key)
        elif chosen is clear_action:
            self._clear_sheet_id_check_field(sheet_name)

    def _add_sheet_id_check_field(self, sheet_name: str, field_key: str) -> None:
        sheet_keys = self._sheet_key_fields_map()
        normalized = normalize_header(field_key)
        fields = list(sheet_keys.get(sheet_name, []))
        if normalized and normalized not in fields:
            fields.append(normalized)
        sheet_keys[sheet_name] = fields
        self.settings["sheet_key_fields"] = sheet_keys
        self._remember_current_settings()
        self._rebuild_sheet_alignment(sheet_name)
        self.statusBar().showMessage(
            f"已更新 {sheet_name} 的组合 ID：{', '.join(fields)}。非空组合键优先对齐；空 ID 行仍按原规则匹配，不会强制排序。",
            7000,
        )

    def _remove_sheet_id_check_field(self, sheet_name: str, field_key: str) -> None:
        sheet_keys = self._sheet_key_fields_map()
        normalized = normalize_header(field_key)
        fields = [field for field in sheet_keys.get(sheet_name, []) if field != normalized]
        if fields:
            sheet_keys[sheet_name] = fields
        else:
            sheet_keys.pop(sheet_name, None)
        self.settings["sheet_key_fields"] = sheet_keys
        self._remember_current_settings()
        self._rebuild_sheet_alignment(sheet_name)
        label = ", ".join(fields) if fields else "未设置"
        self.statusBar().showMessage(f"已更新 {sheet_name} 的组合 ID：{label}。", 5000)

    def _clear_sheet_id_check_field(self, sheet_name: str) -> None:
        sheet_keys = self._sheet_key_fields_map()
        if sheet_name not in sheet_keys:
            return
        del sheet_keys[sheet_name]
        self.settings["sheet_key_fields"] = sheet_keys
        self._remember_current_settings()
        self._rebuild_sheet_alignment(sheet_name)
        self.statusBar().showMessage(f"已清空 {sheet_name} 的组合 ID 字段。", 5000)

    def _rebuild_sheet_alignment(self, sheet_name: str) -> None:
        self.alignments.pop(sheet_name, None)
        self._sheet_state_cache.pop(sheet_name, None)
        if self.current_sheet_name != sheet_name:
            self._refresh_sheet_item_state(sheet_name)
            return
        try:
            alignment = self._get_alignment(sheet_name)
        except Exception as exc:
            self._show_friendly_error("重新对齐工作表失败", exc)
            return
        self._set_alignment(alignment)

    def _set_cell_value(self, row: AlignedRow, logical_column: int, new_value: str, reason: str) -> None:
        cell = row.merged_row.cell_at(logical_column)
        if cell is None:
            cell = CellData(column_index=logical_column, value=new_value, data_type="String")
            row.merged_row.cells.append(cell)
        else:
            cell.value = new_value
        row.conflict_columns.discard(logical_column)
        if row.status == "conflict" and not row.conflict_columns:
            row.status = "same"
        if row.reason and not row.reason.startswith("已手工"):
            row.reason = f"{reason}"
        else:
            row.reason = reason
        if self.middle_model is not None:
            self.middle_model.invalidate_diff_cache(row)
        self._refresh_current_views()

    def _delete_merged_rows(self, visible_rows: list[int]) -> None:
        if self.middle_model is None:
            return
        deleted = 0
        for visible_row in visible_rows:
            source_index = self.middle_model.source_row_index(visible_row)
            if source_index >= len(self.middle_model.alignment.rows):
                continue
            self._mark_row_deleted(self.middle_model.alignment.rows[source_index], "中间合并")
            deleted += 1
        if not deleted:
            return
        self._refresh_current_views()
        self.statusBar().showMessage(f"已删除 {deleted} 行合并结果；导出时会跳过这些行。", 5000)

    def _check_merged_ids(self) -> None:
        if self.middle_model is None or self.current_sheet_name is None:
            if self.compare_only_mode.isChecked() and self.current_sheet_name:
                self.compare_only_mode.setChecked(False)
            if self.middle_model is None:
                QMessageBox.information(self, "检测合并ID", "请先开始比对，并关闭纯比对模式以生成预合并结果。")
                return
        assert self.middle_model is not None
        fields = self._sheet_key_fields_map().get(str(self.current_sheet_name), [])
        if not fields:
            QMessageBox.information(self, "检测合并ID", "当前子表还没有标记组合 ID。请右键中间表或左右表的表头，选择“加入组合 ID 字段”。")
            return
        key_columns: list[int] = []
        missing_fields = set(fields)
        for logical_index, binding in enumerate(self.middle_model.alignment.columns, start=1):
            key = normalize_header(binding.key)
            if key in missing_fields:
                key_columns.append(logical_index)
                missing_fields.remove(key)
        if missing_fields:
            QMessageBox.warning(self, "检测合并ID", f"当前合并结果中找不到 ID 字段：{', '.join(sorted(missing_fields))}")
            return

        seen: dict[str, int] = {}
        duplicates: list[dict[str, object]] = []
        empty_count = 0
        checked_count = 0
        for row_index, row in enumerate(self.middle_model.alignment.rows, start=1):
            if row.status == "deleted" or row.merged_row.kind in {"blank", "header"}:
                continue
            values = [row.merged_row.value_at(column).strip() for column in key_columns]
            if not all(values):
                empty_count += 1
                continue
            value = " / ".join(values)
            checked_count += 1
            previous = seen.get(value)
            if previous is not None:
                duplicates.append({"id": value, "first_row": previous, "row": row_index, "field": ", ".join(fields)})
            else:
                seen[value] = row_index
        self._duplicate_id_rows = duplicates
        self._refresh_duplicate_id_list()

        if not duplicates:
            QMessageBox.information(
                self,
                "检测合并ID",
                f"检测完成：组合 ID {', '.join(fields)}\n非空组合键 {checked_count} 个，未发现重复。\n空 ID 行 {empty_count} 行已按规则忽略。",
            )
            return
        lines = [f'{item["id"]}: 第 {item["first_row"]} 行 / 第 {item["row"]} 行' for item in duplicates[:20]]
        suffix = "\n..." if len(duplicates) > 20 else ""
        QMessageBox.warning(
            self,
            "检测合并ID",
            f"组合 ID {', '.join(fields)} 存在 {len(duplicates)} 处重复，已在下方列表显示，可双击定位：\n" + "\n".join(lines) + suffix,
        )

    def _show_friendly_error(self, title: str, exc: BaseException, hint: str = "") -> None:
        category = type(exc).__name__
        message = str(exc).strip() or "(无具体错误信息)"
        dialog = QMessageBox(self)
        dialog.setWindowTitle(title)
        dialog.setIcon(QMessageBox.Warning)
        main_text = f"{title}：{category}"
        dialog.setText(main_text)
        informative_parts = [message]
        if hint:
            informative_parts.append("\n" + hint)
        dialog.setInformativeText("\n".join(informative_parts))
        dialog.setStandardButtons(QMessageBox.Ok)
        import traceback
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        dialog.setDetailedText(detail)
        dialog.exec()

    def _apply_row_note(self) -> None:
        if self.middle_model is None:
            return
        text = self.row_note_input.text().strip()
        visible_rows = self._current_visible_rows()
        if not visible_rows:
            return
        for visible_row in visible_rows:
            source_index = self.middle_model.source_row_index(visible_row)
            if source_index < len(self.middle_model.alignment.rows):
                self.middle_model.alignment.rows[source_index].note = text
        self._update_detail_panel()
        self.statusBar().showMessage("已记录备注。" if text else "已清空备注。", 3000)

    def _apply_ui_font_size(self, point_size: int) -> None:
        app = QApplication.instance()
        if app is None:
            return
        current = app.font()
        if current.pointSize() == point_size:
            return
        current.setPointSize(max(8, min(18, int(point_size))))
        app.setFont(current)

    def _apply_table_spacing(self) -> None:
        row_height = int(self.row_height_combo.currentData() or self.settings.get("table_row_height", 24) or 24)
        header_height = int(self.header_height_combo.currentData() or self.settings.get("table_header_height", 32) or 32)
        row_height = max(18, min(56, row_height))
        header_height = max(22, min(160, header_height))
        for table in (self.left_table, self.middle_table, self.right_table):
            table.verticalHeader().setDefaultSectionSize(row_height)
            header = table.horizontalHeader()
            header.setDefaultSectionSize(header_height)
            header.setMinimumHeight(header_height)
            header.setMaximumHeight(header_height)
        self._resize_token += 1
        self._column_width_cache.clear()

    def _on_table_spacing_changed(self) -> None:
        self._apply_table_spacing()
        self._remember_current_settings()
        self.statusBar().showMessage(
            f"已调整表格行高 {int(self.row_height_combo.currentData() or 24)} / 表头 {int(self.header_height_combo.currentData() or 32)}",
            3000,
        )

    def _on_font_size_changed(self) -> None:
        size = int(self.font_size_combo.currentData() or 9)
        self._apply_ui_font_size(size)
        self._remember_current_settings()
        self.statusBar().showMessage(f"已切换字号为 {size} pt", 3000)

    def _build_legend_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setContentsMargins(10, 10, 10, 10)
        row.setSpacing(8)
        entries = [
            ("仅左有", "#d9edf7", "#0b5394"),
            ("仅右有", "#fce5cd", "#8a4600"),
            ("冲突", "#f4cccc", "#b42318"),
            ("已删除", "#e6e6e6", "#666666"),
            ("差异", "#fff4ce", "#8a4600"),
        ]
        for text, bg, fg in entries:
            chip = QLabel(text)
            chip.setStyleSheet(
                f"background:{bg}; color:{fg}; border:1px solid #d8dee8; "
                "border-radius:4px; padding:1px 8px; font-size:11px;"
            )
            row.addWidget(chip)
        row.addStretch(1)
        return row

    def _install_shortcuts(self) -> None:
        bindings: list[tuple[str, callable]] = [
            ("Ctrl+Return", self._start_compare),
            ("Ctrl+Enter", self._start_compare),
            ("Ctrl+E", self._export),
            ("Ctrl+F", lambda: self.search_input.setFocus()),
            ("F3", lambda: self._jump_to_diff(1)),
            ("Shift+F3", lambda: self._jump_to_diff(-1)),
            ("Ctrl+L", lambda: self._apply_row_choice("left")),
            ("Ctrl+R", lambda: self._apply_row_choice("right")),
        ]
        for seq, handler in bindings:
            shortcut = QShortcut(QKeySequence(seq), self)
            shortcut.setContext(Qt.ApplicationShortcut)
            shortcut.activated.connect(handler)

    def _jump_to_diff(self, direction: int) -> None:
        if direction not in (1, -1):
            return
        base_model = self.middle_model or self.left_model or self.right_model
        if base_model is None or not base_model.visible_rows:
            return
        count = len(base_model.visible_rows)
        current = -1
        for table in self._selection_tables(include_middle=True):
            selection = table.selectionModel()
            if selection is None:
                continue
            rows = [idx.row() for idx in selection.selectedRows()]
            if rows:
                current = rows[0]
                break
        alignment = base_model.alignment
        start = current + direction if current >= 0 else (0 if direction > 0 else count - 1)
        for offset in range(count):
            idx = (start + offset * direction) % count
            source_index = base_model.visible_rows[idx]
            row = alignment.rows[source_index]
            if row.status in {"conflict", "left_only", "right_only", "deleted"} or base_model._compute_row_difference(row):
                self._select_alignment_row(source_index)
                self.statusBar().showMessage(
                    f"跳转到第 {idx + 1} 行 / 共 {count} 行 可见",
                    3000,
                )
                return
        self.statusBar().showMessage("当前筛选下未找到差异行。", 3000)

    def _visible_sheet_row(self, preferred_sheet: str | None = None) -> int | None:
        first_visible: int | None = None
        for row_index in range(self.sheet_list.count()):
            item = self.sheet_list.item(row_index)
            if item is None or item.isHidden():
                continue
            if first_visible is None:
                first_visible = row_index
            if preferred_sheet is not None and item.data(Qt.UserRole) == preferred_sheet:
                return row_index
        return first_visible

    def _filter_sheet_list(self, text: str) -> None:
        keyword = (text or "").strip().lower()
        current_sheet = self.current_sheet_name
        for row_index in range(self.sheet_list.count()):
            item = self.sheet_list.item(row_index)
            if item is None:
                continue
            name = str(item.data(Qt.UserRole) or "").lower()
            item.setHidden(bool(keyword) and keyword not in name)
        target_row = self._visible_sheet_row(current_sheet)
        if target_row is None:
            self.sheet_list.setCurrentRow(-1)
            self.current_sheet_name = None
            self.detail_panel.setHtml("<b>当前筛选下没有子表。</b>")
            return
        current_item = self.sheet_list.currentItem()
        if current_item is None or current_item.isHidden() or current_item.data(Qt.UserRole) != self.sheet_list.item(target_row).data(Qt.UserRole):
            self.sheet_list.setCurrentRow(target_row)

    def _on_sheet_changed(self, current: QListWidgetItem | None, previous: QListWidgetItem | None) -> None:
        del previous
        if current is None:
            self.current_sheet_name = None
            return
        sheet_name = current.data(Qt.UserRole)
        self.current_sheet_name = sheet_name
        try:
            alignment = self._get_alignment(sheet_name)
        except Exception as exc:
            self._show_friendly_error("对齐工作表失败", exc)
            return
        if self.current_sheet_name != sheet_name:
            return
        if not self._confirm_large_difference_if_needed(alignment):
            self.detail_panel.setHtml(
                "<b>已暂停渲染当前子表。</b><br>"
                "这张表的差异量非常高，可能是左右文件选错或 ID 字段不合适。"
            )
            self.statusBar().showMessage("已取消渲染高差异子表；可调整文件或组合 ID 后重新开始比对。", 6000)
            return
        self._set_alignment(alignment)

    def _confirm_large_difference_if_needed(self, alignment: SheetAlignment) -> bool:
        sheet_name = str(alignment.sheet_name)
        if sheet_name in self._large_diff_confirmed_sheets:
            return True
        total_rows = max(len(alignment.rows), 1)
        if total_rows < 50:
            return True
        changed_rows = sum(1 for row in alignment.rows if row.status != "same")
        conflict_rows = sum(1 for row in alignment.rows if row.status == "conflict")
        only_side_rows = sum(1 for row in alignment.rows if row.status in {"left_only", "right_only", "deleted"})
        ratio = changed_rows / total_rows
        if changed_rows < 30 or ratio < 0.65:
            return True
        box = QMessageBox(self)
        box.setWindowTitle("差异量过高")
        box.setIcon(QMessageBox.Warning)
        box.setText(f"当前子表“{sheet_name}”差异量非常高，可能选错了左右文件。")
        box.setInformativeText(
            f"总行数：{total_rows}\n"
            f"差异行：{changed_rows}（{ratio:.0%}）\n"
            f"冲突行：{conflict_rows}\n"
            f"仅单侧存在/删除行：{only_side_rows}\n\n"
            "如果继续，界面仍会渲染该表，遇到大表可能明显卡顿。"
        )
        continue_button = box.addButton("继续渲染", QMessageBox.AcceptRole)
        cancel_button = box.addButton("先不渲染", QMessageBox.RejectRole)
        box.setDefaultButton(cancel_button)
        box.exec()
        if box.clickedButton() is continue_button:
            self._large_diff_confirmed_sheets.add(sheet_name)
            return True
        return False

    def _get_alignment(self, sheet_name: str) -> SheetAlignment:
        if sheet_name in self.alignments:
            return self.alignments[sheet_name]
        left_sheet = self.left_workbook.get_sheet(sheet_name) if self.left_workbook else None
        right_sheet = self.right_workbook.get_sheet(sheet_name) if self.right_workbook else None
        if self.base_workbook is not None:
            base_sheet = self.base_workbook.get_sheet(sheet_name)
            alignment = align_sheets_three_way(
                base_sheet,
                left_sheet,
                right_sheet,
                self._active_merge_rule(),
                self._comparison_options(),
                self._key_fields_for_sheet(sheet_name),
                strict_single_key=self._strict_single_key_enabled_for_sheet(sheet_name),
            )
        else:
            alignment = align_sheets(
                left_sheet,
                right_sheet,
                self._active_merge_rule(),
                self._comparison_options(),
                self._key_fields_for_sheet(sheet_name),
                strict_single_key=self._strict_single_key_enabled_for_sheet(sheet_name),
            )
        self.alignments[sheet_name] = alignment
        return alignment

    def _available_sheet_names(self) -> list[str]:
        left_names = set(self.left_workbook.sheet_names) if self.left_workbook else set()
        right_names = set(self.right_workbook.sheet_names) if self.right_workbook else set()
        return sorted(left_names | right_names)

    def _ensure_sheet_states(self, sheet_names: list[str] | None = None) -> None:
        for sheet_name in sheet_names or self._available_sheet_names():
            self._sheet_state(sheet_name)

    def _cancel_sheet_state_warmup(self) -> None:
        self._sheet_state_warmup_token += 1
        self._sheet_state_warmup_queue = []
        self._sheet_state_warmup_index = 0
        self._warmup_started_at = 0.0

    def _schedule_sheet_state_warmup(self, sheet_names: list[str] | None = None) -> None:
        if self.left_workbook is None or self.right_workbook is None:
            return
        if self.changed_sheets_only.isChecked():
            return
        pending = [name for name in (sheet_names or self._available_sheet_names()) if name not in self._sheet_state_cache]
        if not pending:
            self._update_sheet_summary(self.alignments.get(self.current_sheet_name) if self.current_sheet_name else None)
            return
        self._sheet_state_warmup_token += 1
        self._sheet_state_warmup_queue = pending
        self._sheet_state_warmup_index = 0
        self._warmup_started_at = perf_counter()
        token = self._sheet_state_warmup_token
        QTimer.singleShot(60, lambda: self._warmup_sheet_states(token))

    def _warmup_sheet_states(self, token: int) -> None:
        if token != self._sheet_state_warmup_token:
            return
        batch_size = 1
        end_index = min(self._sheet_state_warmup_index + batch_size, len(self._sheet_state_warmup_queue))
        for index in range(self._sheet_state_warmup_index, end_index):
            sheet_name = self._sheet_state_warmup_queue[index]
            self._sheet_state(sheet_name)
            self._refresh_sheet_item_state(sheet_name)
        self._sheet_state_warmup_index = end_index
        self._update_sheet_summary(self.alignments.get(self.current_sheet_name) if self.current_sheet_name else None)
        if self._sheet_state_warmup_index < len(self._sheet_state_warmup_queue):
            QTimer.singleShot(16, lambda: self._warmup_sheet_states(token))
            return
        if self._warmup_started_at > 0:
            warmup_ms = (perf_counter() - self._warmup_started_at) * 1000.0
            self._last_compare_metrics["warmup_ms"] = warmup_ms
            self._warmup_started_at = 0.0
            self.statusBar().showMessage(
                "后台分析完成：共 {sheet_count} 张工作表 | 后台耗时 {warmup_ms} | 首屏 {first_ready_ms}".format(
                    sheet_count=len(self._available_sheet_names()),
                    warmup_ms=self._format_duration_ms(warmup_ms),
                    first_ready_ms=self._format_duration_ms(self._last_compare_metrics.get("first_ready_ms", 0.0)),
                ),
                6000,
            )

    def _apply_sheet_item_state(self, item: QListWidgetItem, sheet_name: str, state: dict[str, object] | None) -> None:
        item.setSizeHint(QSize(0, 29))
        font = item.font()
        font.setBold(False)
        item.setFont(font)
        item.setBackground(QBrush(QColor("#ffffff")))
        item.setData(Qt.UserRole + 1, "")
        item.setData(Qt.UserRole + 2, "")
        if state is None:
            item.setText(f"○  {sheet_name}")
            item.setToolTip(f"{sheet_name}\n状态: 未加载，切换到该子表后计算变更。")
            item.setForeground(QBrush(QColor("#667085")))
            return
        state_key = str(state.get("state_key", "same"))
        tag = str(state.get("tag", "")).strip("[]")
        marker_by_state = {
            "error": "!",
            "deleted": "-",
            "added": "+",
            "conflict": "!",
            "changed": "*",
            "same": " ",
        }
        marker = marker_by_state.get(state_key, "*")
        label = f"{marker}  {tag}  {sheet_name}" if tag else f"   {sheet_name}"
        item.setText(label)
        item.setToolTip(str(state["tooltip"]))
        item.setData(Qt.UserRole + 1, state_key)
        item.setData(Qt.UserRole + 2, str(state.get("summary", "")))
        if bool(state["changed"]):
            item.setForeground(QBrush(QColor(str(state["foreground"]))))
            font.setBold(True)
            item.setFont(font)
            return
        item.setForeground(QBrush(QColor("#667085")))

    def _refresh_sheet_item_state(self, sheet_name: str) -> None:
        state = self._sheet_state(sheet_name)
        for row_index in range(self.sheet_list.count()):
            item = self.sheet_list.item(row_index)
            if item.data(Qt.UserRole) == sheet_name:
                self._apply_sheet_item_state(item, sheet_name, state)
                break

    @staticmethod
    def _column_delta_counts(alignment: SheetAlignment) -> tuple[int, int]:
        right_added = sum(1 for column in alignment.columns if column.left_index is None and column.right_index is not None)
        left_only = sum(1 for column in alignment.columns if column.right_index is None and column.left_index is not None)
        return right_added, left_only

    def _sheet_state(self, sheet_name: str) -> dict[str, object]:
        cached = self._sheet_state_cache.get(sheet_name)
        if cached is not None:
            return cached
        try:
            alignment = self._get_alignment(sheet_name)
        except Exception as exc:
            state = {
                "changed": True,
                "has_conflict": True,
                "state_key": "error",
                "summary": f"对齐失败：{exc}",
                "tag": "[错误]",
                "background": "#fff1f3",
                "foreground": "#b42318",
                "tooltip": f"{sheet_name}\n状态: 对齐失败\n{exc}",
            }
            self._sheet_state_cache[sheet_name] = state
            return state
        changed_rows = sum(1 for row in alignment.rows if row.status != "same")
        right_added_columns, left_only_columns = self._column_delta_counts(alignment)
        only_left_sheet = alignment.left_sheet is not None and alignment.right_sheet is None
        only_right_sheet = alignment.right_sheet is not None and alignment.left_sheet is None
        has_conflict = alignment.conflict_count > 0
        changed = changed_rows > 0 or only_left_sheet or only_right_sheet or right_added_columns > 0 or left_only_columns > 0
        if only_left_sheet:
            state_key = "deleted"
            summary = "左侧独有，右侧已删除"
            tag = "[删除]"
            background = "#fde7e9"
            foreground = "#b42318"
        elif only_right_sheet:
            state_key = "added"
            summary = "右侧新增子表"
            tag = "[新增]"
            background = "#e8f5e9"
            foreground = "#027a48"
        elif has_conflict:
            state_key = "conflict"
            summary = self._compose_sheet_summary(f"{alignment.conflict_count} 行冲突", right_added_columns, left_only_columns)
            tag = "[冲突]"
            background = "#fff1f3"
            foreground = "#b42318"
        elif changed:
            state_key = "changed"
            summary = self._compose_sheet_summary(f"{changed_rows} 行变更", right_added_columns, left_only_columns)
            tag = "[变更]"
            background = "#fff4ce"
            foreground = "#8a4600"
        else:
            state_key = "same"
            summary = "无变更"
            tag = ""
            background = "#ffffff"
            foreground = "#667085"
        tooltip = (
            f"{sheet_name}\n"
            f"状态: {summary}\n"
            f"变更行: {changed_rows}\n"
            f"右侧新增列: {right_added_columns}\n"
            f"左侧独有列: {left_only_columns}\n"
            f"冲突行: {alignment.conflict_count}\n"
            f"未处理: {alignment.unresolved_count}"
        )
        state = {
            "changed": changed,
            "has_conflict": has_conflict,
            "state_key": state_key,
            "summary": summary,
            "tag": tag,
            "background": background,
            "foreground": foreground,
            "right_added_columns": right_added_columns,
            "left_only_columns": left_only_columns,
            "tooltip": tooltip,
        }
        self._sheet_state_cache[sheet_name] = state
        return state

    @staticmethod
    def _compose_sheet_summary(base: str, right_added_columns: int, left_only_columns: int) -> str:
        extras: list[str] = []
        if right_added_columns:
            extras.append(f"右增{right_added_columns}列")
        if left_only_columns:
            extras.append(f"左独有{left_only_columns}列")
        if not extras:
            return base
        return f"{base}，{'，'.join(extras)}"

    def _update_sheet_summary(self, current_alignment: SheetAlignment | None = None) -> None:
        total = len(self._available_sheet_names())
        inspected = len(self._sheet_state_cache)
        changed = sum(1 for state in self._sheet_state_cache.values() if bool(state["changed"]))
        if inspected < total:
            base = f"共 {total} 张工作表 | 已分析 {inspected} 张 | 已确认有变更 {changed} 张"
        else:
            base = f"共 {total} 张工作表 | 有变更 {changed} 张"
        if current_alignment is None:
            self.sheet_summary_label.setText(base)
            return
        right_added_columns, left_only_columns = self._column_delta_counts(current_alignment)
        column_text = ""
        if right_added_columns or left_only_columns:
            parts = []
            if right_added_columns:
                parts.append(f"右增 {right_added_columns} 列")
            if left_only_columns:
                parts.append(f"左独有 {left_only_columns} 列")
            column_text = " | " + " | ".join(parts)
        self.sheet_summary_label.setText(
            f"{base} | 当前表冲突 {current_alignment.conflict_count} 行 | 未处理 {current_alignment.unresolved_count} 行{column_text}"
        )

    def _set_alignment(self, alignment: SheetAlignment) -> None:
        if self.current_sheet_name is not None and alignment.sheet_name != self.current_sheet_name:
            return
        current_item = self.sheet_list.currentItem()
        if current_item is not None and current_item.data(Qt.UserRole) != alignment.sheet_name:
            return
        comparison_options = self._comparison_options()
        if self.left_model is None or self.right_model is None:
            self.left_model = MergeTableModel(alignment, "left", comparison_options, self)
            self.right_model = MergeTableModel(alignment, "right", comparison_options, self)
            self.left_table.setModel(self.left_model)
            self.right_table.setModel(self.right_model)
            self._connect_selection_models()
        else:
            self.left_model.set_alignment(alignment, comparison_options)
            self.right_model.set_alignment(alignment, comparison_options)

        if self.compare_only_mode.isChecked():
            if self.middle_model is not None:
                self.middle_model.set_alignment(alignment, comparison_options)
        else:
            self._ensure_middle_model(alignment)

        self._apply_filters()
        self._schedule_resize_columns()
        self._update_detail_panel()
        self._update_sheet_summary(alignment)
        self._update_sheet_overview(alignment)
        self._refresh_sheet_item_state(alignment.sheet_name)
        right_added_columns, left_only_columns = self._column_delta_counts(alignment)
        column_parts = []
        if right_added_columns:
            column_parts.append(f"右增 {right_added_columns} 列")
        if left_only_columns:
            column_parts.append(f"左独有 {left_only_columns} 列")
        column_suffix = f"，{ '，'.join(column_parts) }" if column_parts else ""
        sheet_keys = self._sheet_key_fields_map().get(alignment.sheet_name, [])
        key_label = sheet_keys if sheet_keys else (alignment.key_fields or ["自动内容对齐"])
        key_prefix = "组合ID" if sheet_keys else "键"
        self.statusBar().showMessage(
            f"表 {alignment.sheet_name}: 共 {len(alignment.rows)} 行，冲突 {alignment.conflict_count} 行{column_suffix}，{key_prefix} {key_label}，规则 {self._active_merge_rule().title}",
            8000,
        )

    def _update_sheet_overview(self, alignment: SheetAlignment) -> None:
        added = sum(1 for row in alignment.rows if row.status == "right_only")
        deleted = sum(1 for row in alignment.rows if row.status in {"left_only", "deleted"})
        conflicts = sum(1 for row in alignment.rows if row.status == "conflict")
        changed_columns = self._changed_column_count(alignment)
        duplicates = self._find_merged_id_duplicates(alignment)
        self._duplicate_id_rows = duplicates
        self.overview_added_button.setText(f"新增 {added}")
        self.overview_deleted_button.setText(f"删除 {deleted}")
        self.overview_conflict_button.setText(f"冲突 {conflicts}")
        self.overview_changed_columns_button.setText(f"改动列 {changed_columns}")
        self.overview_duplicate_id_button.setText(f"重复ID {len(duplicates)}")
        self.overview_added_button.setEnabled(added > 0)
        self.overview_deleted_button.setEnabled(deleted > 0)
        self.overview_conflict_button.setEnabled(conflicts > 0)
        self.overview_changed_columns_button.setEnabled(changed_columns > 0)
        self.overview_duplicate_id_button.setEnabled(bool(duplicates))
        self._refresh_duplicate_id_list()

    def _changed_column_count(self, alignment: SheetAlignment) -> int:
        changed: set[int] = set()
        for logical_index, column in enumerate(alignment.columns, start=1):
            if column.left_index is None or column.right_index is None:
                changed.add(logical_index)
                continue
            for row in alignment.rows:
                if row.status == "deleted":
                    continue
                left_value = row.left_row.value_at(column.left_index) if row.left_row and column.left_index else ""
                right_value = row.right_row.value_at(column.right_index) if row.right_row and column.right_index else ""
                if not compare_text_values(left_value, right_value, self._comparison_options()):
                    changed.add(logical_index)
                    break
        return len(changed)

    def _find_merged_id_duplicates(self, alignment: SheetAlignment) -> list[dict[str, object]]:
        fields = self._sheet_key_fields_map().get(alignment.sheet_name, [])
        if not fields:
            return []
        key_columns: list[int] = []
        missing_fields = set(fields)
        for logical_index, binding in enumerate(alignment.columns, start=1):
            key = normalize_header(binding.key)
            if key in missing_fields:
                key_columns.append(logical_index)
                missing_fields.remove(key)
        if missing_fields:
            return []
        seen: dict[str, int] = {}
        duplicates: list[dict[str, object]] = []
        for row_index, row in enumerate(alignment.rows, start=1):
            if row.status == "deleted" or row.merged_row.kind in {"blank", "header"}:
                continue
            values = [row.merged_row.value_at(column).strip() for column in key_columns]
            if not all(values):
                continue
            value = " / ".join(values)
            previous = seen.get(value)
            if previous is not None:
                duplicates.append({"id": value, "first_row": previous, "row": row_index, "field": ", ".join(fields)})
            else:
                seen[value] = row_index
        return duplicates

    def _refresh_duplicate_id_list(self) -> None:
        self.duplicate_id_list.clear()
        if not self._duplicate_id_rows:
            self.duplicate_id_label.setVisible(False)
            self.duplicate_id_list.setVisible(False)
            return
        for item in self._duplicate_id_rows:
            text = f'{item["id"]}: 第 {item["first_row"]} 行 / 第 {item["row"]} 行'
            list_item = QListWidgetItem(text)
            list_item.setData(Qt.UserRole, int(item["row"]))
            self.duplicate_id_list.addItem(list_item)
        self.duplicate_id_label.setVisible(True)
        self.duplicate_id_list.setVisible(True)

    def _jump_to_duplicate_item(self, item: QListWidgetItem) -> None:
        row_number = int(item.data(Qt.UserRole) or 0)
        if row_number > 0:
            self._select_alignment_row(row_number - 1)

    def _show_duplicate_ids(self) -> None:
        if not self._duplicate_id_rows:
            QMessageBox.information(self, "重复ID", "当前子表没有重复 ID，或尚未右键标记检测 ID 列。")
            return
        self.duplicate_id_label.setVisible(True)
        self.duplicate_id_list.setVisible(True)
        self.duplicate_id_list.setFocus()
        self.duplicate_id_list.setCurrentRow(0)
        first = self.duplicate_id_list.item(0)
        if first is not None:
            self._jump_to_duplicate_item(first)

    def _jump_to_next_duplicate_id(self) -> None:
        if not self._duplicate_id_rows:
            QMessageBox.information(self, "重复ID", "当前子表没有重复 ID，或尚未右键标记检测 ID 列。")
            return
        self.duplicate_id_label.setVisible(True)
        self.duplicate_id_list.setVisible(True)
        current_row = self.duplicate_id_list.currentRow()
        next_row = (current_row + 1) % self.duplicate_id_list.count() if current_row >= 0 else 0
        self.duplicate_id_list.setCurrentRow(next_row)
        item = self.duplicate_id_list.item(next_row)
        if item is not None:
            self._jump_to_duplicate_item(item)
            self.statusBar().showMessage(
                f"跳转到重复ID {next_row + 1} / {self.duplicate_id_list.count()}",
                3000,
            )

    def _show_changed_columns_summary(self) -> None:
        if self.middle_model is None and self.left_model is None:
            return
        model = self.middle_model or self.left_model
        alignment = model.alignment
        lines: list[str] = []
        for logical_index, column in enumerate(alignment.columns, start=1):
            if column.left_index is None:
                lines.append(f"{column.title}: 右侧新增列")
            elif column.right_index is None:
                lines.append(f"{column.title}: 左侧独有列")
            else:
                for row in alignment.rows:
                    left_value = row.left_row.value_at(column.left_index) if row.left_row and column.left_index else ""
                    right_value = row.right_row.value_at(column.right_index) if row.right_row and column.right_index else ""
                    if not compare_text_values(left_value, right_value, self._comparison_options()):
                        lines.append(f"{column.title}: 存在内容差异")
                        break
        QMessageBox.information(self, "改动列", "\n".join(lines[:80]) if lines else "当前子表没有改动列。")

    def _jump_to_changed_row(self) -> None:
        model = self.middle_model or self.left_model or self.right_model
        if model is None:
            return

        def matcher(row: AlignedRow) -> bool:
            if row.status in {"conflict", "left_only", "right_only", "deleted"}:
                return False
            return model._compute_row_difference(row)

        self._jump_to_matching_alignment_row("改动", matcher)

    def _jump_to_row_status(self, statuses: set[str], label: str = "对应") -> None:
        self._jump_to_matching_alignment_row(label, lambda row: row.status in statuses)

    def _current_selected_source_index(self, model: MergeTableModel) -> int | None:
        for table in self._selection_tables(include_middle=True):
            selection = table.selectionModel()
            if selection is None:
                continue
            rows = [idx.row() for idx in selection.selectedRows()]
            if not rows:
                continue
            visible_row = rows[0]
            if 0 <= visible_row < len(model.visible_rows):
                return model.visible_rows[visible_row]
        return None

    def _jump_to_matching_alignment_row(self, label: str, matcher) -> None:
        model = self.middle_model or self.left_model or self.right_model
        if model is None:
            return
        matches = [
            source_index
            for source_index, row in enumerate(model.alignment.rows)
            if matcher(row)
        ]
        if not matches:
            self.statusBar().showMessage(f"当前子表没有{label}行。", 3000)
            return
        current_source_index = self._current_selected_source_index(model)
        target = matches[0]
        if current_source_index is not None:
            target = next((source_index for source_index in matches if source_index > current_source_index), matches[0])
        self._select_alignment_row(target)
        position = matches.index(target) + 1
        self.statusBar().showMessage(
            f"跳转到{label} {position} / {len(matches)}，合并行 {target + 1}",
            3000,
        )

    def _select_alignment_row(self, source_index: int) -> None:
        model = self.middle_model or self.left_model or self.right_model
        if model is None:
            return
        try:
            visible_row = model.visible_rows.index(source_index)
        except ValueError:
            self.view_mode_combo.setCurrentIndex(max(0, self.view_mode_combo.findData("all")))
            self._apply_filters()
            try:
                visible_row = model.visible_rows.index(source_index)
            except ValueError:
                return
        self._select_rows([visible_row], source=None)
        self._update_detail_panel(visible_row)
        for table in self._selection_tables(include_middle=True):
            table_model = table.model()
            if table_model is not None:
                table.scrollTo(table_model.index(visible_row, 0), QTableView.PositionAtCenter)

    def _ensure_middle_model(self, alignment: SheetAlignment) -> None:
        comparison_options = self._comparison_options()
        if self.middle_model is None:
            self.middle_model = MergeTableModel(alignment, "middle", comparison_options, self)
            self.middle_table.setModel(self.middle_model)
            self._connect_selection_models()
            return
        self.middle_model.set_alignment(alignment, comparison_options)

    def _apply_filters(self) -> None:
        if not self.left_model or not self.right_model:
            return
        view_mode = str(self.view_mode_combo.currentData() or "all")
        base_model = self.middle_model or self.left_model
        visible_rows = base_model.compute_visible_rows(view_mode, "")
        self.left_model.set_visible_rows(visible_rows)
        if self.middle_model is not None:
            self.middle_model.set_visible_rows(visible_rows)
        self.right_model.set_visible_rows(visible_rows)
        search_text = self.search_input.text().strip()
        if search_text:
            restart = search_text.lower() != self._last_search_text
            QTimer.singleShot(0, lambda restart=restart: self._jump_to_search(1, restart=restart))
        else:
            self._last_search_text = ""
            self._last_search_match = None

    def _jump_to_search(self, direction: int = 1, *, restart: bool = False) -> None:
        if direction not in (1, -1):
            return
        keyword = self.search_input.text().strip().lower()
        if not keyword:
            self._last_search_text = ""
            self._last_search_match = None
            return
        base_model = self.middle_model or self.left_model or self.right_model
        if base_model is None or not base_model.visible_rows:
            return
        count = len(base_model.visible_rows)
        current_visible = -1
        current_column = -1
        if not restart:
            current_visible, current_column = self._current_search_anchor()
        start = current_visible + direction if current_visible >= 0 else (0 if direction > 0 else count - 1)
        for offset in range(count):
            visible_row = (start + offset * direction) % count
            source_index = base_model.visible_rows[visible_row]
            row = base_model.alignment.rows[source_index]
            column_index = self._first_matching_column(row, keyword, start_after=current_column if visible_row == current_visible else -1)
            if column_index is None:
                continue
            self._last_search_text = keyword
            self._last_search_match = (source_index, column_index)
            self._select_rows([visible_row], source=None)
            self._update_detail_panel(visible_row)
            self._scroll_tables_to_cell(visible_row, column_index)
            self.statusBar().showMessage(f"搜索命中：第 {visible_row + 1} 行，字段 {column_index}", 3000)
            return
        self.statusBar().showMessage("当前视图未找到匹配内容。", 3000)

    def _current_search_anchor(self) -> tuple[int, int]:
        if self._last_search_match:
            source_index, column_index = self._last_search_match
            model = self.middle_model or self.left_model or self.right_model
            if model is not None:
                try:
                    return model.visible_rows.index(source_index), column_index
                except ValueError:
                    pass
        for table in self._selection_tables(include_middle=True):
            selection = table.selectionModel()
            if selection is None:
                continue
            indexes = selection.selectedIndexes()
            if indexes:
                first = indexes[0]
                return first.row(), first.column() + 1
            rows = selection.selectedRows()
            if rows:
                return rows[0].row(), -1
        return -1, -1

    def _first_matching_column(self, row: AlignedRow, keyword: str, *, start_after: int = -1) -> int | None:
        columns = self.left_model.alignment.columns if self.left_model is not None else []
        if not columns:
            return None
        ordered_columns = list(range(1, len(columns) + 1))
        if start_after > 0:
            ordered_columns = [column for column in ordered_columns if column > start_after]
        for logical_index in ordered_columns:
            binding = columns[logical_index - 1]
            values = [
                row.left_row.value_at(binding.left_index or -1) if row.left_row and binding.left_index else "",
                row.merged_row.value_at(logical_index) if row.merged_row else "",
                row.right_row.value_at(binding.right_index or -1) if row.right_row and binding.right_index else "",
            ]
            if any(keyword in str(value or "").lower() for value in values):
                return logical_index
        return None

    def _scroll_tables_to_cell(self, visible_row: int, logical_column: int) -> None:
        for table in self._selection_tables(include_middle=True):
            table_model = table.model()
            if table_model is None:
                continue
            column = max(0, min(logical_column - 1, table_model.columnCount() - 1))
            if 0 <= visible_row < table_model.rowCount():
                table.scrollTo(table_model.index(visible_row, column), QTableView.PositionAtCenter)

    def _sync_selection(self, visible_row: int) -> None:
        self._select_rows([visible_row], source=self.sender() if isinstance(self.sender(), QTableView) else None)
        self._update_detail_panel(visible_row)

    def _current_visible_rows(self) -> list[int]:
        for table in self._selection_tables():
            selection = table.selectionModel()
            if selection is None:
                continue
            rows = sorted(index.row() for index in selection.selectedRows())
            if rows:
                return rows
        return []

    def _apply_row_choice(self, mode: str) -> None:
        if self.middle_model is None or self.current_sheet_name is None:
            return
        visible_rows = self._current_visible_rows()
        if not visible_rows:
            QMessageBox.information(self, "提示", "请先选择至少一行。")
            return
        if self.compare_only_mode.isChecked():
            QMessageBox.information(self, "提示", "纯比对模式下不显示预合并表格，也不能执行取左/取右。")
            return

        if mode == "auto":
            rebuilt = align_sheets(
                self.middle_model.alignment.left_sheet,
                self.middle_model.alignment.right_sheet,
                self._active_merge_rule(),
                self._comparison_options(),
                self._key_fields_for_sheet(self.current_sheet_name),
                strict_single_key=self._strict_single_key_enabled_for_sheet(self.current_sheet_name),
            )
            current = self.middle_model.alignment
            for visible_row in visible_rows:
                source_index = self.middle_model.source_row_index(visible_row)
                if source_index < len(current.rows) and source_index < len(rebuilt.rows):
                    current.rows[source_index] = rebuilt.rows[source_index]
            self._refresh_current_views()
            self._select_rows(visible_rows, source=self.middle_table)
            return

        skipped_rows = 0
        deleted_rows = 0
        for visible_row in visible_rows:
            source_row_index = self.middle_model.source_row_index(visible_row)
            row = self.middle_model.alignment.rows[source_row_index]
            reference = row.left_row if mode == "left" else row.right_row
            if reference is None:
                self._mark_row_deleted(row, mode)
                deleted_rows += 1
                continue
            values_by_column: dict[int, str] = {}
            for logical_index, binding in enumerate(self.middle_model.alignment.columns, start=1):
                source_index = binding.left_index if mode == "left" else binding.right_index
                values_by_column[logical_index] = reference.value_at(source_index or -1) if source_index else ""
            row.merged_row = clone_row_with_values(reference, values_by_column)
            row.status = "same"
            row.reason = f"已采用{mode}侧行"
            row.conflict_columns.clear()
        self._refresh_current_views()
        self._select_rows(visible_rows, source=self.middle_table)
        if deleted_rows or skipped_rows:
            parts = []
            if deleted_rows:
                parts.append(f"{deleted_rows} 行因{mode}侧不存在而标记为删除")
            if skipped_rows:
                parts.append(f"{skipped_rows} 行已跳过")
            QMessageBox.information(self, "批量处理完成", "；".join(parts))

    def _refresh_current_views(self) -> None:
        for table in (self.left_table, self.middle_table, self.right_table):
            model = table.model()
            if model is not None:
                if isinstance(model, MergeTableModel):
                    model.invalidate_diff_cache()
                model.layoutChanged.emit()
        if self.current_sheet_name:
            alignment = self.alignments[self.current_sheet_name]
            self._sheet_state_cache.pop(self.current_sheet_name, None)
            self._refresh_sheet_list(self.current_sheet_name)
            self.statusBar().showMessage(
                f"表 {alignment.sheet_name}: 冲突 {alignment.conflict_count} 行，未自动处理 {alignment.unresolved_count} 行",
                5000,
            )
        self._update_detail_panel()

    def _resize_columns(self) -> None:
        for table, side in ((self.left_table, "left"), (self.middle_table, "middle"), (self.right_table, "right")):
            model = table.model()
            if model is None:
                continue
            row_count = model.rowCount()
            col_count = model.columnCount()
            if col_count <= 0:
                continue
            alignment = getattr(model, "alignment", None)
            sheet_name = getattr(alignment, "sheet_name", "") or ""
            cache_key = (str(sheet_name), side, col_count, row_count)
            cached_widths = self._column_width_cache.get(cache_key)
            if cached_widths is not None:
                table.setUpdatesEnabled(False)
                try:
                    for column, width in enumerate(cached_widths[:col_count]):
                        table.setColumnWidth(column, width)
                finally:
                    table.setUpdatesEnabled(True)
                continue
            header = table.horizontalHeader()
            precision = 12 if row_count > 500 else 20 if row_count > 200 else 40
            auto_fit_columns = min(col_count, 10 if row_count > 300 else 14 if row_count > 120 else col_count)
            widths: list[int] = []
            table.setUpdatesEnabled(False)
            try:
                if header is not None:
                    header.setResizeContentsPrecision(precision)
                for column in range(auto_fit_columns):
                    table.resizeColumnToContents(column)
                for column in range(col_count):
                    width = table.columnWidth(column)
                    if column >= auto_fit_columns:
                        width = 110 if column < 6 else 96
                    elif column < 6:
                        width = min(max(width, 90), 260)
                    else:
                        width = min(max(width, 80), 260)
                    table.setColumnWidth(column, width)
                    widths.append(width)
            finally:
                table.setUpdatesEnabled(True)
            self._column_width_cache[cache_key] = widths

    def _schedule_resize_columns(self, delay_ms: int = 0) -> None:
        self._resize_token += 1
        token = self._resize_token
        QTimer.singleShot(delay_ms, lambda: self._run_scheduled_resize(token))

    def _run_scheduled_resize(self, token: int) -> None:
        if token != self._resize_token:
            return
        self._resize_columns()

    @staticmethod
    def _format_duration_ms(value: float) -> str:
        if value <= 0:
            return "0 ms"
        if value < 1000:
            return f"{value:.0f} ms"
        return f"{value / 1000.0:.2f} s"

    def _export(self) -> None:
        if not self._ensure_export_alignments():
            return

        template_workbook = self.left_workbook if self.template_source_combo.currentData() == "left" else self.right_workbook
        compare_workbook = self.right_workbook if template_workbook is self.left_workbook else self.left_workbook
        if not self._confirm_export_precheck(template_workbook):
            return

        output_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出结果",
            str(Path.cwd() / f"{self._suggest_output_stem(self.left_workbook)}.xml"),
            "Excel XML (*.xml)",
        )
        if not output_path:
            return

        template_format = template_workbook.source_meta.get("file_format", "xml")
        if template_workbook.source_kind == SOURCE_GOOGLE_SHEETS:
            QMessageBox.warning(
                self,
                "无法导出",
                "当前版本不能把 Google Sheets 作为导出模板来源。\n请把模板来源切到本地 XML 或 SVN XML。",
            )
            return
        if template_format != "xml" and template_workbook.source_kind == SOURCE_LOCAL:
            QMessageBox.warning(
                self,
                "无法导出",
                "当前版本不能把 xlsx/csv 作为导出模板来源。\n请把模板来源切到本地 XML 或 SVN XML。",
            )
            return
        output_path = self._ensure_extension(output_path, "xml")

        export_workbook(
            template_workbook,
            compare_workbook,
            self.alignments,
            output_path,
            self._active_merge_rule(),
        )
        QMessageBox.information(self, "导出完成", f"已导出到:\n{output_path}")

    def _confirm_export_precheck(self, template_workbook: WorkbookData) -> bool:
        template_format = template_workbook.source_meta.get("file_format", "xml")
        blockers: list[str] = []
        warnings: list[str] = []
        if template_workbook.source_kind == SOURCE_GOOGLE_SHEETS:
            blockers.append("模板来源是 Google Sheets，当前不能作为 XML 导出模板。")
        if template_format != "xml" and template_workbook.source_kind == SOURCE_LOCAL:
            blockers.append(f"模板来源是 {template_format}，当前只能用 XML 作为导出模板。")
        for sheet_name in sorted(self.alignments):
            alignment = self.alignments[sheet_name]
            if alignment.unresolved_count:
                warnings.append(f"{sheet_name}: 未处理 {alignment.unresolved_count} 行")
            if alignment.conflict_count:
                warnings.append(f"{sheet_name}: 冲突 {alignment.conflict_count} 行")
            right_added_columns, left_only_columns = self._column_delta_counts(alignment)
            if right_added_columns:
                warnings.append(f"{sheet_name}: 右侧新增 {right_added_columns} 列")
            if left_only_columns:
                warnings.append(f"{sheet_name}: 左侧独有 {left_only_columns} 列")
            duplicates = self._find_merged_id_duplicates(alignment)
            if duplicates:
                warnings.append(f"{sheet_name}: 重复 ID {len(duplicates)} 处")
        if blockers:
            QMessageBox.warning(self, "导出前检查未通过", "\n".join(blockers))
            return False
        if not warnings:
            return True
        message = "导出前检查发现以下风险：\n\n" + "\n".join(warnings[:80])
        if len(warnings) > 80:
            message += f"\n... 还有 {len(warnings) - 80} 项"
        message += "\n\n是否继续导出？"
        box = QMessageBox(self)
        box.setWindowTitle("导出前检查")
        box.setIcon(QMessageBox.Warning)
        box.setText("导出前检查发现风险")
        box.setInformativeText(message)
        continue_button = box.addButton("继续导出", QMessageBox.AcceptRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(continue_button)
        box.exec()
        return box.clickedButton() == continue_button

    def _ensure_export_alignments(self) -> bool:
        if self.left_workbook is None or self.right_workbook is None:
            QMessageBox.warning(self, "无法导出", "请先加载左右两侧内容。")
            return False
        sheet_names = sorted(set(self.left_workbook.sheet_names) | set(self.right_workbook.sheet_names))
        for sheet_name in sheet_names:
            self.alignments[sheet_name] = self._get_alignment(sheet_name)
        if not self.alignments:
            QMessageBox.warning(self, "无法导出", "当前还没有可导出的表。")
            return False
        return True

    def _export_diff_report(self) -> None:
        if not self._ensure_export_alignments():
            return

        output_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出差异报告",
            str(Path.cwd() / f"{self._suggest_output_stem(self.left_workbook)}_diff_report.html"),
            "HTML 报告 (*.html);;Markdown 报告 (*.md);;CSV 明细 (*.csv);;文本报告 (*.txt)",
        )
        if not output_path:
            return
        export_mode = self._detect_export_mode(output_path, selected_filter)
        if export_mode == "xml":
            export_mode = "html"
        output_path = self._ensure_extension(output_path, export_mode)
        export_diff_report(self.alignments, output_path, self._comparison_options())
        QMessageBox.information(self, "导出完成", f"已导出差异报告到:\n{output_path}")

    @staticmethod
    def _suggest_output_stem(workbook: WorkbookData) -> str:
        if workbook.path is not None:
            return workbook.path.stem
        return (workbook.source_label or "merged").replace("@", "_").replace(" ", "_")

    @staticmethod
    def _detect_export_mode(output_path: str, selected_filter: str) -> str:
        suffix = Path(output_path).suffix.lower()
        if suffix == ".html":
            return "html"
        if suffix == ".md":
            return "md"
        if suffix == ".csv":
            return "csv"
        if suffix == ".txt":
            return "txt"
        if "HTML" in selected_filter:
            return "html"
        if "Markdown" in selected_filter:
            return "md"
        if "CSV" in selected_filter:
            return "csv"
        if "文本" in selected_filter:
            return "txt"
        return "xml"

    @staticmethod
    def _ensure_extension(output_path: str, export_mode: str) -> str:
        suffix = Path(output_path).suffix.lower()
        expected = {
            "xml": ".xml",
            "html": ".html",
            "md": ".md",
            "csv": ".csv",
            "txt": ".txt",
        }.get(export_mode, "")
        if expected and suffix != expected:
            return f"{output_path}{expected}"
        return output_path

    def _start_batch_compare(self) -> None:
        default_output = str(Path.cwd() / "batch_diff_report.html")
        dialog = BatchCompareDialog(
            str(self.left_folder_path or ""),
            str(self.right_folder_path or ""),
            default_output,
            self._selected_revision("left"),
            self._selected_revision("right"),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        if not values["left_root"] or not values["right_root"] or not values["output_path"]:
            QMessageBox.information(self, "提示", "请填写左右根路径和报告文件。")
            return
        left_kind = infer_source_kind(values["left_root"])
        right_kind = infer_source_kind(values["right_root"])
        if SOURCE_GOOGLE_SHEETS in {left_kind, right_kind}:
            QMessageBox.warning(self, "无法批量对比", "批量对比当前支持本地目录和 SVN 根路径，不支持 Google Sheets。")
            return

        try:
            result = self._run_batch_compare(values, left_kind, right_kind)
        except Exception as exc:
            self._show_friendly_error("批量对比失败", exc, hint="请检查左右根路径、SVN 鉴权和输出路径是否可用。")
            return
        QMessageBox.information(
            self,
            "批量对比完成",
            "已对比 {files} 个同名文件，发现 {diff_files} 个有差异。\n报告已保存到:\n{path}".format(
                files=result["matched_files"],
                diff_files=result["diff_files"],
                path=result["output_path"],
            ),
        )

    def _run_batch_compare(self, values: dict[str, str], left_kind: str, right_kind: str) -> dict[str, object]:
        left_root = values["left_root"]
        right_root = values["right_root"]
        output_format = values["format"]
        output_path = self._ensure_extension(values["output_path"], output_format)
        left_revision = self._batch_revision_for_kind(left_kind, values.get("left_revision") or self._selected_revision("left"))
        right_revision = self._batch_revision_for_kind(right_kind, values.get("right_revision") or self._selected_revision("right"))
        left_entries = self._batch_file_entries(left_root, left_kind, left_revision)
        right_entries = self._batch_file_entries(right_root, right_kind, right_revision)
        all_names = sorted(set(left_entries) | set(right_entries), key=str.lower)
        if not all_names:
            raise ValueError("左右根路径下没有可对比的表格文件。")

        comparison_options = self._comparison_options()
        manual_key_fields = self._manual_key_fields()
        file_summaries: list[dict] = []
        matched_files = 0
        diff_files = 0

        for name in all_names:
            left_entry = left_entries.get(name)
            right_entry = right_entries.get(name)
            if left_entry is None:
                file_summaries.append({"file": name, "status": "右独有", "note": "左侧缺少同名文件。", "rows": []})
                continue
            if right_entry is None:
                file_summaries.append({"file": name, "status": "左独有", "note": "右侧缺少同名文件。", "rows": []})
                continue
            matched_files += 1
            try:
                left_source = self._batch_workbook_source(left_kind, left_entry.path, left_root, left_revision)
                right_source = self._batch_workbook_source(right_kind, right_entry.path, right_root, right_revision)
                left_workbook = load_workbook_from_source(left_source)
                right_workbook = load_workbook_from_source(right_source)
                alignments = self._build_batch_alignments(
                    left_workbook,
                    right_workbook,
                    comparison_options,
                    manual_key_fields,
                )
                rows = build_diff_report_rows(alignments, comparison_options)
                if rows:
                    diff_files += 1
                file_summaries.append(
                    {
                        "file": name,
                        "status": "有差异" if rows else "无差异",
                        "note": f"{len(rows)} 个差异字段" if rows else "",
                        "rows": rows,
                    }
                )
            except Exception as exc:
                file_summaries.append({"file": name, "status": "失败", "note": str(exc), "rows": []})

        self._write_batch_compare_report(output_path, output_format, file_summaries)
        self.statusBar().showMessage(f"批量对比完成，报告已保存到: {output_path}", 8000)
        return {"output_path": output_path, "matched_files": matched_files, "diff_files": diff_files}

    def _start_batch_merge(self) -> None:
        default_output_dir = str(Path.cwd() / "batch_merged")
        dialog = BatchMergeDialog(
            str(self.left_folder_path or ""),
            str(self.right_folder_path or ""),
            default_output_dir,
            str(self.template_source_combo.currentData() or "left"),
            self._selected_revision("left"),
            self._selected_revision("right"),
            self,
        )
        if dialog.exec() != QDialog.Accepted:
            return
        values = dialog.values()
        if not values["left_root"] or not values["right_root"] or not values["output_dir"]:
            QMessageBox.information(self, "提示", "请填写左右根路径和输出目录。")
            return
        left_kind = infer_source_kind(values["left_root"])
        right_kind = infer_source_kind(values["right_root"])
        if SOURCE_GOOGLE_SHEETS in {left_kind, right_kind}:
            QMessageBox.warning(self, "无法批量合并", "批量合并当前支持本地目录和 SVN 根路径，不支持 Google Sheets。")
            return

        try:
            result = self._run_batch_merge(values, left_kind, right_kind)
        except Exception as exc:
            self._show_friendly_error("批量合并失败", exc, hint="请检查左右根路径、SVN 鉴权和输出目录是否可用。")
            return
        skipped = result["skipped_files"]
        message = "已合并 {merged} 个文件，跳过 {skipped} 个。\n输出目录:\n{path}".format(
            merged=result["merged_files"],
            skipped=skipped,
            path=result["output_dir"],
        )
        if result["summary_path"]:
            message += f"\n摘要:\n{result['summary_path']}"
        QMessageBox.information(self, "批量合并完成", message)

    def _run_batch_merge(self, values: dict[str, str], left_kind: str, right_kind: str) -> dict[str, object]:
        left_root = values["left_root"]
        right_root = values["right_root"]
        output_dir = Path(values["output_dir"])
        template_side = values["template_source"]
        output_dir.mkdir(parents=True, exist_ok=True)

        left_revision = self._batch_revision_for_kind(left_kind, values.get("left_revision") or self._selected_revision("left"))
        right_revision = self._batch_revision_for_kind(right_kind, values.get("right_revision") or self._selected_revision("right"))
        left_entries = self._batch_file_entries(left_root, left_kind, left_revision)
        right_entries = self._batch_file_entries(right_root, right_kind, right_revision)
        all_names = sorted(set(left_entries) | set(right_entries), key=str.lower)
        if not all_names:
            raise ValueError("左右根路径下没有可合并的表格文件。")

        comparison_options = self._comparison_options()
        manual_key_fields = self._manual_key_fields()
        summary_rows: list[dict[str, str]] = []
        merged_files = 0

        for name in all_names:
            left_entry = left_entries.get(name)
            right_entry = right_entries.get(name)
            if left_entry is None:
                summary_rows.append({"file": name, "status": "跳过", "output": "", "note": "左侧缺少同名文件。"})
                continue
            if right_entry is None:
                summary_rows.append({"file": name, "status": "跳过", "output": "", "note": "右侧缺少同名文件。"})
                continue
            try:
                left_source = self._batch_workbook_source(left_kind, left_entry.path, left_root, left_revision)
                right_source = self._batch_workbook_source(right_kind, right_entry.path, right_root, right_revision)
                left_workbook = load_workbook_from_source(left_source)
                right_workbook = load_workbook_from_source(right_source)
                template_workbook = left_workbook if template_side == "left" else right_workbook
                compare_workbook = right_workbook if template_side == "left" else left_workbook
                template_format = template_workbook.source_meta.get("file_format", "xml")
                if template_workbook.source_kind == SOURCE_GOOGLE_SHEETS or template_format != "xml":
                    summary_rows.append(
                        {
                            "file": name,
                            "status": "跳过",
                            "output": "",
                            "note": "模板来源必须是 XML；xlsx/csv/Google Sheets 不能作为合并输出模板。",
                        }
                    )
                    continue

                alignments = self._build_batch_alignments(
                    left_workbook,
                    right_workbook,
                    comparison_options,
                    manual_key_fields,
                )
                output_path = output_dir / f"{Path(name).stem}.xml"
                export_workbook(
                    template_workbook,
                    compare_workbook,
                    alignments,
                    output_path,
                    self._active_merge_rule(),
                )
                merged_files += 1
                summary_rows.append(
                    {
                        "file": name,
                        "status": "已合并",
                        "output": str(output_path),
                        "note": f"{len(alignments)} 个工作表",
                    }
                )
            except Exception as exc:
                summary_rows.append({"file": name, "status": "失败", "output": "", "note": str(exc)})

        summary_path = output_dir / "batch_merge_summary.csv"
        self._write_batch_merge_summary(summary_path, summary_rows)
        skipped_files = sum(1 for row in summary_rows if row["status"] != "已合并")
        self.statusBar().showMessage(f"批量合并完成，输出目录: {output_dir}", 8000)
        return {
            "output_dir": str(output_dir),
            "summary_path": str(summary_path),
            "merged_files": merged_files,
            "skipped_files": skipped_files,
        }

    @staticmethod
    def _write_batch_merge_summary(path: Path, rows: list[dict[str, str]]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["file", "status", "output", "note"])
            writer.writeheader()
            writer.writerows(rows)

    def _batch_file_entries(self, root: str, source_kind: str, revision: str = "HEAD") -> dict[str, SourceFileEntry]:
        if source_kind == SOURCE_LOCAL:
            entries = list_local_table_files(root)
        elif source_kind == SOURCE_SVN:
            entries = list_svn_xml_files(root, revision=revision)
        else:
            raise ValueError(f"不支持批量对比的来源类型: {source_kind}")
        return {entry.name: entry for entry in entries}

    @staticmethod
    def _batch_revision_for_kind(source_kind: str, revision: str) -> str:
        normalized = str(revision or "").strip() or "HEAD"
        if source_kind == SOURCE_SVN and normalized.upper() == "WORKING":
            return "HEAD"
        return normalized

    def _batch_workbook_source(self, source_kind: str, file_path: str, root: str, revision: str) -> WorkbookSource:
        if source_kind == SOURCE_LOCAL:
            path = Path(file_path)
            return WorkbookSource(
                kind=SOURCE_LOCAL,
                file_path=str(path),
                display_name=path.name,
                source_root=str(path.parent),
            )
        if source_kind == SOURCE_SVN:
            target = file_path if file_path.lower().startswith("svn://") else join_source_target(root, file_path)
            return WorkbookSource(
                kind=SOURCE_SVN,
                file_path=target,
                revision=revision,
                display_name=f"{source_path_name(target)}@{revision}",
                source_root=root,
            )
        raise ValueError(f"不支持的来源类型: {source_kind}")

    def _build_batch_alignments(
        self,
        left_workbook: WorkbookData,
        right_workbook: WorkbookData,
        comparison_options: ComparisonOptions,
        manual_key_fields: list[str],
    ) -> dict[str, SheetAlignment]:
        alignments: dict[str, SheetAlignment] = {}
        sheet_names = sorted(set(left_workbook.sheet_names) | set(right_workbook.sheet_names))
        for sheet_name in sheet_names:
            key_fields = self._key_fields_for_sheet(sheet_name)
            alignments[sheet_name] = align_sheets(
                left_workbook.get_sheet(sheet_name),
                right_workbook.get_sheet(sheet_name),
                self._active_merge_rule(),
                comparison_options,
                preferred_key_fields=key_fields,
                strict_single_key=self._strict_single_key_enabled_for_sheet(sheet_name),
            )
        return alignments

    def _write_batch_compare_report(self, output_path: str, output_format: str, file_summaries: list[dict]) -> None:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        if output_format == "md":
            path.write_text(
                format_diff_report_markdown([], title="批量差异报告", file_summaries=file_summaries),
                encoding="utf-8",
            )
            return
        if output_format == "csv":
            self._write_batch_compare_csv(path, file_summaries)
            return
        path.write_text(
            format_diff_report_html([], title="批量差异报告", file_summaries=file_summaries),
            encoding="utf-8",
        )

    @staticmethod
    def _write_batch_compare_csv(path: Path, file_summaries: list[dict]) -> None:
        row_fields = [
            "sheet_name",
            "alignment_row",
            "row_status",
            "row_reason",
            "row_note",
            "left_row",
            "right_row",
            "field_key",
            "field_title",
            "first_diff_char",
            "left_value",
            "right_value",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["file", "status", "note", *row_fields])
            writer.writeheader()
            for entry in file_summaries:
                rows = entry.get("rows") or []
                if not rows:
                    writer.writerow(
                        {
                            "file": entry.get("file", ""),
                            "status": entry.get("status", ""),
                            "note": entry.get("note", ""),
                        }
                    )
                    continue
                for row in rows:
                    writer.writerow(
                        {
                            "file": entry.get("file", ""),
                            "status": entry.get("status", ""),
                            "note": entry.get("note", ""),
                            **{field: row.get(field, "") for field in row_fields},
                        }
                    )

    def _update_detail_panel(self, visible_row: int | None = None) -> None:
        alignment_model = self.middle_model or self.left_model or self.right_model
        if alignment_model is None:
            self.detail_panel.clear()
            return
        selected_rows = self._current_visible_rows()
        if visible_row is not None and visible_row not in selected_rows:
            selected_rows = [visible_row]
        if not selected_rows:
            self.detail_panel.setHtml("<b>未选择行</b>")
            return
        if len(selected_rows) > 1:
            action_hint = (
                "可以直接批量执行“取左 / 取右 / 重置自动”。<br/>"
                if not self.compare_only_mode.isChecked()
                else "当前为纯比对模式；如需人工合并，请先切到合并模式。<br/>"
            )
            self.detail_panel.setHtml(
                f"<b>已选中 {len(selected_rows)} 行</b><br/>"
                f"{action_hint}"
                "如需查看差异详情，请只选中一行。"
            )
            return
        visible_row = selected_rows[0]
        if visible_row >= alignment_model.rowCount():
            self.detail_panel.setHtml("<b>未选择行</b>")
            return

        row = alignment_model._aligned_row(visible_row)
        if hasattr(self, "row_note_input"):
            if self.row_note_input.text() != row.note:
                self.row_note_input.blockSignals(True)
                self.row_note_input.setText(row.note)
                self.row_note_input.blockSignals(False)
        header_bits = [f"<b>状态:</b> {row.status}", f"<b>说明:</b> {row.reason}"]
        if row.note:
            header_bits.append(f"<b>备注:</b> <span style='color:#8a4600;'>{self._escape(row.note)}</span>")
        html = [
            " &nbsp;&nbsp; ".join(header_bits),
            "<hr/>",
        ]
        for logical_index, binding in enumerate(alignment_model.alignment.columns, start=1):
            left_value = row.left_row.value_at(binding.left_index or -1) if row.left_row and binding.left_index else ""
            merged_value = row.merged_row.value_at(logical_index)
            right_value = row.right_row.value_at(binding.right_index or -1) if row.right_row and binding.right_index else ""
            has_diff = not compare_text_values(left_value, right_value, self._comparison_options()) or logical_index in row.conflict_columns
            if not has_diff and not (left_value or merged_value or right_value):
                continue
            title = binding.title or binding.key
            if has_diff:
                diff_hint = self._first_difference_hint(left_value, right_value)
                html.append(
                    "<div style='margin-bottom:6px;'>"
                    f"<b>{title}</b><br/>"
                    f"{diff_hint}"
                    f"<span style='color:#1155cc;'>左:</span> {self._render_diff_value_html(left_value, right_value, 'left')}<br/>"
                    f"<span style='color:#b45f06;'>中:</span> {self._escape(merged_value)}<br/>"
                    f"<span style='color:#cc0000;'>右:</span> {self._render_diff_value_html(right_value, left_value, 'right')}"
                    "</div>"
                )
        if len(html) == 2:
            html.append("<span style='color:#666;'>当前行没有可展示的差异字段。</span>")
        self.detail_panel.setHtml("".join(html))

    def _first_difference_hint(self, left_value: str, right_value: str) -> str:
        matcher = SequenceMatcher(None, left_value or "", right_value or "", autojunk=False)
        for tag, left_start, left_end, right_start, right_end in matcher.get_opcodes():
            if tag == "equal":
                continue
            index = min(left_start, right_start) + 1
            return (
                "<span style='color:#8a4600; font-size:11px;'>"
                f"首个差异位置: 第 {index} 个字符"
                "</span><br/>"
            )
        return ""

    def _render_diff_value_html(self, value: str, other_value: str, side: str) -> str:
        text = value or ""
        other = other_value or ""
        if text == other:
            return self._escape(text)
        background = "#dbeafe" if side == "left" else "#ffe7cc"
        foreground = "#1d4ed8" if side == "left" else "#b54708"
        parts: list[str] = []
        matcher = SequenceMatcher(None, text, other, autojunk=False)
        for tag, start, end, _, _ in matcher.get_opcodes():
            segment = text[start:end]
            if not segment:
                continue
            escaped = self._escape(segment)
            if tag == "equal":
                parts.append(escaped)
            else:
                parts.append(
                    f"<span style='background:{background}; color:{foreground}; font-weight:700;'>{escaped}</span>"
                )
        if not parts:
            return "<span style='color:#999;'>(空)</span>"
        return "".join(parts)

    @staticmethod
    def _escape(value: str) -> str:
        return (
            value.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br/>")
        )

    def _connect_selection_models(self) -> None:
        for table in self._selection_tables(include_middle=True):
            selection_model = table.selectionModel()
            if selection_model is not None:
                token = id(selection_model)
                if token in self._connected_selection_models:
                    continue
                self._connected_selection_models.add(token)
                selection_model.selectionChanged.connect(
                    lambda selected, deselected, source=table: self._on_selection_changed(source)
                )

    def _selection_tables(self, include_middle: bool = False) -> list[QTableView]:
        tables = [self.left_table, self.right_table]
        if include_middle or (self.middle_model is not None and not self.compare_only_mode.isChecked()):
            tables.insert(1, self.middle_table)
        return tables

    def _on_selection_changed(self, source_table: QTableView) -> None:
        if self._syncing_selection:
            return
        selection_model = source_table.selectionModel()
        if selection_model is None:
            return
        rows = sorted(index.row() for index in selection_model.selectedRows())
        if not rows:
            self._update_detail_panel()
            return
        self._select_rows(rows, source=source_table)
        self._update_detail_panel(rows[0] if len(rows) == 1 else None)

    def _select_rows(self, rows: list[int], source: QTableView | None = None) -> None:
        self._syncing_selection = True
        try:
            for table in self._selection_tables(include_middle=True):
                if table is source:
                    continue
                selection_model = table.selectionModel()
                model = table.model()
                if selection_model is None or model is None:
                    continue
                selection_model.clearSelection()
                for row in rows:
                    if 0 <= row < model.rowCount():
                        index = model.index(row, 0)
                        selection_model.select(index, QItemSelectionModel.Select | QItemSelectionModel.Rows)
        finally:
            self._syncing_selection = False

    def _mark_row_deleted(self, row: AlignedRow, mode: str) -> None:
        source_row = row.left_row or row.right_row or row.merged_row
        row.merged_row = clone_row_with_values(source_row, {})
        row.merged_row.cells = []
        row.merged_row.kind = "blank"
        row.status = "deleted"
        row.reason = f"按{mode}侧结果删除"
        row.conflict_columns.clear()

    def apply_initial_file_paths(self, file_paths: list[str]) -> None:
        local_paths = []
        for raw_path in file_paths[:2]:
            path = Path(str(raw_path or "").strip())
            if path.is_file() and path.suffix.lower() in {".xml", ".xlsx", ".csv"}:
                local_paths.append(path)
        if not local_paths:
            return
        self._suspend_auto_selection = True
        try:
            for side, path in zip(("left", "right"), local_paths):
                combo = self.left_source_combo if side == "left" else self.right_source_combo
                combo.setCurrentIndex(max(0, combo.findData(SOURCE_LOCAL)))
                if side == "left":
                    self.left_folder_path = str(path.parent)
                else:
                    self.right_folder_path = str(path.parent)
                self._update_folder_labels()
                self._refresh_file_combo(side)
                self._set_pending_file(side, str(path), remember=False, load_revisions=False)
        finally:
            self._suspend_auto_selection = False
        self._remember_current_settings()
        if len(local_paths) >= 2:
            self._start_compare()
            return
        self.statusBar().showMessage("已从命令行载入文件，可直接开始比对。", 5000)


def _apply_global_style(app: QApplication) -> None:
    qss = """
    QWidget { color: #334155; background-color: #f8fafc; }
    QPushButton { background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px 12px; }
    QPushButton:hover { background-color: #f1f5f9; border: 1px solid #94a3b8; }
    QPushButton#primaryAction { background-color: #0ea5e9; color: #ffffff; border: none; }
    QPushButton#primaryAction:hover { background-color: #0284c7; }
    QLineEdit, QTextEdit { background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px 8px; }
    QLineEdit:focus, QTextEdit:focus { border: 1px solid #0ea5e9; }
    QTableView { background-color: #ffffff; gridline-color: #f1f5f9; border: none; alternate-background-color: #f8fafc; }
    QHeaderView::section { background-color: #f1f5f9; border: none; border-right: 1px solid #e2e8f0; border-bottom: 2px solid #cbd5e1; padding: 4px; padding-left: 6px; font-weight: bold; color: #475569; }
    QListWidget, QTreeWidget { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; }
    QListWidget::item, QTreeWidget::item { padding: 4px; border-radius: 4px; margin: 2px 4px; }
    QListWidget::item:hover, QTreeWidget::item:hover { background-color: #f1f5f9; }
    QListWidget::item:selected, QTreeWidget::item:selected { background-color: #e0f2fe; color: #0369a1; }
    """
    app.setStyleSheet(qss)
def run(initial_paths: list[str] | None = None) -> None:
    app = QApplication(sys.argv)
    _apply_preferred_font(app)
    _apply_global_style(app)
    window = MainWindow()
    if initial_paths:
        window.apply_initial_file_paths(initial_paths)
    window.show()
    sys.exit(app.exec())


def _apply_preferred_font(app: QApplication) -> None:
    available_families = set(QFontDatabase.families())
    for family in PREFERRED_UI_FONTS:
        if family in available_families:
            font = QFont(family, 9)
            font.setStyleHint(QFont.SansSerif)
            app.setFont(font)
            return
