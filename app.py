"""
app.py — Other Recode Tool GUI (PyQt6)
Pastel dark theme, 2-tab layout:
  Tab 1: Export Coding Sheet  (Rawdata + SPSS → Excel)
  Tab 2: Apply Recodes        (Rawdata + filled Coding Sheet → save)
"""

from __future__ import annotations

import logging
import sys
import traceback
import ctypes
import json
import os
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest

import pandas as pd
from PyQt6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QRunnable,
    QSize,
    Qt,
    QThreadPool,
    QTimer,
    pyqtSignal,
    QObject,
)
from PyQt6.QtGui import QColor, QFont, QIcon, QPalette
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QScrollArea,
    QSplitter,
    QStatusBar,
    QTableView,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import core


APP_ID = "songklod.toolsothercev1"
ICON_FILE = "Iconapp.ico"
APP_VERSION = "1.0.6"
UPDATE_CONFIG_FILE = "update_config.json"
UPDATE_CONFIG_EXAMPLE_FILE = "update_config.example.json"
GITHUB_REPO_DEFAULT = "Icezy159753/Tools-CE-Other"
GITHUB_ASSET_NAME_DEFAULT = "Tools Other CE V1.exe"
GITHUB_UPDATER_ASSET_NAME_DEFAULT = "Tools Other CE Updater.exe"


def _resource_path(name: str) -> Path:
    """Return resource path that works in source and PyInstaller onefile mode."""
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / name


def _app_base_dir() -> Path:
    """Return directory of the running app/exe for external config and downloads."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _parse_version(version: str) -> tuple[int, ...]:
    parts = []
    for token in str(version).strip().split("."):
        digits = "".join(ch for ch in token if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts or [0])


def _is_version_newer(latest: str, current: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def _load_update_config() -> dict:
    config = {
        "provider": os.environ.get("TOOLS_OTHER_UPDATE_PROVIDER", "github").strip() or "github",
        "repo": os.environ.get("TOOLS_OTHER_UPDATE_REPO", GITHUB_REPO_DEFAULT).strip(),
        "asset_name": os.environ.get("TOOLS_OTHER_UPDATE_ASSET", GITHUB_ASSET_NAME_DEFAULT).strip(),
        "updater_asset_name": os.environ.get("TOOLS_OTHER_UPDATE_UPDATER_ASSET", GITHUB_UPDATER_ASSET_NAME_DEFAULT).strip(),
        "auto_check": True,
    }
    config_path = _app_base_dir() / UPDATE_CONFIG_FILE
    if config_path.exists():
        try:
            loaded = json.loads(config_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                config.update(loaded)
        except Exception:
            pass
    return config


def _ensure_update_config_example() -> None:
    example_path = _app_base_dir() / UPDATE_CONFIG_EXAMPLE_FILE
    if example_path.exists():
        return
    example = {
        "provider": "github",
        "repo": GITHUB_REPO_DEFAULT,
        "asset_name": GITHUB_ASSET_NAME_DEFAULT,
        "updater_asset_name": GITHUB_UPDATER_ASSET_NAME_DEFAULT,
        "auto_check": True,
    }
    try:
        example_path.write_text(
            json.dumps(example, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _fetch_json(url: str) -> dict:
    req = urlrequest.Request(
        url,
        headers={
            "User-Agent": f"Tools-Other-CE/{APP_VERSION}",
            "Accept": "application/json",
        },
    )
    with urlrequest.urlopen(req, timeout=20) as resp:
        payload = resp.read().decode("utf-8")
    data = json.loads(payload)
    if not isinstance(data, dict):
        raise ValueError("update metadata must be a JSON object")
    return data


def _fetch_github_release_metadata(repo: str, asset_name: str) -> dict:
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    data = _fetch_json(api_url)
    version = str(data.get("tag_name", "")).strip().lstrip("vV")
    if not version:
        raise ValueError("GitHub release missing tag_name")

    assets = data.get("assets") or []
    assets_map: dict[str, str] = {}
    for asset in assets:
        name = str(asset.get("name", "")).strip()
        url = str(asset.get("browser_download_url", "")).strip()
        if name and url:
            assets_map[name] = url
    download_url = _find_asset_download_url(assets_map, asset_name, prefer_keyword="")

    return {
        "version": version,
        "download_url": download_url,
        "notes": str(data.get("body", "")).strip(),
        "published_at": str(data.get("published_at", "")).strip(),
        "release_url": str(data.get("html_url", "")).strip(),
        "asset_name": asset_name,
        "assets_map": assets_map,
        "available_assets": list(assets_map.keys()),
    }


def _find_asset_download_url(assets_map: dict[str, str], preferred_name: str, prefer_keyword: str) -> str:
    preferred_name = str(preferred_name).strip()
    if preferred_name and preferred_name in assets_map:
        return assets_map[preferred_name]

    lower_map = {name.lower(): url for name, url in assets_map.items()}
    if preferred_name and preferred_name.lower() in lower_map:
        return lower_map[preferred_name.lower()]

    if prefer_keyword:
        keyword = prefer_keyword.lower()
        for name, url in assets_map.items():
            lname = name.lower()
            if keyword in lname and lname.endswith(".exe"):
                return url

    for name, url in assets_map.items():
        if name.lower().endswith(".exe"):
            return url
    return ""


def _check_for_updates() -> dict:
    try:
        config = _load_update_config()
        provider = str(config.get("provider", "github")).strip().lower()
        if provider != "github":
            return {
                "configured": False,
                "current_version": APP_VERSION,
            }
        repo = str(config.get("repo", "")).strip()
        if not repo:
            return {
                "configured": False,
                "current_version": APP_VERSION,
            }
        metadata = _fetch_github_release_metadata(
            repo,
            str(config.get("asset_name", "")).strip(),
        )
        updater_asset_name = str(config.get("updater_asset_name", GITHUB_UPDATER_ASSET_NAME_DEFAULT)).strip()
        assets_map = metadata.get("assets_map", {})
        updater_download_url = _find_asset_download_url(
            assets_map,
            updater_asset_name,
            prefer_keyword="updater",
        )
        return {
            "configured": True,
            "current_version": APP_VERSION,
            "latest_version": metadata["version"],
            "download_url": metadata["download_url"],
            "notes": metadata["notes"],
            "published_at": metadata["published_at"],
            "release_url": metadata.get("release_url", ""),
            "asset_name": metadata.get("asset_name", ""),
            "updater_asset_name": updater_asset_name,
            "updater_download_url": updater_download_url,
            "available_assets": metadata.get("available_assets", []),
            "repo": repo,
            "update_available": _is_version_newer(metadata["version"], APP_VERSION),
        }
    except urlerror.URLError as ex:
        raise ValueError(f"เช็กอัปเดตไม่สำเร็จ: เชื่อมต่อ GitHub ไม่ได้ ({ex.reason})") from ex


def _download_update(download_url: str, output_path: str) -> str:
    req = urlrequest.Request(
        download_url,
        headers={"User-Agent": f"Tools-Other-CE/{APP_VERSION}"},
    )
    with urlrequest.urlopen(req, timeout=120) as resp, open(output_path, "wb") as fh:
        fh.write(resp.read())
    return output_path


def _prepare_update_package(
    app_download_url: str,
    updater_download_url: str,
    latest_version: str,
) -> dict:
    try:
        base_dir = _app_base_dir()
        updates_dir = base_dir / "_updates"
        updates_dir.mkdir(parents=True, exist_ok=True)
        app_temp = updates_dir / f"Tools Other CE V1-{latest_version}.exe"
        updater_temp = updates_dir / "Tools Other CE Updater.exe"
        _download_update(app_download_url, str(app_temp))
        if updater_download_url:
            _download_update(updater_download_url, str(updater_temp))
        return {
            "app_path": str(app_temp),
            "updater_path": str(updater_temp),
        }
    except urlerror.URLError as ex:
        raise ValueError(f"ดาวน์โหลดอัปเดตไม่สำเร็จ: เชื่อมต่อ GitHub ไม่ได้ ({ex.reason})") from ex


def _set_windows_taskbar_icon(hwnd: int, icon_path: Path) -> None:
    """Force small/big window icons so Windows taskbar picks them up consistently."""
    if sys.platform != "win32" or not icon_path.exists():
        return
    try:
        user32 = ctypes.windll.user32
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        hicon = user32.LoadImageW(
            None,
            str(icon_path),
            IMAGE_ICON,
            0,
            0,
            LR_LOADFROMFILE | LR_DEFAULTSIZE,
        )
        if hicon:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Logging → GUI console
# ---------------------------------------------------------------------------

class _LogEmitter(QObject):
    append_log = pyqtSignal(str, int)


class _QtLogHandler(logging.Handler):
    """Routes logger messages to a QTextEdit widget."""

    def __init__(self, text_widget: QTextEdit) -> None:
        super().__init__()
        self._widget = text_widget
        self._emitter = _LogEmitter()
        self._emitter.append_log.connect(self._append_to_widget)

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self._emitter.append_log.emit(msg, record.levelno)

    def _append_to_widget(self, msg: str, level: int) -> None:
        if level >= logging.ERROR:
            color = "#F9A8D4"  # pastel pink
        elif level >= logging.WARNING:
            color = "#FDE68A"  # pastel amber
        else:
            color = "#94A3B8"  # muted slate
        self._widget.append(f'<span style="color:{color};">{msg}</span>')
        self._widget.verticalScrollBar().setValue(
            self._widget.verticalScrollBar().maximum()
        )


# ---------------------------------------------------------------------------
# Pandas → QAbstractTableModel (for QTableView)
# ---------------------------------------------------------------------------

class PandasModel(QAbstractTableModel):
    """Wraps a pandas DataFrame for display in QTableView."""

    def __init__(self, df: pd.DataFrame | None = None) -> None:
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()

    def load(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df.copy()
        self.endResetModel()

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._df.columns)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or self._df.empty:
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            val = self._df.iloc[index.row(), index.column()]
            return "" if pd.isna(val) else str(val)
        if role == Qt.ItemDataRole.BackgroundRole:
            col_name = self._df.columns[index.column()]
            if col_name == core.NEW_CODE_COL:
                val = self._df.iloc[index.row(), index.column()]
                if pd.isna(val) or str(val).strip() == "":
                    return QColor("#3B2040")  # muted plum = empty
                return QColor("#1A3A2A")      # muted teal = filled
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._df.columns[section]) if section < len(self._df.columns) else None
        return str(section + 1)


# ---------------------------------------------------------------------------
# Worker (runs core logic in background thread)
# ---------------------------------------------------------------------------

class _WorkerSignals(QObject):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)


class _Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs) -> None:
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = _WorkerSignals()

    def run(self) -> None:
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.finished.emit(result)
        except Exception as ex:
            if isinstance(ex, ValueError):
                self.signals.error.emit(str(ex))
            else:
                self.signals.error.emit(traceback.format_exc())


# ---------------------------------------------------------------------------
# Color palette — Pastel on Dark
# ---------------------------------------------------------------------------

BG       = "#0F0E17"   # deep navy-black
SURFACE  = "#1A1926"   # card surface
SURFACE2 = "#232136"   # elevated surface
ROSE     = "#F9A8D4"   # pastel rose/pink
LAVENDER = "#C4B5FD"   # pastel lavender
MINT     = "#86EFAC"   # pastel mint
SKY      = "#7DD3FC"   # pastel sky blue
PEACH    = "#FDBA74"   # pastel peach/orange
LILAC    = "#D8B4FE"   # pastel lilac
CREAM    = "#FDE68A"   # pastel cream/yellow
TEXT     = "#E2E0F0"   # primary text
TEXT2    = "#8E8CA8"   # secondary text
BORDER   = "#2E2C42"   # subtle border
HOVER    = "#2A283E"   # hover state


def _stylesheet() -> str:
    return f"""
    * {{
        font-family: 'Segoe UI', 'Noto Sans Thai', sans-serif;
    }}
    QMainWindow, QWidget {{
        background-color: {BG};
        color: {TEXT};
        font-size: 13px;
    }}

    /* ── Tabs ──────────────────────────────────────────────── */
    QTabWidget::pane {{
        border: 1px solid {BORDER};
        background: {SURFACE};
        border-radius: 0px 10px 10px 10px;
        top: -1px;
    }}
    QTabBar {{
        background: transparent;
    }}
    QTabBar::tab {{
        background: {SURFACE2};
        color: {TEXT2};
        padding: 12px 40px;
        margin-right: 2px;
        border: 1px solid {BORDER};
        border-bottom: none;
        border-radius: 10px 10px 0 0;
        font-size: 14px;
        font-weight: 600;
        min-width: 120px;
    }}
    QTabBar::tab:selected {{
        color: {LAVENDER};
        background: {SURFACE};
        border-color: {LAVENDER};
        border-bottom: 2px solid {SURFACE};
        margin-bottom: -1px;
        font-weight: 700;
    }}
    QTabBar::tab:hover:!selected {{
        color: {TEXT};
        background: {HOVER};
        border-color: #3D3B58;
    }}

    /* ── Buttons ───────────────────────────────────────────── */
    QPushButton {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 {LAVENDER}, stop:1 {ROSE});
        color: {BG};
        border: none;
        border-radius: 10px;
        padding: 10px 28px;
        font-weight: 700;
        font-size: 13px;
    }}
    QPushButton:hover {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #D8C8FE, stop:1 #FBC0DC);
    }}
    QPushButton:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #B8A0F0, stop:1 #E890B8);
    }}
    QPushButton:disabled {{
        background: {SURFACE2};
        color: #4A4868;
    }}

    QPushButton#browse {{
        background: {SURFACE2};
        border: 1px solid #3D3A58;
        color: {LILAC};
        padding: 4px 12px;
        font-size: 12px;
        font-weight: 600;
        border-radius: 8px;
    }}
    QPushButton#browse:hover {{
        background: #2E2C48;
        border-color: {LAVENDER};
        color: {LAVENDER};
    }}

    QPushButton#success {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 {MINT}, stop:1 {SKY});
        color: {BG};
    }}
    QPushButton#success:hover {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #A0F5C0, stop:1 #A0E0FC);
    }}
    QPushButton#success:disabled {{
        background: {SURFACE2};
        color: #4A4868;
    }}

    /* ── Inputs ────────────────────────────────────────────── */
    QLineEdit {{
        background-color: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 8px;
        padding: 8px 12px;
        color: {TEXT};
        selection-background-color: {LAVENDER};
    }}
    QLineEdit:focus {{ border-color: {LAVENDER}; }}

    /* ── Table ─────────────────────────────────────────────── */
    QTableView {{
        background-color: {SURFACE};
        alternate-background-color: #1E1D30;
        border: 1px solid {BORDER};
        border-radius: 10px;
        gridline-color: #26243A;
        selection-background-color: rgba(196, 181, 253, 0.2);
        selection-color: {TEXT};
        font-size: 13px;
    }}
    QTableView::item {{
        padding: 6px 10px;
        border-bottom: 1px solid #1E1D30;
    }}
    QHeaderView::section {{
        background-color: {SURFACE2};
        color: {LILAC};
        padding: 10px 12px;
        border: none;
        border-right: 1px solid {BORDER};
        border-bottom: 1px solid {BORDER};
        font-weight: 700;
        font-size: 12px;
        text-transform: uppercase;
    }}

    /* ── Console ───────────────────────────────────────────── */
    QTextEdit {{
        background-color: #0C0B14;
        color: #7A7890;
        border: 1px solid {BORDER};
        border-radius: 8px;
        font-family: 'Cascadia Code', 'Consolas', 'Courier New', monospace;
        font-size: 12px;
        padding: 4px;
    }}

    /* ── Labels ────────────────────────────────────────────── */
    QLabel#section_title {{
        color: {LAVENDER};
        font-size: 13px;
        font-weight: 700;
    }}

    QLabel#version_badge {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 {SKY}, stop:1 {LAVENDER});
        color: {BG};
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 11px;
        padding: 3px 12px;
        font-size: 11px;
        font-weight: 800;
    }}

    QPushButton#update_action {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 {MINT}, stop:1 {SKY});
        color: {BG};
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 11px;
        padding: 4px 16px;
        font-size: 12px;
        font-weight: 800;
    }}
    QPushButton#update_action:hover {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #A8F8C8, stop:1 #A9E6FF);
    }}
    QPushButton#update_action:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #74D7A0, stop:1 #69C9F4);
    }}
    QPushButton#update_action:disabled {{
        background: {SURFACE2};
        color: #6A6888;
        border-color: #3B3956;
    }}

    /* ── Progress ──────────────────────────────────────────── */
    QProgressBar {{
        border: none;
        border-radius: 3px;
        background-color: {SURFACE};
        height: 6px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {ROSE}, stop:0.5 {LAVENDER}, stop:1 {SKY});
        border-radius: 3px;
    }}

    /* ── Cards ─────────────────────────────────────────────── */
    QFrame#card {{
        background-color: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 12px;
    }}
    QFrame#kpi_card {{
        background-color: {SURFACE2};
        border: 1px solid #2E2C42;
        border-radius: 14px;
    }}
    QLabel#kpi_value {{
        color: {TEXT};
        font-size: 26px;
        font-weight: 800;
    }}
    QLabel#kpi_title {{
        color: {TEXT2};
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}

    /* ── Misc ──────────────────────────────────────────────── */
    QSplitter::handle {{
        background: {BORDER};
        height: 1px;
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 6px;
        border-radius: 3px;
    }}
    QScrollBar::handle:vertical {{
        background: #3A3858;
        border-radius: 3px;
        min-height: 20px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {LAVENDER};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    QStatusBar {{
        background: {SURFACE};
        color: {TEXT2};
        border-top: 1px solid {BORDER};
        font-size: 12px;
        padding: 2px 8px;
    }}
    """


# ---------------------------------------------------------------------------
# Path display
# ---------------------------------------------------------------------------

class _PathDisplayEdit(QLineEdit):
    """Read-only field that keeps long file names readable."""

    def __init__(self) -> None:
        super().__init__()
        self._display_name = ""
        self._full_path_text = ""
        self._name_label: QLabel | None = None

    def bind_name_label(self, label: QLabel) -> None:
        self._name_label = label

    def set_display_path(self, path: str) -> None:
        self.setProperty("full_path", path)
        self.setToolTip(path)
        self._display_name = Path(path).name
        self._full_path_text = path
        self._refresh_text()

    def _refresh_text(self) -> None:
        if not self._display_name:
            return
        self.setText(self._display_name)
        if self._name_label is not None:
            self._name_label.setText(self._display_name)
            self._name_label.setToolTip(self._full_path_text)
            self._name_label.setStyleSheet(f"color: {MINT}; font-size: 11px; font-weight: 600;")

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_text()


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------

def _file_row(label: str, hint: str) -> tuple[QWidget, QLineEdit, QPushButton]:
    """Return (container_widget, hidden path holder, browse button)."""
    container = QFrame()
    container.setMinimumHeight(64)
    container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
    row = QHBoxLayout(container)
    row.setContentsMargins(8, 0, 8, 0)
    row.setSpacing(12)

    # Left: colored bar
    bar = QFrame()
    bar.setFixedSize(3, 36)
    bar.setStyleSheet(f"background: {LAVENDER}; border-radius: 1px;")

    # Middle: title + hint stacked
    info = QWidget()
    info_lay = QVBoxLayout(info)
    info_lay.setContentsMargins(0, 0, 0, 0)
    info_lay.setSpacing(3)

    title_lbl = QLabel(label)
    title_lbl.setWordWrap(True)
    title_lbl.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 600;")
    title_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

    status_lbl = QLabel(hint)
    status_lbl.setWordWrap(True)
    status_lbl.setStyleSheet(f"color: {TEXT2}; font-size: 11px;")
    status_lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

    info_lay.addWidget(title_lbl)
    info_lay.addWidget(status_lbl)
    info.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    info.setMinimumWidth(0)

    # Right: browse button
    browse_btn = QPushButton("Browse")
    browse_btn.setObjectName("browse")
    browse_btn.setMinimumSize(76, 32)
    browse_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
    browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)

    row.addWidget(bar)
    row.addWidget(info)
    row.addWidget(browse_btn)
    row.setStretch(1, 1)

    # Hidden edit for path storage
    edit = _PathDisplayEdit()
    edit.bind_name_label(status_lbl)
    edit.setVisible(False)

    return container, edit, browse_btn


def _stat_badge(value: str, label: str, color: str) -> QFrame:
    """Pastel KPI card."""
    frame = QFrame()
    frame.setObjectName("kpi_card")
    frame.setMinimumHeight(74)
    frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(16, 12, 16, 12)
    layout.setSpacing(4)

    val_lbl = QLabel(value)
    val_lbl.setObjectName("kpi_value")
    val_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
    val_lbl.setStyleSheet(f"color: {color}; font-size: 28px;")

    txt_lbl = QLabel(label)
    txt_lbl.setObjectName("kpi_title")
    txt_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

    layout.addWidget(val_lbl)
    layout.addWidget(txt_lbl)
    return frame


def _separator() -> QFrame:
    line = QFrame()
    line.setFixedHeight(1)
    line.setStyleSheet(f"background-color: {BORDER};")
    return line


# ---------------------------------------------------------------------------
# Tab 1 — Export Coding Sheet
# ---------------------------------------------------------------------------

class ExportTab(QWidget):
    """Tab 1: load Rawdata + SPSS → preview → Export coding sheet."""

    def __init__(self, log_widget: QTextEdit, status_bar: QStatusBar) -> None:
        super().__init__()
        self._log = log_widget
        self._status = status_bar
        self._pool = QThreadPool.globalInstance()
        self._model = PandasModel()
        self._result: core.Phase1Result | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 12)
        root.setSpacing(14)

        # ── Header ────────────────────────────────────────────────────
        title = QLabel("Export Other Coding Sheet")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {TEXT};")
        root.addWidget(title)

        desc = QLabel(
            "โหลด Rawdata (.xlsx) และไฟล์ SPSS (.sav) เพื่อดึงแถวที่ตอบ 'อื่นๆ ระบุ' "
            "และ Export เป็น Coding Sheet ให้ทีมลง New Code"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {TEXT2}; font-size: 12px; line-height: 1.4;")
        root.addWidget(desc)

        # ── File pickers card ─────────────────────────────────────────
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 8, 12, 8)
        card_layout.setSpacing(0)

        raw_w, self._raw_edit, raw_btn = _file_row(
            "Rawdata Excel (.xlsx)",
            "ไฟล์ข้อมูลดิบ — ต้องมีคอลัมน์ _oth เช่น Q3_oth, Q5_oth",
        )
        raw_btn.clicked.connect(lambda: self._pick_file(self._raw_edit, "Excel (*.xlsx *.xls)"))

        spss_w, self._spss_edit, spss_btn = _file_row(
            "SPSS Labels (.sav)",
            "ไฟล์ SPSS ที่มี Value Labels — ใช้ detect code 'อื่นๆ' อัตโนมัติ",
        )
        spss_btn.clicked.connect(lambda: self._pick_file(self._spss_edit, "SPSS (*.sav)"))

        card_layout.addWidget(raw_w)
        card_layout.addWidget(_separator())
        card_layout.addWidget(spss_w)

        root.addWidget(card)

        # ── Run button + progress ─────────────────────────────────────
        self._run_btn = QPushButton("CodeSheet")
        self._run_btn.setMinimumHeight(44)
        self._run_btn.setMinimumWidth(220)
        self._run_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._run_btn.clicked.connect(self._run)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.setVisible(False)
        self._progress.setFixedHeight(6)

        run_row = QHBoxLayout()
        run_row.addStretch()
        run_row.addWidget(self._run_btn)
        root.addLayout(run_row)
        root.addWidget(self._progress)

        # ── Dashboard badges ──────────────────────────────────────────
        stats_row = QHBoxLayout()
        stats_row.setSpacing(12)
        self._badge_q = _stat_badge("—", "Questions", SKY)
        self._badge_r = _stat_badge("—", "Other Rows", PEACH)
        stats_row.addWidget(self._badge_q)
        stats_row.addWidget(self._badge_r)
        stats_row.addStretch()
        root.addLayout(stats_row)

        # ── Preview table ─────────────────────────────────────────────
        preview_lbl = QLabel("PREVIEW")
        preview_lbl.setObjectName("section_title")
        root.addWidget(preview_lbl)

        self._table = QTableView()
        self._table.setModel(self._model)
        self._table.setAlternatingRowColors(True)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self._table.verticalHeader().setDefaultSectionSize(32)
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(False)
        self._table.setMinimumHeight(150)
        self._table.setSortingEnabled(True)
        root.addWidget(self._table, stretch=1)

    # ── Helpers ─────────────────────────────────────────────────────

    @staticmethod
    def _set_path(edit: QLineEdit, path: str) -> None:
        if isinstance(edit, _PathDisplayEdit):
            edit.set_display_path(path)
            return
        edit.setProperty("full_path", path)
        edit.setText(Path(path).name)
        edit.setToolTip(path)

    @staticmethod
    def _get_path(edit: QLineEdit) -> str:
        return edit.property("full_path") or edit.text()

    def _pick_file(self, edit: QLineEdit, filter_str: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์", "", filter_str)
        if path:
            self._set_path(edit, path)

    def _run(self) -> None:
        try:
            raw = self._get_path(self._raw_edit).strip()
            spss = self._get_path(self._spss_edit).strip()
            if not raw or not spss:
                QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือก Rawdata และ SPSS ก่อน")
                return

            out = str(Path(raw).parent / "CodeSheet.xlsx")
            mapping_report = core.inspect_phase1_column_mapping(Path(raw), Path(spss), allow_order_fallback=False)

            if mapping_report.unresolved_excel_cols:
                preview_report = core.inspect_phase1_column_mapping(Path(raw), Path(spss), allow_order_fallback=True)
                unresolved_after_fallback = preview_report.unresolved_excel_cols
                if unresolved_after_fallback:
                    examples = "\n".join(f"- {col}" for col in unresolved_after_fallback[:15])
                    more = ""
                    if len(unresolved_after_fallback) > 15:
                        more = f"\n... และอีก {len(unresolved_after_fallback) - 15} คอลัมน์"
                    QMessageBox.warning(
                        self,
                        "พบตัวแปรที่แมพไม่ได้",
                        "โปรแกรมพบตัวแปรจาก Excel ที่แมพกับ SPSS ไม่ได้ และ fallback ตามลำดับแล้วก็ยังไม่เจอ\n\n"
                        f"{examples}{more}\n\n"
                        "กรุณาตรวจสอบไฟล์ก่อน แล้วค่อยสร้าง CodeSheet ใหม่",
                    )
                    return
                chosen_report = preview_report
                allow_order_fallback = True
            else:
                chosen_report = mapping_report
                allow_order_fallback = False

            if chosen_report.mismatched_matches:
                mismatch_lines = [
                    f"{excel_col} > {spss_col}"
                    for excel_col, spss_col in chosen_report.mismatched_matches
                ]
                preview_text = "\n".join(mismatch_lines[:20])
                more = ""
                if len(mismatch_lines) > 20:
                    more = f"\n... และอีก {len(mismatch_lines) - 20} รายการ"
                mode_text = (
                    "โปรแกรมสามารถ fallback ตามลำดับคอลัมน์ได้ดังนี้:"
                    if chosen_report.fallback_matches
                    else "โปรแกรมสามารถแมพตัวแปรชื่อไม่ตรงกันได้ดังนี้:"
                )
                reply = QMessageBox.question(
                    self,
                    "ยืนยัน Variable Mapping",
                    "พบตัวแปรที่ชื่อไม่ตรงกันระหว่าง Excel กับ SPSS\n"
                    f"{mode_text}\n\n"
                    f"{preview_text}{more}\n\n"
                    "ยืนยันสร้าง Code Sheet ไหม",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    self._status.showMessage("ยกเลิกการสร้าง CodeSheet")
                    return

            self._run_btn.setEnabled(False)
            self._progress.setVisible(True)
            self._status.showMessage("กำลังดึงข้อมูล...")

            worker = _Worker(core.phase1_export, Path(raw), Path(spss), Path(out), allow_order_fallback)
            worker.signals.finished.connect(self._on_done)
            worker.signals.error.connect(self._on_error)
            self._pool.start(worker)
        except Exception:
            self._run_btn.setEnabled(True)
            self._progress.setVisible(False)
            QMessageBox.critical(self, "Error", traceback.format_exc())
            self._status.showMessage("เกิดข้อผิดพลาด")

    def _on_done(self, result: core.Phase1Result) -> None:
        try:
            self._result = result
            self._progress.setVisible(False)
            self._run_btn.setEnabled(True)

            labels_q = self._badge_q.findChildren(QLabel)
            if labels_q:
                labels_q[0].setText(str(result.n_questions))
            labels_r = self._badge_r.findChildren(QLabel)
            if labels_r:
                labels_r[0].setText(str(result.n_rows))

            self._model.load(result.coding_df)
            self._table.resizeColumnsToContents()
            msg = f"Saved {result.n_rows} rows across {result.n_questions} questions to CodeSheet"
            self._status.showMessage(msg)
            QMessageBox.information(
                self,
                "Export Complete",
                f"{msg}\n\nFile: {result.output_path}",
            )
        except Exception:
            self._progress.setVisible(False)
            self._run_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", traceback.format_exc())
            self._status.showMessage("เกิดข้อผิดพลาด")

    def _on_error(self, msg: str) -> None:
        self._progress.setVisible(False)
        self._run_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)
        self._status.showMessage("เกิดข้อผิดพลาด")


# ---------------------------------------------------------------------------
# Tab 2 — Apply Recodes
# ---------------------------------------------------------------------------

class ApplyTab(QWidget):
    """Tab 2: load Rawdata + filled Coding Sheet → apply recodes → save."""

    def __init__(self, log_widget: QTextEdit, status_bar: QStatusBar) -> None:
        super().__init__()
        self._log = log_widget
        self._status = status_bar
        self._pool = QThreadPool.globalInstance()
        self._model_log = PandasModel()
        self._result: core.Phase2Result | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 12)
        root.setSpacing(14)

        # ── Header ────────────────────────────────────────────────────
        title = QLabel("ลง Code ให้ Rawdata")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {TEXT};")
        root.addWidget(title)

        desc = QLabel(
            "โหลด Rawdata ต้นฉบับ และ Coding Sheet ที่ลง New_Code เรียบร้อยแล้ว "
            "จากนั้น Save เป็น Rawdata ฉบับ Recode พร้อม Log สรุป"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {TEXT2}; font-size: 12px;")
        root.addWidget(desc)

        # ── File pickers card ─────────────────────────────────────────
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 8, 12, 8)
        card_layout.setSpacing(0)

        raw_w, self._raw_edit, raw_btn = _file_row(
            "Rawdata ต้นฉบับ (.xlsx)",
            "ไฟล์ rawdata เดิม (ก่อน recode) — จะไม่ถูกแก้ไขโดยตรง",
        )
        raw_btn.clicked.connect(lambda: self._pick_file(self._raw_edit, "Excel (*.xlsx *.xls)"))

        coding_w, self._coding_edit, coding_btn = _file_row(
            "Coding Sheet (.xlsx)",
            "ไฟล์ที่ Export จาก Phase 1 และลง New_Code เสร็จแล้ว",
        )
        coding_btn.clicked.connect(lambda: self._pick_file(self._coding_edit, "Excel (*.xlsx *.xls)"))

        out_raw_w, self._out_raw_edit, out_raw_btn = _file_row(
            "Save Rawdata Recoded (.xlsx)",
            "ไฟล์ Rawdata ใหม่ที่ถูก recode แล้ว",
        )
        out_raw_btn.clicked.connect(self._pick_out_raw)

        card_layout.addWidget(raw_w)
        card_layout.addWidget(_separator())
        card_layout.addWidget(coding_w)
        card_layout.addWidget(_separator())
        card_layout.addWidget(out_raw_w)

        root.addWidget(card)

        # ── Run button + progress ─────────────────────────────────────
        self._run_btn = QPushButton("ลง Code")
        self._run_btn.setMinimumHeight(44)
        self._run_btn.setMinimumWidth(220)
        self._run_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._run_btn.clicked.connect(self._run)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.setVisible(False)
        self._progress.setFixedHeight(6)

        run_row = QHBoxLayout()
        run_row.addStretch()
        run_row.addWidget(self._run_btn)
        root.addLayout(run_row)
        root.addWidget(self._progress)

        # ── Stats badges ──────────────────────────────────────────────
        stats_row = QGridLayout()
        stats_row.setHorizontalSpacing(12)
        stats_row.setVerticalSpacing(10)
        self._badge_applied  = _stat_badge("—", "Applied", MINT)
        self._badge_skipped  = _stat_badge("—", "Skipped", CREAM)
        self._badge_notfound = _stat_badge("—", "Not Found", ROSE)
        stats_row.addWidget(self._badge_applied, 0, 0)
        stats_row.addWidget(self._badge_skipped, 0, 1)
        stats_row.addWidget(self._badge_notfound, 1, 0)

        self._open_raw_btn = QPushButton("Open Output Folder")
        self._open_raw_btn.setObjectName("success")
        self._open_raw_btn.setMinimumHeight(38)
        self._open_raw_btn.setEnabled(False)
        self._open_raw_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._open_raw_btn.clicked.connect(self._open_folder)
        stats_row.addWidget(self._open_raw_btn, 1, 1, alignment=Qt.AlignmentFlag.AlignRight)
        stats_row.setColumnStretch(0, 1)
        stats_row.setColumnStretch(1, 1)

        root.addLayout(stats_row)

        # ── Preview log table ─────────────────────────────────────────
        preview_lbl = QLabel("RECODE LOG")
        preview_lbl.setObjectName("section_title")
        root.addWidget(preview_lbl)

        self._table = QTableView()
        self._table.setModel(self._model_log)
        self._table.setAlternatingRowColors(True)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self._table.verticalHeader().setDefaultSectionSize(32)
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(False)
        self._table.setMinimumHeight(150)
        self._table.setSortingEnabled(True)
        root.addWidget(self._table, stretch=1)

    # ── Helpers ─────────────────────────────────────────────────────

    @staticmethod
    def _set_path(edit: QLineEdit, path: str) -> None:
        if isinstance(edit, _PathDisplayEdit):
            edit.set_display_path(path)
            return
        edit.setProperty("full_path", path)
        edit.setText(Path(path).name)
        edit.setToolTip(path)

    @staticmethod
    def _get_path(edit: QLineEdit) -> str:
        return edit.property("full_path") or edit.text()

    def _pick_file(self, edit: QLineEdit, filter_str: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์", "", filter_str)
        if path:
            self._set_path(edit, path)

    def _pick_out_raw(self) -> None:
        raw_path = self._get_path(self._raw_edit)
        default = str(Path(raw_path).parent / "Rawdata_CE Complete.xlsx") if raw_path else "Rawdata_CE Complete.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "บันทึก Rawdata Recoded", default, "Excel (*.xlsx)")
        if path:
            self._set_path(self._out_raw_edit, path if path.endswith(".xlsx") else path + ".xlsx")

    def _run(self) -> None:
        raw = self._get_path(self._raw_edit).strip()
        coding = self._get_path(self._coding_edit).strip()
        if not raw or not coding:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือก Rawdata และ Coding Sheet ก่อน")
            return

        raw_path = Path(raw)
        out_raw = self._get_path(self._out_raw_edit).strip() or str(raw_path.parent / "Rawdata_CE Complete.xlsx")
        self._set_path(self._out_raw_edit, out_raw)

        self._run_btn.setEnabled(False)
        self._open_raw_btn.setEnabled(False)
        self._progress.setVisible(True)
        self._status.showMessage("กำลัง Apply Recodes...")

        worker = _Worker(
            core.phase2_apply,
            raw_path, Path(coding), Path(out_raw),
        )
        worker.signals.finished.connect(self._on_done)
        worker.signals.error.connect(self._on_error)
        self._pool.start(worker)

    def _on_done(self, result: core.Phase2Result) -> None:
        self._result = result
        self._progress.setVisible(False)
        self._run_btn.setEnabled(True)
        self._open_raw_btn.setEnabled(True)

        def _set_badge(badge: QFrame, value: str) -> None:
            labels = badge.findChildren(QLabel)
            if labels:
                labels[0].setText(value)

        _set_badge(self._badge_applied,  str(result.n_applied))
        _set_badge(self._badge_skipped,  str(result.n_skipped))
        _set_badge(self._badge_notfound, str(result.n_not_found))

        self._model_log.load(result.log_df)
        self._table.resizeColumnsToContents()

        msg = (
            f"Recode applied: {result.n_applied}"
            + (f"  |  Skipped: {result.n_skipped}" if result.n_skipped else "")
            + (f"  |  Not found: {result.n_not_found}" if result.n_not_found else "")
        )
        self._status.showMessage(msg)
        QMessageBox.information(self, "Done!", msg)

    def _on_error(self, msg: str) -> None:
        self._progress.setVisible(False)
        self._run_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)
        self._status.showMessage("เกิดข้อผิดพลาด")

    def _open_folder(self) -> None:
        if self._result:
            import subprocess, platform
            folder = str(self._result.output_rawdata_path.parent)
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer "{folder}"')
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])


# ---------------------------------------------------------------------------
# Tab 3 — AI CodeFrame
# ---------------------------------------------------------------------------

class CodeFrameTab(QWidget):
    """Tab 3: load Coding Sheet and generate AI codeframe workbook."""

    def __init__(self, log_widget: QTextEdit, status_bar: QStatusBar) -> None:
        super().__init__()
        self._log = log_widget
        self._status = status_bar
        self._pool = QThreadPool.globalInstance()
        self._model_preview = PandasModel()
        self._result: core.CodeFrameResult | None = None
        self._build_ui()

    @staticmethod
    def _set_path(edit: QLineEdit, path: str) -> None:
        if isinstance(edit, _PathDisplayEdit):
            edit.set_display_path(path)
            return
        edit.setProperty("full_path", path)
        edit.setText(Path(path).name)
        edit.setToolTip(path)

    @staticmethod
    def _get_path(edit: QLineEdit) -> str:
        return edit.property("full_path") or edit.text()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 12)
        root.setSpacing(14)

        title = QLabel("AI Group Code (Demo)")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {TEXT};")
        root.addWidget(title)

        desc = QLabel(
            "เลือก Coding Sheet ที่มี Open_Text แล้วส่งแต่ละข้อให้ OpenRouter "
            "ช่วยจัดกลุ่มคำตอบเป็น CodeFrame แยกชีต"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {TEXT2}; font-size: 12px;")
        root.addWidget(desc)

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 8, 12, 8)
        card_layout.setSpacing(0)

        coding_w, self._coding_edit, coding_btn = _file_row(
            "Coding Sheet (.xlsx)",
            "ไฟล์ codesheet.xlsx ที่มี Open_Text และพร้อมใช้สร้าง CodeFrame",
        )
        coding_btn.clicked.connect(lambda: self._pick_file(self._coding_edit, "Excel (*.xlsx *.xls)"))
        card_layout.addWidget(coding_w)
        card_layout.addWidget(_separator())

        api_row = QFrame()
        api_layout = QVBoxLayout(api_row)
        api_layout.setContentsMargins(8, 10, 8, 10)
        api_layout.setSpacing(4)
        api_lbl = QLabel("OpenRouter API Key")
        api_lbl.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 600;")
        self._api_key_edit = QLineEdit()
        self._api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self._api_key_edit.setPlaceholderText("เว้นว่างได้ถ้าตั้ง OPENROUTER_API_KEY ไว้แล้ว")
        api_hint = QLabel("ใช้ค่าจากช่องนี้ก่อน ถ้าว่างจะ fallback ไปที่ตัวแปรแวดล้อม OPENROUTER_API_KEY")
        api_hint.setWordWrap(True)
        api_hint.setStyleSheet(f"color: {TEXT2}; font-size: 11px;")
        api_layout.addWidget(api_lbl)
        api_layout.addWidget(self._api_key_edit)
        api_layout.addWidget(api_hint)
        card_layout.addWidget(api_row)
        card_layout.addWidget(_separator())

        model_row = QFrame()
        model_layout = QVBoxLayout(model_row)
        model_layout.setContentsMargins(8, 10, 8, 10)
        model_layout.setSpacing(4)
        model_lbl = QLabel("Model")
        model_lbl.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 600;")
        self._model_edit = QLineEdit(core.OPENROUTER_MODEL_DEFAULT)
        self._model_edit.setPlaceholderText(core.OPENROUTER_MODEL_DEFAULT)
        model_layout.addWidget(model_lbl)
        model_layout.addWidget(self._model_edit)
        card_layout.addWidget(model_row)

        root.addWidget(card)

        self._run_btn = QPushButton("AI Group Code (Demo)")
        self._run_btn.setMinimumHeight(44)
        self._run_btn.setMinimumWidth(220)
        self._run_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self._run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._run_btn.clicked.connect(self._run)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.setVisible(False)
        self._progress.setFixedHeight(6)

        run_row = QHBoxLayout()
        run_row.addStretch()
        run_row.addWidget(self._run_btn)
        root.addLayout(run_row)
        root.addWidget(self._progress)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(12)
        self._badge_groups = _stat_badge("—", "Groups", SKY)
        self._badge_rows = _stat_badge("—", "AI Rows", PEACH)
        stats_row.addWidget(self._badge_groups)
        stats_row.addWidget(self._badge_rows)
        stats_row.addStretch()
        root.addLayout(stats_row)

        preview_lbl = QLabel("PREVIEW")
        preview_lbl.setObjectName("section_title")
        root.addWidget(preview_lbl)

        self._table = QTableView()
        self._table.setModel(self._model_preview)
        self._table.setAlternatingRowColors(True)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self._table.verticalHeader().setDefaultSectionSize(32)
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(False)
        self._table.setMinimumHeight(150)
        self._table.setSortingEnabled(True)
        root.addWidget(self._table, stretch=1)

    def _pick_file(self, edit: QLineEdit, filter_str: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์", "", filter_str)
        if path:
            self._set_path(edit, path)

    def _run(self) -> None:
        coding = self._get_path(self._coding_edit).strip()
        if not coding:
            QMessageBox.warning(self, "ข้อมูลไม่ครบ", "กรุณาเลือก Coding Sheet ก่อน")
            return

        api_key = self._api_key_edit.text().strip()
        model = self._model_edit.text().strip() or core.OPENROUTER_MODEL_DEFAULT
        out = str(Path(coding).parent / "codeframe.xlsx")

        self._run_btn.setEnabled(False)
        self._progress.setVisible(True)
        self._status.showMessage("กำลังสร้าง AI CodeFrame...")

        worker = _Worker(
            core.generate_codeframe_with_ai,
            Path(coding),
            Path(out),
            api_key,
            model,
        )
        worker.signals.finished.connect(self._on_done)
        worker.signals.error.connect(self._on_error)
        self._pool.start(worker)

    def _on_done(self, result: core.CodeFrameResult) -> None:
        self._result = result
        self._progress.setVisible(False)
        self._run_btn.setEnabled(True)

        labels_g = self._badge_groups.findChildren(QLabel)
        if labels_g:
            labels_g[0].setText(str(result.n_groups))
        labels_r = self._badge_rows.findChildren(QLabel)
        if labels_r:
            labels_r[0].setText(str(result.n_rows))

        self._model_preview.load(result.codeframe_df)
        self._table.resizeColumnsToContents()
        msg = f"Saved {result.n_rows} AI rows across {result.n_groups} groups to CodeFrame"
        self._status.showMessage(msg)
        QMessageBox.information(
            self,
            "CodeFrame Complete",
            f"{msg}\n\nModel: {result.model}\nFile: {result.output_path}",
        )

    def _on_error(self, msg: str) -> None:
        self._progress.setVisible(False)
        self._run_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)
        self._status.showMessage("เกิดข้อผิดพลาด")


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        _ensure_update_config_example()
        self.setWindowTitle("Tools Other CE V1")
        self.setMinimumSize(QSize(900, 750))
        self._pool = QThreadPool.globalInstance()
        self._update_check_running = False
        self._download_running = False
        self._pending_update_result: dict | None = None
        icon_path = _resource_path(ICON_FILE)
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self._setup_logging()
        config = _load_update_config()
        if bool(config.get("auto_check")) and str(config.get("repo", "")).strip():
            QTimer.singleShot(1200, self._check_updates_silent)

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Status bar
        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status.showMessage("Ready")

        # Splitter: tabs (top) + log console (bottom)
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setHandleWidth(2)

        # Tabs
        self._tabs = QTabWidget()
        self._log_widget = QTextEdit()
        self._log_widget.setReadOnly(True)
        self._log_widget.setMaximumHeight(120)
        self._log_widget.setPlaceholderText("Log output will appear here...")

        self._tab1 = ExportTab(self._log_widget, self._status)
        self._tab2 = ApplyTab(self._log_widget, self._status)
        self._tab3 = CodeFrameTab(self._log_widget, self._status)
        self._tabs.addTab(self._make_scroll_tab(self._tab1), "  Phase1 - CodeSheet  ")
        self._tabs.addTab(self._make_scroll_tab(self._tab2), "  Phase2 - ลง Code  ")
        self._tabs.addTab(self._make_scroll_tab(self._tab3), "  Phase2 - AI Group Code (Demo)  ")

        splitter.addWidget(self._tabs)

        # Log panel
        log_frame = QFrame()
        log_frame.setObjectName("card")
        log_layout = QVBoxLayout(log_frame)
        log_layout.setContentsMargins(12, 8, 12, 8)
        log_layout.setSpacing(4)

        log_title_row = QHBoxLayout()
        log_lbl = QLabel("CONSOLE")
        log_lbl.setObjectName("section_title")
        log_lbl.setStyleSheet(f"color: {TEXT2}; font-size: 10px; font-weight: 700; letter-spacing: 1px;")

        version_lbl = QLabel(f"Version {APP_VERSION}")
        version_lbl.setObjectName("version_badge")
        version_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_lbl.setMinimumHeight(26)

        self._update_btn = QPushButton("Check Update")
        self._update_btn.setObjectName("update_action")
        self._update_btn.setMinimumSize(128, 28)
        self._update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._update_btn.clicked.connect(self._check_updates_manual)

        clear_btn = QPushButton("Clear")
        clear_btn.setObjectName("browse")
        clear_btn.setFixedSize(56, 22)
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.clicked.connect(self._log_widget.clear)

        log_title_row.addWidget(log_lbl)
        log_title_row.addStretch()
        log_title_row.addWidget(version_lbl)
        log_title_row.addWidget(self._update_btn)
        log_title_row.addWidget(clear_btn)
        log_layout.addLayout(log_title_row)
        log_layout.addWidget(self._log_widget)

        splitter.addWidget(log_frame)
        splitter.setSizes([640, 120])

        root.addWidget(splitter)

    @staticmethod
    def _make_scroll_tab(content: QWidget) -> QScrollArea:
        area = QScrollArea()
        area.setWidget(content)
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        return area

    def center_on_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if not screen:
            return
        available = screen.availableGeometry()
        x = available.x() + (available.width() - self.width()) // 2
        y = available.y() + (available.height() - self.height()) // 2
        self.move(x, y)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        icon_path = _resource_path(ICON_FILE)
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
            _set_windows_taskbar_icon(int(self.winId()), icon_path)

    def _setup_logging(self) -> None:
        handler = _QtLogHandler(self._log_widget)
        handler.setFormatter(
            logging.Formatter("%(asctime)s  %(levelname)s  %(message)s", "%H:%M:%S")
        )
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def _check_updates_manual(self) -> None:
        self._start_update_check(silent=False)

    def _check_updates_silent(self) -> None:
        self._start_update_check(silent=True)

    def _start_update_check(self, silent: bool) -> None:
        if self._update_check_running:
            return
        self._update_check_running = True
        self._update_btn.setEnabled(False)
        if not silent:
            self._status.showMessage("กำลังเช็กอัปเดต...")
        worker = _Worker(_check_for_updates)
        worker.signals.finished.connect(lambda result: self._on_update_check_done(result, silent))
        worker.signals.error.connect(lambda msg: self._on_update_check_error(msg, silent))
        self._pool.start(worker)

    def _on_update_check_done(self, result: dict, silent: bool) -> None:
        self._update_check_running = False
        self._update_btn.setEnabled(True)
        if not result.get("configured"):
            self._status.showMessage("ยังไม่ได้ตั้งค่า GitHub update")
            if not silent:
                QMessageBox.information(
                    self,
                    "Check Update",
                    "ยังไม่ได้ตั้งค่า GitHub update\n\n"
                    f"ให้สร้างไฟล์ `{UPDATE_CONFIG_FILE}` ไว้ข้างโปรแกรม แล้วใส่ `repo`",
                )
            return

        if not result.get("update_available"):
            self._status.showMessage(f"เป็นเวอร์ชันล่าสุดแล้ว (v{APP_VERSION})")
            if not silent:
                QMessageBox.information(
                    self,
                    "Check Update",
                    f"เป็นเวอร์ชันล่าสุดแล้ว\n\nCurrent version: v{APP_VERSION}",
                )
            return

        latest = result.get("latest_version", "")
        notes = result.get("notes", "")
        published_at = result.get("published_at", "")
        self._pending_update_result = result
        detail_lines = [
            f"Current version: v{APP_VERSION}",
            f"Latest version: v{latest}",
            f"Repo: {result.get('repo', '')}",
        ]
        if published_at:
            detail_lines.append(f"Published: {published_at}")
        if notes:
            detail_lines.append("")
            detail_lines.append("Release notes:")
            detail_lines.append(notes)
        detail_lines.append("")
        detail_lines.append("ต้องการอัปเดตตอนนี้ไหม")
        reply = QMessageBox.question(
            self,
            "Update Available",
            "\n".join(detail_lines),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage(f"พบอัปเดต v{latest}")
            return

        download_url = str(result.get("download_url", "")).strip()
        if not download_url:
            QMessageBox.warning(
                self,
                "Update Available",
                "พบเวอร์ชันใหม่ แต่ release ยังไม่มีไฟล์ .exe ให้ดาวน์โหลด",
            )
            return
        updater_download_url = str(result.get("updater_download_url", "")).strip()
        if not updater_download_url:
            available_assets = result.get("available_assets") or []
            asset_text = ""
            if available_assets:
                asset_text = "\n\nAssets ใน release ล่าสุด:\n" + "\n".join(
                    f"- {name}" for name in available_assets
                )
            QMessageBox.warning(
                self,
                "Update Available",
                "พบเวอร์ชันใหม่ แต่ release ยังไม่มีไฟล์ Updater.exe\n\n"
                "ให้ปล่อย release ใหม่ที่มีทั้ง Tools Other CE V1.exe และ Tools Other CE Updater.exe ก่อน"
                + asset_text,
            )
            return
        self._start_update_download(download_url, updater_download_url, latest)

    def _on_update_check_error(self, msg: str, silent: bool) -> None:
        self._update_check_running = False
        self._update_btn.setEnabled(True)
        self._status.showMessage("เช็กอัปเดตไม่สำเร็จ")
        if silent:
            logging.warning(f"Update check failed: {msg}")
            return
        QMessageBox.warning(
            self,
            "Check Update",
            f"เช็กอัปเดตไม่สำเร็จ\n\n{msg}",
        )

    def _start_update_download(self, download_url: str, updater_download_url: str, latest_version: str) -> None:
        if self._download_running:
            return
        self._download_running = True
        self._update_btn.setEnabled(False)
        self._status.showMessage("กำลังดาวน์โหลดอัปเดต...")
        worker = _Worker(_prepare_update_package, download_url, updater_download_url, latest_version)
        worker.signals.finished.connect(self._on_update_download_done)
        worker.signals.error.connect(self._on_update_download_error)
        self._pool.start(worker)

    def _launch_updater_and_exit(self, updater_path: str, new_app_path: str) -> None:
        import subprocess

        current_exe = str(Path(sys.executable if getattr(sys, "frozen", False) else Path(__file__).resolve()))
        launch_target = current_exe if getattr(sys, "frozen", False) else sys.executable
        cmd = [
            updater_path,
            "--target", current_exe,
            "--source", new_app_path,
            "--launch", launch_target,
            "--pid", str(os.getpid()),
        ]
        if not getattr(sys, "frozen", False):
            cmd.extend(["--launch-args", str(Path(__file__).resolve())])
        subprocess.Popen(cmd, cwd=str(_app_base_dir()))
        QApplication.instance().quit()

    def _on_update_download_done(self, payload: dict) -> None:
        self._download_running = False
        self._update_btn.setEnabled(True)
        app_path = str(payload.get("app_path", "")).strip()
        updater_path = str(payload.get("updater_path", "")).strip()
        if not app_path or not updater_path or not Path(updater_path).exists():
            QMessageBox.warning(
                self,
                "Download Update",
                "ดาวน์โหลดอัปเดตแล้ว แต่ไม่พบไฟล์ Updater ที่ใช้แทนเวอร์ชันเดิม",
            )
            return
        reply = QMessageBox.question(
            self,
            "Ready To Install Update",
            "ดาวน์โหลดอัปเดตเรียบร้อยแล้ว\n\n"
            "โปรแกรมจะปิดตัวเดิม แทนที่ไฟล์ด้วยเวอร์ชันใหม่ และเปิดเวอร์ชันใหม่ขึ้นมาให้อัตโนมัติ\n\n"
            "ต้องการดำเนินการต่อไหม",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            self._status.showMessage("ยกเลิกการติดตั้งอัปเดต")
            return
        self._status.showMessage("กำลังติดตั้งอัปเดต...")
        self._launch_updater_and_exit(updater_path, app_path)

    def _on_update_download_error(self, msg: str) -> None:
        self._download_running = False
        self._update_btn.setEnabled(True)
        self._status.showMessage("ดาวน์โหลดอัปเดตไม่สำเร็จ")
        QMessageBox.warning(
            self,
            "Download Update",
            f"ดาวน์โหลดอัปเดตไม่สำเร็จ\n\n{msg}",
        )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setApplicationName("Tools Other CE V1")
    icon_path = _resource_path(ICON_FILE)
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyleSheet(_stylesheet())

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(BG))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.Base, QColor(SURFACE))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#1E1D30"))
    palette.setColor(QPalette.ColorRole.Text, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(LAVENDER))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(BG))
    app.setPalette(palette)

    win = MainWindow()
    win.center_on_screen()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
