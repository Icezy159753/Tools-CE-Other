from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

from PyQt6.QtCore import QObject, QRunnable, Qt, QThreadPool, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtWidgets import QApplication, QLabel, QMainWindow, QMessageBox, QProgressBar, QVBoxLayout, QWidget


APP_TITLE = "Tools Other CE Updater"
ICON_FILE = "Iconapp.ico"

BG = "#0F0E17"
SURFACE = "#1A1926"
LAVENDER = "#C4B5FD"
ROSE = "#F9A8D4"
TEXT = "#E2E0F0"
TEXT2 = "#8E8CA8"
BORDER = "#2E2C42"


def _resource_path(name: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / name


def _wait_for_process_exit(pid: int, timeout_sec: int = 120) -> None:
    if pid <= 0:
        return
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        try:
            if sys.platform == "win32":
                import ctypes

                SYNCHRONIZE = 0x00100000
                handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, pid)
                if not handle:
                    return
                result = ctypes.windll.kernel32.WaitForSingleObject(handle, 500)
                ctypes.windll.kernel32.CloseHandle(handle)
                if result == 0:
                    return
            else:
                os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.5)


def _replace_file_with_retry(source: Path, target: Path, attempts: int = 40) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    for _ in range(attempts):
        try:
            os.replace(source, target)
            return
        except PermissionError:
            time.sleep(0.5)
        except OSError:
            time.sleep(0.5)
    raise RuntimeError(f"ไม่สามารถแทนที่ไฟล์ได้: {target}")


def _run_update(target: str, source: str, launch: str, launch_args: str, pid: int) -> str:
    target_path = Path(target)
    source_path = Path(source)
    launch_path = Path(launch)

    if not source_path.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์อัปเดต: {source_path}")

    _wait_for_process_exit(pid)
    time.sleep(0.8)
    _replace_file_with_retry(source_path, target_path)

    cmd = [str(launch_path)]
    if launch_args.strip():
        cmd.append(launch_args.strip())
    subprocess.Popen(cmd, cwd=str(target_path.parent))
    return str(target_path)


class _Signals(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)


class _Worker(QRunnable):
    def __init__(self, fn, *args) -> None:
        super().__init__()
        self.fn = fn
        self.args = args
        self.signals = _Signals()

    def run(self) -> None:
        try:
            result = self.fn(*self.args)
            self.signals.finished.emit(result)
        except Exception as ex:
            self.signals.error.emit(str(ex))


class UpdaterWindow(QMainWindow):
    def __init__(self, target: str, source: str, launch: str, launch_args: str, pid: int) -> None:
        super().__init__()
        self._pool = QThreadPool.globalInstance()
        self._build_ui()
        self._worker = _Worker(_run_update, target, source, launch, launch_args, pid)
        self._worker.signals.finished.connect(self._on_done)
        self._worker.signals.error.connect(self._on_error)
        self._pool.start(self._worker)

    def _build_ui(self) -> None:
        self.setWindowTitle(APP_TITLE)
        self.setFixedSize(520, 220)
        icon_path = _resource_path(ICON_FILE)
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(14)

        title = QLabel("กำลังอัปเดตโปรแกรม")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet(f"color: {TEXT};")

        self._status = QLabel("กำลังดาวน์โหลดเสร็จแล้ว และกำลังแทนที่ไฟล์เวอร์ชันเดิม...")
        self._status.setWordWrap(True)
        self._status.setStyleSheet(f"color: {TEXT2}; font-size: 12px;")

        progress = QProgressBar()
        progress.setRange(0, 0)
        progress.setFixedHeight(8)
        progress.setStyleSheet(
            f"""
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {SURFACE};
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ROSE}, stop:1 {LAVENDER});
                border-radius: 4px;
            }}
            """
        )

        note = QLabel("โปรแกรมจะเปิดเวอร์ชันใหม่ให้อัตโนมัติหลังอัปเดตเสร็จ")
        note.setStyleSheet(f"color: {TEXT2}; font-size: 11px;")

        layout.addWidget(title)
        layout.addWidget(self._status)
        layout.addWidget(progress)
        layout.addWidget(note)

        self.setStyleSheet(
            f"""
            QMainWindow, QWidget {{
                background-color: {BG};
            }}
            """
        )

    def _on_done(self, target_path: str) -> None:
        self._status.setText(f"อัปเดตเสร็จแล้ว กำลังเปิดเวอร์ชันใหม่...\n{target_path}")
        QApplication.instance().quit()

    def _on_error(self, msg: str) -> None:
        QMessageBox.critical(self, "Updater Error", f"อัปเดตไม่สำเร็จ\n\n{msg}")
        QApplication.instance().quit()


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--launch", required=True)
    parser.add_argument("--launch-args", default="")
    parser.add_argument("--pid", type=int, default=0)
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    window = UpdaterWindow(
        target=args.target,
        source=args.source,
        launch=args.launch,
        launch_args=args.launch_args,
        pid=args.pid,
    )
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
