from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PySide6 import QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy System Shell")
        self.resize(960, 560)
        self.root = Path(__file__).resolve().parents[2]
        self._build_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)

        card = QtWidgets.QFrame()
        card.setObjectName("card")
        c = QtWidgets.QVBoxLayout(card)
        c.setContentsMargins(18, 18, 18, 18)

        title = QtWidgets.QLabel("Integrated Operator Shell")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Launch converted subsystem panels with one click")
        subtitle.setObjectName("subtitle")
        c.addWidget(title)
        c.addWidget(subtitle)

        buttons = [
            ("VRM Logger", "vrm_logger/main.py"),
            ("AF Tuner", "af_tuner/main.py"),
            ("Changer XY", "changer_xy_control/main.py"),
            ("Up/Down", "updown_control/main.py"),
            ("DC Motor Control", "dc_motor_control/main.py"),
        ]
        for label, relpath in buttons:
            btn = QtWidgets.QPushButton(label)
            btn.setObjectName("accent")
            btn.clicked.connect(lambda _checked=False, p=relpath: self._launch(p))
            c.addWidget(btn)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        c.addWidget(self.console, stretch=1)

        layout.addWidget(card)
        apply_card_shadow(card)

    def _launch(self, relative_path: str) -> None:
        target = self.root / relative_path
        if not target.exists():
            self.console.appendPlainText(f"Missing app entrypoint: {target}")
            return

        python = Path(sys.executable)
        subprocess.Popen([str(python), str(target)], cwd=str(target.parent))
        self.console.appendPlainText(f"Launched: {relative_path}")


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()
