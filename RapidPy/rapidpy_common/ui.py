from __future__ import annotations

from pathlib import Path

from PySide6 import QtGui, QtWidgets

from .palette import GOLD, MAROON


def apply_liquid_glass_theme(app: QtWidgets.QApplication) -> None:
    """Apply the shared Apple-inspired liquid glass styling."""
    assets_dir = Path(__file__).resolve().parent / "assets"
    arrow_down = (assets_dir / "arrow_down.svg").as_posix()
    arrow_up = (assets_dir / "arrow_up.svg").as_posix()

    app.setStyle("Fusion")
    font = QtGui.QFont("SF Pro Text", 10)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Avenir Next", 10)
    if not QtGui.QFontInfo(font).exactMatch():
        font = QtGui.QFont("Segoe UI", 10)
    app.setFont(font)

    app.setStyleSheet(
        f"""
        QWidget {{
            background: #f3eee2;
            color: #2f2827;
        }}
        QFrame#card {{
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(255, 255, 255, 0.65);
            border-radius: 24px;
        }}
        QLabel#title {{
            font-size: 24px;
            font-weight: 760;
            color: {MAROON};
        }}
        QLabel#subtitle {{
            color: #61534d;
            margin-bottom: 4px;
        }}
        QLabel#valuePill {{
            background: rgba(255, 255, 255, 0.82);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 16px;
            padding: 8px 10px;
            font-weight: 650;
        }}
        QPlainTextEdit#console {{
            background: rgba(28, 20, 19, 0.88);
            color: #fff2c9;
            border-radius: 14px;
            border: 1px solid rgba(255, 205, 52, 0.32);
            padding: 8px;
            selection-background-color: {MAROON};
        }}
        QPushButton {{
            background: rgba(255, 255, 255, 0.76);
            border: 1px solid rgba(255, 255, 255, 0.75);
            border-radius: 14px;
            padding: 9px 14px;
        }}
        QPushButton:hover {{
            background: rgba(255, 255, 255, 0.92);
        }}
        QPushButton:pressed {{
            background: rgba(232, 226, 216, 0.95);
        }}
        QPushButton#accent {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {MAROON}, stop:1 #5a0013);
            color: #fff9eb;
            border: 1px solid rgba(255, 255, 255, 0.26);
            font-weight: 680;
        }}
        QPushButton#accent:hover {{
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8a0220, stop:1 #650016);
        }}
        QPushButton#accent:pressed {{
            background: #5a0013;
        }}
        QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {{
            border: 1px solid rgba(255, 255, 255, 0.82);
            background: rgba(255, 255, 255, 0.72);
            border-radius: 12px;
            padding: 7px;
            selection-background-color: {MAROON};
            selection-color: #ffffff;
        }}
        QComboBox {{
            padding-right: 34px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            margin: 3px;
            border: none;
            border-radius: 10px;
            background: rgba(122, 2, 25, 0.12);
        }}
        QComboBox::drop-down:hover {{
            background: rgba(122, 2, 25, 0.2);
        }}
        QComboBox::drop-down:pressed {{
            background: rgba(122, 2, 25, 0.28);
        }}
        QComboBox::down-arrow {{
            image: url({arrow_down});
            width: 14px;
            height: 14px;
        }}
        QAbstractSpinBox {{
            padding-right: 52px;
        }}
        QAbstractSpinBox::up-button,
        QAbstractSpinBox::down-button {{
            width: 24px;
            border: none;
            border-radius: 9px;
            background: rgba(122, 2, 25, 0.12);
            margin-right: 3px;
        }}
        QAbstractSpinBox::up-button {{
            subcontrol-origin: border;
            subcontrol-position: top right;
            margin-top: 3px;
            margin-bottom: 1px;
        }}
        QAbstractSpinBox::down-button {{
            subcontrol-origin: border;
            subcontrol-position: bottom right;
            margin-top: 1px;
            margin-bottom: 3px;
        }}
        QAbstractSpinBox::up-button:hover,
        QAbstractSpinBox::down-button:hover {{
            background: rgba(122, 2, 25, 0.2);
        }}
        QAbstractSpinBox::up-button:pressed,
        QAbstractSpinBox::down-button:pressed {{
            background: rgba(122, 2, 25, 0.28);
        }}
        QAbstractSpinBox::up-arrow {{
            image: url({arrow_up});
            width: 13px;
            height: 13px;
        }}
        QAbstractSpinBox::down-arrow {{
            image: url({arrow_down});
            width: 13px;
            height: 13px;
        }}
        QHeaderView::section {{
            background: rgba(255, 255, 255, 0.85);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 6px;
            padding: 6px;
            color: #4d3a39;
        }}
        QTableWidget {{
            background: rgba(255, 255, 255, 0.8);
            alternate-background-color: rgba(255, 255, 255, 0.65);
            border: 1px solid rgba(122, 2, 25, 0.16);
            border-radius: 12px;
            gridline-color: rgba(122, 2, 25, 0.12);
        }}
        QTableWidget::item:selected {{
            background: rgba(255, 205, 52, 0.38);
            color: #251f1e;
        }}
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
            background-color: {GOLD};
            border: 1px solid {MAROON};
        }}
        """
    )


def apply_card_shadow(widget: QtWidgets.QWidget) -> None:
    shadow = QtWidgets.QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(34)
    shadow.setOffset(0, 10)
    shadow.setColor(QtGui.QColor(35, 25, 25, 48))
    widget.setGraphicsEffect(shadow)
