"""Omarchy/Hyprland-style tiling canvas for RapidPy windows and panels.

Every panel (and, optionally, every tool window) lives in a *tile*.  Tiles
fill the canvas automatically: opening a new tile splits the focused tile
along its longer side ("dwindle"), closing one gives its space back to its
sibling, and every split can be dragged to any ratio.  Like a tiling window
manager the canvas has numbered workspaces, a monocle (fullscreen-tile)
toggle, keyboard focus/swap/resize, a launcher, and tool windows can float
out of the grid and be tiled back in.

Mouse
    * drag a split gap to resize neighbouring tiles;
    * drag a tile by its title bar onto another tile -- the left/right/top/
      bottom band docks it on that side, the centre swaps the two tiles;
    * double-click a title bar to toggle monocle.

Keyboard (``MOD`` defaults to Ctrl+Alt; Windows reserves the Super key)
    MOD+Arrows focus          MOD+Shift+Arrows move/swap
    MOD+= / MOD+-  grow/shrink   MOD+J toggle split   MOD+B balance
    MOD+F monocle  MOD+W close   MOD+T float/tile tool window
    MOD+1..9 workspace           MOD+Shift+1..9 send tile to workspace
    MOD+Space launcher           MOD+K key bindings   MOD+C classic pages

``TilingCanvas`` also exposes the small ``QStackedWidget`` API RapidPy used
before (``addWidget``/``setCurrentIndex``/``currentWidget``/``currentChanged``
...), so code that "switches pages" now opens or focuses the matching tile.
*Classic pages* mode shows one panel at a time for operators who prefer the
VB6-style single page.
"""

from __future__ import annotations

import json
import math
from dataclasses import dataclass, field
from typing import Callable, Iterable

import shiboken6
from PySide6 import QtCore, QtGui, QtWidgets

TILE_MIME = "application/x-rapidpy-tile"
DIRECTIONS = ("left", "right", "up", "down")
DEFAULT_MODIFIER = "Ctrl+Alt"


# ── layout tree ─────────────────────────────────────────────────────────────


@dataclass
class Leaf:
    key: str

    def to_dict(self) -> dict:
        return {"leaf": self.key}


@dataclass
class Split:
    orientation: str  # "h" = side by side, "v" = stacked
    children: list = field(default_factory=list)
    ratios: list[float] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "split": self.orientation,
            "children": [child.to_dict() for child in self.children],
            "ratios": [round(float(r), 5) for r in self.ratios],
        }


def node_from_dict(payload: object, known: Iterable[str]) -> Leaf | Split | None:
    """Rebuild a layout tree, dropping tiles that no longer exist."""

    known = set(known)
    if not isinstance(payload, dict):
        return None
    if "leaf" in payload:
        key = str(payload["leaf"])
        return Leaf(key) if key in known else None
    orientation = payload.get("split")
    if orientation not in ("h", "v"):
        return None
    children, ratios = [], []
    raw_ratios = payload.get("ratios") or []
    for index, child_payload in enumerate(payload.get("children") or []):
        child = node_from_dict(child_payload, known)
        if child is None:
            continue
        children.append(child)
        try:
            ratio = float(raw_ratios[index])
        except (IndexError, TypeError, ValueError):
            ratio = 1.0
        ratios.append(ratio if math.isfinite(ratio) and ratio > 0 else 1.0)
    if not children:
        return None
    if len(children) == 1:
        return children[0]
    total = sum(ratios)
    return Split(orientation, children, [r / total for r in ratios])


def leaves(node: Leaf | Split | None) -> list[str]:
    if node is None:
        return []
    if isinstance(node, Leaf):
        return [node.key]
    out: list[str] = []
    for child in node.children:
        out.extend(leaves(child))
    return out


def _find_parent(node, key: str, parent=None):
    """Return ``(parent_split, index)`` for the leaf ``key`` (parent None for root)."""

    if isinstance(node, Leaf):
        return (parent, None) if node.key == key else None
    for index, child in enumerate(node.children):
        if isinstance(child, Leaf) and child.key == key:
            return node, index
        found = _find_parent(child, key, node)
        if found is not None:
            return found
    return None


def remove_leaf(root, key: str):
    """Remove ``key``; a split left with one child is replaced by that child."""

    if root is None:
        return None
    if isinstance(root, Leaf):
        return None if root.key == key else root
    found = _find_parent(root, key)
    if found is None:
        return root
    parent, index = found
    del parent.children[index]
    del parent.ratios[index]
    total = sum(parent.ratios) or 1.0
    parent.ratios = [r / total for r in parent.ratios]
    return _collapse(root)


def _collapse(node):
    if isinstance(node, Leaf):
        return node
    node.children = [_collapse(child) for child in node.children]
    if len(node.children) == 1:
        return node.children[0]
    return node


def replace_leaf(root, key: str, replacement):
    if isinstance(root, Leaf):
        return replacement if root.key == key else root
    for index, child in enumerate(root.children):
        root.children[index] = replace_leaf(child, key, replacement)
    return root


def insert_beside(root, target: str | None, key: str, side: str):
    """Dock ``key`` on ``side`` (left/right/up/down) of ``target``."""

    orientation = "h" if side in ("left", "right") else "v"
    new_first = side in ("left", "up")
    if root is None:
        return Leaf(key)
    if target is None or target not in leaves(root):
        children = [Leaf(key), root] if new_first else [root, Leaf(key)]
        return Split(orientation, children, [0.5, 0.5])
    pair = [Leaf(key), Leaf(target)] if new_first else [Leaf(target), Leaf(key)]
    return replace_leaf(root, target, Split(orientation, pair, [0.5, 0.5]))


def swap_leaves(root, first: str, second: str):
    if isinstance(root, Leaf):
        if root.key == first:
            return Leaf(second)
        if root.key == second:
            return Leaf(first)
        return root
    root.children = [swap_leaves(child, first, second) for child in root.children]
    return root


# ── tile chrome ─────────────────────────────────────────────────────────────


class TileScrollArea(QtWidgets.QScrollArea):
    """Fit content to the tile width; scroll vertically when the tile is short.

    Panels already adapt their width (their minimum width is 0), so like the
    former page stack the tile hands them exactly its width.  Height is never
    squeezed below the panel's layout minimum -- the tile scrolls instead.
    """

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("tileScroll")
        self.setWidgetResizable(False)
        self.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)

    def setWidget(self, widget: QtWidgets.QWidget) -> None:  # type: ignore[override]
        super().setWidget(widget)
        widget.installEventFilter(self)
        self.fit_content()

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self.fit_content()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:  # type: ignore[override]
        if obj is self.widget() and event.type() == QtCore.QEvent.Type.LayoutRequest:
            QtCore.QTimer.singleShot(0, self.fit_content)
        return super().eventFilter(obj, event)

    def fit_content(self) -> None:
        if not shiboken6.isValid(self):
            return  # a deferred fit can outlive its tile
        widget = self.widget()
        if widget is None:
            return
        width = self.viewport().width()
        needed = widget.minimumSizeHint().height() if widget.layout() is not None else widget.minimumHeight()
        widget.resize(width, max(self.viewport().height(), needed))


class TrafficLight(QtWidgets.QAbstractButton):
    """macOS window control: a coloured dot that shows its glyph on header hover.

    Inactive tiles show grey dots, like background windows on macOS; a control
    that does not apply to this tile (e.g. float for a panel) stays grey and is
    disabled.
    """

    COLORS = {"close": ("#FF5F57", "#E0443E"), "float": ("#FEBC2E", "#DEA123"), "monocle": ("#28C840", "#1AAB29")}
    GLYPHS = {"close": "×", "float": "−", "monocle": "+"}

    def __init__(self, role: str, tip: str) -> None:
        super().__init__()
        self.role = role
        self.setToolTip(tip)
        self.setAccessibleName(tip)
        self.setFixedSize(14, 14)
        self.setCursor(QtCore.Qt.CursorShape.ArrowCursor)
        self.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        self.active = False
        self.reveal = False

    def sizeHint(self) -> QtCore.QSize:  # type: ignore[override]
        return QtCore.QSize(14, 14)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        rect = QtCore.QRectF(1, 1, 12, 12)
        coloured = self.isEnabled() and (self.active or self.reveal)
        fill, rim = self.COLORS[self.role] if coloured else ("#D7D2D3", "#C3BDBE")
        painter.setPen(QtGui.QPen(QtGui.QColor(rim), 0.8))
        painter.setBrush(QtGui.QColor(fill))
        painter.drawEllipse(rect)
        if coloured and self.reveal:
            painter.setPen(QtGui.QPen(QtGui.QColor(0, 0, 0, 150), 1.4))
            font = QtGui.QFont(self.font())
            font.setPixelSize(11)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect.adjusted(0, -1, 0, 0), QtCore.Qt.AlignmentFlag.AlignCenter, self.GLYPHS[self.role])
        painter.end()


class TileFrame(QtWidgets.QFrame):
    """One tile: a draggable title bar and a scrollable body.

    ``chrome="mac"`` draws macOS-style traffic lights (close / float /
    monocle) on the left with a centred title; ``"compact"`` keeps small
    glyph buttons on the right.
    """

    focus_requested = QtCore.Signal(str)
    close_requested = QtCore.Signal(str)
    monocle_requested = QtCore.Signal(str)
    float_requested = QtCore.Signal(str)

    def __init__(self, key: str, title: str, widget: QtWidgets.QWidget, *, icon: str = "", closable: bool = True,
                 floatable: bool = False, scroll: bool = True, chrome: str = "compact") -> None:
        super().__init__()
        self.key = key
        self.widget = widget
        self.chrome = chrome
        self.setObjectName("tile")
        self.setProperty("chrome", chrome)
        self.setProperty("active", False)
        self.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(1, 1, 1, 1)
        outer.setSpacing(0)

        self.header = QtWidgets.QWidget()
        self.header.setObjectName("tileHeader")
        self.header.setCursor(QtCore.Qt.CursorShape.OpenHandCursor)
        self.header.installEventFilter(self)
        head = QtWidgets.QHBoxLayout(self.header)
        self.title_label = QtWidgets.QLabel(f"{icon}  {title}".strip())
        self.title_label.setObjectName("tileTitle")
        self.title_label.installEventFilter(self)
        if chrome == "mac":
            head.setContentsMargins(12, 7, 12, 7)
            head.setSpacing(8)
            self.close_button = self._traffic("close", "Close tile (MOD+W)", self.close_requested)
            self.float_button = self._traffic("float", "Float / tile this window (MOD+T)", self.float_requested)
            self.monocle_button = self._traffic("monocle", "Monocle: fill the canvas (MOD+F)", self.monocle_requested)
            self.close_button.setEnabled(closable)
            self.float_button.setEnabled(floatable)
            lights = QtWidgets.QHBoxLayout()
            lights.setSpacing(8)
            for button in (self.close_button, self.float_button, self.monocle_button):
                lights.addWidget(button)
            head.addLayout(lights)
            self.title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            head.addWidget(self.title_label, 1)
            # Balance the traffic lights so the title stays visually centred.
            head.addSpacing(3 * 14 + 2 * 8)
            self.header.setAttribute(QtCore.Qt.WidgetAttribute.WA_Hover, True)
        else:
            head.setContentsMargins(12, 5, 6, 5)
            head.setSpacing(4)
            head.addWidget(self.title_label, 1)
            self.float_button = self._tool_button("⇱", "Float / tile this window (MOD+T)", self.float_requested)
            self.float_button.setVisible(floatable)
            self.monocle_button = self._tool_button("⤢", "Monocle: fill the canvas (MOD+F)", self.monocle_requested)
            self.close_button = self._tool_button("✕", "Close tile (MOD+W)", self.close_requested)
            self.close_button.setVisible(closable)
            for button in (self.float_button, self.monocle_button, self.close_button):
                head.addWidget(button)
        outer.addWidget(self.header)

        if scroll:
            self.body = TileScrollArea()
            self.body.setWidget(widget)
        else:
            self.body = widget
        outer.addWidget(self.body, 1)
        self._press_pos: QtCore.QPoint | None = None

    def _traffic(self, role: str, tip: str, signal) -> TrafficLight:
        button = TrafficLight(role, tip)
        button.clicked.connect(lambda: signal.emit(self.key))
        return button

    def _traffic_lights(self) -> list[TrafficLight]:
        return [b for b in (self.close_button, self.float_button, self.monocle_button) if isinstance(b, TrafficLight)]

    def _tool_button(self, text: str, tip: str, signal) -> QtWidgets.QToolButton:
        button = QtWidgets.QToolButton()
        button.setObjectName("tileButton")
        button.setText(text)
        button.setToolTip(tip)
        button.setAutoRaise(True)
        button.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        button.clicked.connect(lambda: signal.emit(self.key))
        return button

    def set_title(self, title: str) -> None:
        self.title_label.setText(title)

    def set_active(self, active: bool) -> None:
        if bool(self.property("active")) == bool(active):
            return
        self.setProperty("active", bool(active))
        for light in self._traffic_lights():
            light.active = bool(active)
            light.update()
        self.style().unpolish(self)
        self.style().polish(self)
        for child in (self.header, self.title_label):
            child.style().unpolish(child)
            child.style().polish(child)
        self.update()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        super().paintEvent(event)
        if not self.property("active"):
            return
        painter = QtGui.QPainter(self)
        if self.chrome == "mac":
            # macOS-style key window: a quiet accent ring, no loud gradient.
            painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
            painter.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25, 90), 1.5))
            painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QtCore.QRectF(self.rect()).adjusted(0.75, 0.75, -0.75, -0.75), 12, 12)
            painter.end()
            return
        # Omarchy-style gradient border on the focused tile.
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        gradient = QtGui.QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(0.0, QtGui.QColor(122, 2, 25, 235))
        gradient.setColorAt(1.0, QtGui.QColor(253, 181, 21, 235))
        pen = QtGui.QPen(QtGui.QBrush(gradient), 2.0)
        painter.setPen(pen)
        painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(QtCore.QRectF(self.rect()).adjusted(1, 1, -1, -1), 14, 14)
        painter.end()

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:  # type: ignore[override]
        kind = event.type()
        if obj is self.header and kind in (QtCore.QEvent.Type.HoverEnter, QtCore.QEvent.Type.HoverLeave):
            for light in self._traffic_lights():
                light.reveal = kind == QtCore.QEvent.Type.HoverEnter
                light.update()
            return False
        if kind == QtCore.QEvent.Type.MouseButtonPress and event.button() == QtCore.Qt.MouseButton.LeftButton:
            self._press_pos = event.position().toPoint()
            self.focus_requested.emit(self.key)
        elif kind == QtCore.QEvent.Type.MouseMove and self._press_pos is not None:
            if (event.position().toPoint() - self._press_pos).manhattanLength() >= QtWidgets.QApplication.startDragDistance():
                self._press_pos = None
                self._start_drag()
                return True
        elif kind == QtCore.QEvent.Type.MouseButtonRelease:
            self._press_pos = None
        elif kind == QtCore.QEvent.Type.MouseButtonDblClick:
            self.monocle_requested.emit(self.key)
            return True
        return super().eventFilter(obj, event)

    def _start_drag(self) -> None:
        drag = QtGui.QDrag(self)
        mime = QtCore.QMimeData()
        mime.setData(TILE_MIME, self.key.encode("utf-8"))
        drag.setMimeData(mime)
        preview = self.grab().scaled(
            260, 170, QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation
        )
        drag.setPixmap(preview)
        drag.setHotSpot(QtCore.QPoint(preview.width() // 2, 16))
        drag.exec(QtCore.Qt.DropAction.MoveAction)


class _DropPreview(QtWidgets.QWidget):
    """Translucent rectangle showing where a dragged tile will land."""

    def __init__(self, parent: QtWidgets.QWidget) -> None:
        super().__init__(parent)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.hide()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        painter.setPen(QtGui.QPen(QtGui.QColor(122, 2, 25, 200), 2, QtCore.Qt.PenStyle.DashLine))
        painter.setBrush(QtGui.QColor(253, 181, 21, 70))
        painter.drawRoundedRect(QtCore.QRectF(self.rect()).adjusted(2, 2, -2, -2), 14, 14)
        painter.end()


@dataclass
class TileSpec:
    key: str
    title: str
    widget: QtWidgets.QWidget
    icon: str = ""
    closable: bool = True
    kind: str = "panel"  # "panel" (persistent) or "window" (docked tool window)
    frame: TileFrame | None = None
    on_close: Callable[[], None] | None = None


# ── canvas ─────────────────────────────────────────────────────────────────


class TilingCanvas(QtWidgets.QWidget):
    """Tiling workspace host with a ``QStackedWidget``-compatible facade."""

    currentChanged = QtCore.Signal(int)
    focusChanged = QtCore.Signal(str)
    layoutChanged = QtCore.Signal()
    workspaceChanged = QtCore.Signal(int)
    floated = QtCore.Signal(str)

    def __init__(self, parent: QtWidgets.QWidget | None = None, *, workspaces: int = 5, gap: int = 8,
                 chrome: str = "compact") -> None:
        super().__init__(parent)
        self.setObjectName("tilingCanvas")
        self._chrome = chrome
        self.setAcceptDrops(True)
        self._gap = int(gap)
        self._workspace_count = max(1, int(workspaces))
        self._specs: dict[str, TileSpec] = {}
        self._order: list[str] = []  # registration order = stacked-widget index
        self._trees: dict[int, Leaf | Split | None] = {n: None for n in range(1, self._workspace_count + 1)}
        self._focus: dict[int, str | None] = {n: None for n in self._trees}
        self._monocle: set[int] = set()
        self._active = 1
        self._classic = False
        self._classic_key: str | None = None
        self._animation_ms = 140
        self._splitters: list[QtWidgets.QSplitter] = []
        self._split_nodes: dict[int, Split] = {}
        self._animations: list[QtCore.QVariantAnimation] = []
        self._launcher_actions: list[tuple[str, Callable[[], None]]] = []
        self._rebuild_pending = False

        self._holder = QtWidgets.QWidget(self)  # parks tiles that are not on screen
        self._holder.hide()
        self._layout = QtWidgets.QVBoxLayout(self)
        self._layout.setContentsMargins(self._gap, self._gap, self._gap, self._gap)
        self._layout.setSpacing(0)
        self._root_widget: QtWidgets.QWidget | None = None
        self._empty = QtWidgets.QLabel("Empty workspace — press Ctrl+Alt+Space to open a panel")
        self._empty.setObjectName("tilingEmpty")
        self._empty.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self._empty.setParent(self._holder)
        self._preview = _DropPreview(self)
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.focusChanged.connect(self._on_app_focus_changed)

    # ── registration ──────────────────────────────────────────────────
    def register(self, key: str, widget: QtWidgets.QWidget, title: str, *, icon: str = "", closable: bool = True,
                 kind: str = "panel", scroll: bool = True, on_close: Callable[[], None] | None = None) -> TileSpec:
        if key in self._specs:
            raise ValueError(f"tile {key!r} is already registered")
        frame = TileFrame(key, title, widget, icon=icon, closable=closable, floatable=(kind == "window"), scroll=scroll,
                          chrome=self._chrome)
        frame.setParent(self._holder)
        frame.focus_requested.connect(self.focus)
        frame.close_requested.connect(self.close_tile)
        frame.monocle_requested.connect(self._monocle_from_tile)
        frame.float_requested.connect(self.float_tile)
        spec = TileSpec(key, title, widget, icon, closable, kind, frame, on_close)
        self._specs[key] = spec
        if kind == "panel":
            self._order.append(key)
        return spec

    def addWidget(self, widget: QtWidgets.QWidget) -> int:  # QStackedWidget compatibility
        key = widget.objectName() or f"panel{len(self._order)}"
        while key in self._specs:
            key += "_"
        self.register(key, widget, widget.windowTitle() or key)
        return len(self._order) - 1

    def add_launcher_action(self, title: str, callback: Callable[[], None]) -> None:
        self._launcher_actions.append((title, callback))

    def set_animation_duration(self, ms: int) -> None:
        self._animation_ms = max(0, int(ms))

    # ── QStackedWidget facade ─────────────────────────────────────────
    def count(self) -> int:
        return len(self._order)

    def widget(self, index: int) -> QtWidgets.QWidget | None:
        if 0 <= index < len(self._order):
            return self._specs[self._order[index]].widget
        return None

    def indexOf(self, widget: QtWidgets.QWidget) -> int:
        for index, key in enumerate(self._order):
            if self._specs[key].widget is widget:
                return index
        return -1

    def currentIndex(self) -> int:
        key = self.focused_key()
        return self._order.index(key) if key in self._order else -1

    def currentWidget(self) -> QtWidgets.QWidget | None:
        key = self.focused_key()
        return self._specs[key].widget if key in self._specs else None

    def setCurrentIndex(self, index: int) -> None:
        if 0 <= index < len(self._order):
            self.open(self._order[index])

    def setCurrentWidget(self, widget: QtWidgets.QWidget) -> None:
        self.setCurrentIndex(self.indexOf(widget))

    def minimumSizeHint(self) -> QtCore.QSize:  # type: ignore[override]
        # Tiles scroll, so the canvas never forces the window wider; keep a
        # screen-capped vertical minimum like the former page stack.
        screen = self.screen() or QtWidgets.QApplication.primaryScreen()
        height = 240
        if screen is not None:
            height = min(height, screen.availableGeometry().height())
        return QtCore.QSize(0, height)

    def sizeHint(self) -> QtCore.QSize:  # type: ignore[override]
        return QtCore.QSize(1000, 700)

    # ── queries ───────────────────────────────────────────────────────
    @property
    def active_workspace(self) -> int:
        return self._active

    @property
    def workspace_count(self) -> int:
        return self._workspace_count

    @property
    def classic(self) -> bool:
        return self._classic

    def is_monocle(self, workspace: int | None = None) -> bool:
        return (workspace or self._active) in self._monocle

    def focused_key(self) -> str | None:
        if self._classic:
            return self._classic_key
        return self._focus.get(self._active)

    def workspace_of(self, key: str) -> int | None:
        for number, tree in self._trees.items():
            if key in leaves(tree):
                return number
        return None

    def visible_keys(self) -> list[str]:
        if self._classic:
            return [self._classic_key] if self._classic_key else []
        tree = self._trees[self._active]
        if self._active in self._monocle and self._focus[self._active]:
            return [self._focus[self._active]]
        return leaves(tree)

    def occupied_workspaces(self) -> list[int]:
        return [n for n, tree in self._trees.items() if tree is not None]

    def tree(self, workspace: int | None = None):
        return self._trees[workspace or self._active]

    def frame(self, key: str) -> TileFrame | None:
        spec = self._specs.get(key)
        return spec.frame if spec else None

    def spec(self, key: str) -> TileSpec | None:
        return self._specs.get(key)

    def registered_keys(self) -> list[str]:
        return list(self._specs)

    # ── tile operations ───────────────────────────────────────────────
    def open(self, key: str, *, workspace: int | None = None, beside: str | None = None, side: str | None = None) -> None:
        """Focus ``key`` where it is, or tile it into ``workspace`` (default: active)."""

        if key not in self._specs:
            raise KeyError(key)
        if self._classic:
            self._classic_key = key
            self._set_focus_key(key)
            self._schedule_rebuild()
            return
        where = self.workspace_of(key)
        if where is not None and workspace in (None, where):
            if where != self._active:
                self.switch_workspace(where)
            self.focus(key)
            return
        if where is not None:
            self._trees[where] = remove_leaf(self._trees[where], key)
            if self._focus[where] == key:
                remaining = leaves(self._trees[where])
                self._focus[where] = remaining[-1] if remaining else None
        target_ws = workspace or self._active
        tree = self._trees[target_ws]
        target = beside or self._focus.get(target_ws)
        if target not in leaves(tree):
            target = (leaves(tree) or [None])[-1]
        self._trees[target_ws] = insert_beside(tree, target, key, side or self._dwindle_side(target))
        self._focus[target_ws] = key
        self._monocle.discard(target_ws)
        if target_ws != self._active:
            self.switch_workspace(target_ws)
        self._set_focus_key(key)
        self._schedule_rebuild(grow=key)

    def _dwindle_side(self, target: str | None) -> str:
        frame = self.frame(target) if target else None
        if frame is not None and frame.isVisible() and frame.height() > frame.width() * 1.05:
            return "down"
        return "right"

    def close_tile(self, key: str) -> None:
        spec = self._specs.get(key)
        if spec is None:
            return
        if self._classic and spec.kind == "panel":
            return  # classic pages always show a panel
        if spec.kind == "window":
            # Ask the tool window itself to close while it is still visible, so
            # dialogs run their own cleanup (reject/finished, ownership leases).
            # A window that refuses (busy hardware work) keeps its tile.
            result = (spec.on_close or spec.widget.close)()
            if result is False:
                return
            self._drop_window(key)
            return
        where = self.workspace_of(key)
        if where is not None:
            self._trees[where] = remove_leaf(self._trees[where], key)
            if self._focus[where] == key:
                remaining = leaves(self._trees[where])
                self._focus[where] = remaining[-1] if remaining else None
                if not remaining:
                    self._monocle.discard(where)
        if spec.frame is not None:
            spec.frame.setParent(self._holder)
            spec.frame.hide()
        self._schedule_rebuild()
        if where == self._active:
            self._emit_focus()

    def focus(self, key: str) -> None:
        if key not in self._specs:
            return
        if not self._classic:
            where = self.workspace_of(key)
            if where is None:
                return
            if where != self._active:
                self.switch_workspace(where)
            self._focus[self._active] = key
            if self._active in self._monocle:
                self._schedule_rebuild()
        self._set_focus_key(key)

    def _set_focus_key(self, key: str | None) -> None:
        for name, spec in self._specs.items():
            if spec.frame is not None:
                spec.frame.set_active(name == key)
        self._sync_keyboard_focus()
        self._emit_focus()

    def _sync_keyboard_focus(self) -> None:
        """Put keyboard focus inside the focused tile unless it is already there."""
        frame = self.frame(self.focused_key()) if self.focused_key() else None
        if frame is None or not frame.isVisible():
            return
        current = QtWidgets.QApplication.focusWidget()
        if current is not None and (current is frame or frame.isAncestorOf(current)):
            return
        if current is not None and not self.isAncestorOf(current):
            return  # focus is outside the canvas (header, sidebar, dialog): leave it
        frame.body.setFocus(QtCore.Qt.FocusReason.OtherFocusReason)

    def _emit_focus(self) -> None:
        key = self.focused_key()
        self.focusChanged.emit(key or "")
        self.currentChanged.emit(self._order.index(key) if key in self._order else -1)

    def switch_workspace(self, number: int) -> None:
        if not 1 <= number <= self._workspace_count or number == self._active:
            return
        self._active = number
        if self._classic:
            self._classic = False
        self._schedule_rebuild()
        self._set_focus_key(self._focus.get(number))
        self.workspaceChanged.emit(number)

    def move_to_workspace(self, number: int, key: str | None = None) -> None:
        key = key or self.focused_key()
        if key is None or not 1 <= number <= self._workspace_count:
            return
        source = self.workspace_of(key)
        if source == number:
            return
        if source is not None:
            self._trees[source] = remove_leaf(self._trees[source], key)
            if self._focus[source] == key:
                remaining = leaves(self._trees[source])
                self._focus[source] = remaining[-1] if remaining else None
        tree = self._trees[number]
        target = self._focus.get(number) if self._focus.get(number) in leaves(tree) else (leaves(tree) or [None])[-1]
        self._trees[number] = insert_beside(tree, target, key, "right")
        self._focus[number] = key
        self._schedule_rebuild()
        self._set_focus_key(self.focused_key())
        self.workspaceChanged.emit(self._active)

    def toggle_monocle(self) -> None:
        if self._classic:
            return
        if self._active in self._monocle:
            self._monocle.discard(self._active)
        elif self._focus.get(self._active):
            self._monocle.add(self._active)
        self._schedule_rebuild()

    def _monocle_from_tile(self, key: str) -> None:
        self.focus(key)
        self.toggle_monocle()

    def toggle_split(self) -> None:
        key = self.focused_key()
        found = _find_parent(self._trees[self._active], key) if key else None
        if not found or found[0] is None:
            return
        parent = found[0]
        parent.orientation = "v" if parent.orientation == "h" else "h"
        self._schedule_rebuild()

    def balance(self) -> None:
        def even(node):
            if isinstance(node, Split):
                node.ratios = [1.0 / len(node.children)] * len(node.children)
                for child in node.children:
                    even(child)

        even(self._trees[self._active])
        self._apply_ratios(animated=True)

    def resize_focused(self, delta: float) -> None:
        """Grow (``delta`` > 0) or shrink the focused tile within its split."""

        key = self.focused_key()
        found = _find_parent(self._trees[self._active], key) if key else None
        if not found or found[0] is None:
            return
        parent, index = found
        count = len(parent.ratios)
        new = min(0.9, max(0.1, parent.ratios[index] + float(delta)))
        others = sum(r for i, r in enumerate(parent.ratios) if i != index) or 1.0
        scale = (1.0 - new) / others
        parent.ratios = [new if i == index else r * scale for i, r in enumerate(parent.ratios)]
        if count:
            self._apply_ratios(animated=True)

    def neighbour(self, direction: str, key: str | None = None) -> str | None:
        key = key or self.focused_key()
        frame = self.frame(key) if key else None
        if frame is None or not frame.isVisible():
            return None
        origin = self._rect_in_canvas(frame)
        best, best_score = None, math.inf
        for other in self.visible_keys():
            if other == key:
                continue
            other_frame = self.frame(other)
            if other_frame is None or not other_frame.isVisible():
                continue
            rect = self._rect_in_canvas(other_frame)
            if direction == "left" and rect.right() <= origin.left() + 2:
                gap, overlap = origin.left() - rect.right(), _overlap(origin.top(), origin.bottom(), rect.top(), rect.bottom())
            elif direction == "right" and rect.left() >= origin.right() - 2:
                gap, overlap = rect.left() - origin.right(), _overlap(origin.top(), origin.bottom(), rect.top(), rect.bottom())
            elif direction == "up" and rect.bottom() <= origin.top() + 2:
                gap, overlap = origin.top() - rect.bottom(), _overlap(origin.left(), origin.right(), rect.left(), rect.right())
            elif direction == "down" and rect.top() >= origin.bottom() - 2:
                gap, overlap = rect.top() - origin.bottom(), _overlap(origin.left(), origin.right(), rect.left(), rect.right())
            else:
                continue
            # Prefer tiles that share an edge span, then the nearest.
            score = gap + (0 if overlap > 0 else 10_000) - overlap * 0.01
            if score < best_score:
                best, best_score = other, score
        return best

    def focus_direction(self, direction: str) -> None:
        target = self.neighbour(direction)
        if target:
            self.focus(target)

    def swap_direction(self, direction: str) -> None:
        key = self.focused_key()
        target = self.neighbour(direction)
        if key and target:
            self.swap(key, target)

    def swap(self, first: str, second: str) -> None:
        ws_first, ws_second = self.workspace_of(first), self.workspace_of(second)
        if ws_first is None or ws_second is None:
            return
        if ws_first == ws_second:
            self._trees[ws_first] = swap_leaves(self._trees[ws_first], first, second)
        else:
            self._trees[ws_first] = replace_leaf(self._trees[ws_first], first, Leaf(second))
            self._trees[ws_second] = replace_leaf(self._trees[ws_second], second, Leaf(first))
        self._schedule_rebuild()
        self._set_focus_key(self.focused_key())

    def dock(self, key: str, target: str, zone: str) -> None:
        """Drop handling: ``zone`` is a side to dock on, or ``"center"`` to swap."""

        if key == target or target not in self._specs:
            return
        if zone == "center":
            if self.workspace_of(key) is None:
                self.open(key, beside=target)
            else:
                self.swap(key, target)
            self.focus(key)
            return
        target_ws = self.workspace_of(target) or self._active
        source_ws = self.workspace_of(key)
        if source_ws is not None:
            self._trees[source_ws] = remove_leaf(self._trees[source_ws], key)
            if self._focus[source_ws] == key:
                remaining = leaves(self._trees[source_ws])
                self._focus[source_ws] = remaining[-1] if remaining else None
        self._trees[target_ws] = insert_beside(self._trees[target_ws], target, key, zone)
        self._focus[target_ws] = key
        self._monocle.discard(target_ws)
        self._schedule_rebuild(grow=key)
        self._set_focus_key(key)

    def set_classic(self, enabled: bool) -> None:
        enabled = bool(enabled)
        if enabled == self._classic:
            return
        if enabled:
            self._classic_key = self.focused_key() if self.focused_key() in self._order else (self._order[0] if self._order else None)
        self._classic = enabled
        if not enabled and self._classic_key and self.workspace_of(self._classic_key) is None:
            key, self._classic_key = self._classic_key, None
            self.open(key)
            return
        if not enabled and self._classic_key:
            self.focus(self._classic_key)
        self._schedule_rebuild()
        self._emit_focus()

    # ── tool windows ──────────────────────────────────────────────────
    def dock_window(self, widget: QtWidgets.QWidget, key: str, title: str, *, icon: str = "",
                    on_close: Callable[[], None] | None = None, side: str | None = None) -> None:
        """Tile a top-level tool window (e.g. a non-modal dialog) into the canvas."""

        if key in self._specs:
            self.open(key)
            return
        widget.setWindowFlags(QtCore.Qt.WindowType.Widget)
        spec = self.register(key, widget, title, icon=icon, kind="window", on_close=on_close)
        widget.destroyed.connect(lambda *_: self._window_destroyed(key))
        finished = getattr(widget, "finished", None)
        if finished is not None:
            finished.connect(lambda *_: self._window_finished(key))
        widget.show()
        self.open(spec.key, side=side)

    def _window_finished(self, key: str) -> None:
        if key in self._specs:
            QtCore.QTimer.singleShot(0, lambda: self._drop_window(key))

    def _window_destroyed(self, key: str) -> None:
        if not shiboken6.isValid(self):
            return  # the canvas itself is being torn down
        self._drop_window(key, destroyed=True)

    def _drop_window(self, key: str, *, destroyed: bool = False) -> None:
        spec = self._specs.get(key)
        if spec is None or spec.kind != "window":
            return
        where = self.workspace_of(key)
        if where is not None:
            self._trees[where] = remove_leaf(self._trees[where], key)
            if self._focus[where] == key:
                remaining = leaves(self._trees[where])
                self._focus[where] = remaining[-1] if remaining else None
        self._forget_window(key, destroyed=destroyed)
        self._schedule_rebuild()
        self._emit_focus()

    def _forget_window(self, key: str, *, destroyed: bool = False) -> None:
        spec = self._specs.pop(key, None)
        if spec is None or spec.frame is None:
            return
        frame = spec.frame
        if not shiboken6.isValid(frame):
            return
        if not destroyed and isinstance(frame.body, QtWidgets.QScrollArea):
            try:
                widget = frame.body.takeWidget()
                if widget is not None:
                    widget.setParent(self._holder)
            except RuntimeError:
                pass
        frame.setParent(None)
        frame.deleteLater()

    def float_tile(self, key: str | None = None) -> None:
        """Pop a docked tool window out of the grid (Omarchy MOD+T)."""

        key = key or self.focused_key()
        spec = self._specs.get(key) if key else None
        if spec is None or spec.kind != "window":
            return
        widget = spec.widget
        where = self.workspace_of(key)
        if where is not None:
            self._trees[where] = remove_leaf(self._trees[where], key)
            if self._focus[where] == key:
                remaining = leaves(self._trees[where])
                self._focus[where] = remaining[-1] if remaining else None
        frame = spec.frame
        if frame is not None and isinstance(frame.body, QtWidgets.QScrollArea):
            frame.body.takeWidget()
        self._specs.pop(key, None)
        if frame is not None:
            frame.setParent(None)
            frame.deleteLater()
        geometry = frame.geometry() if frame is not None else QtCore.QRect(0, 0, 640, 480)
        widget.setParent(self.window(), QtCore.Qt.WindowType.Dialog)
        widget.resize(max(420, geometry.width()), max(320, geometry.height()))
        widget.show()
        widget.raise_()
        self._schedule_rebuild()
        self._emit_focus()
        self.floated.emit(key)

    # ── persistence ───────────────────────────────────────────────────
    def save_state(self) -> dict:
        panel_keys = {key for key, spec in self._specs.items() if spec.kind == "panel"}

        def panels_only(node):
            if node is None:
                return None
            if isinstance(node, Leaf):
                return node if node.key in panel_keys else None
            kept = [(panels_only(child), ratio) for child, ratio in zip(node.children, node.ratios)]
            kept = [(child, ratio) for child, ratio in kept if child is not None]
            if not kept:
                return None
            if len(kept) == 1:
                return kept[0][0]
            return Split(node.orientation, [c for c, _ in kept], [r for _, r in kept])

        return {
            "version": 1,
            "active": self._active,
            "classic": self._classic,
            "classic_key": self._classic_key,
            "monocle": sorted(self._monocle),
            "focus": {str(n): k for n, k in self._focus.items() if k in panel_keys},
            "workspaces": {
                str(n): (panels_only(tree).to_dict() if panels_only(tree) is not None else None)
                for n, tree in self._trees.items()
            },
        }

    def restore_state(self, state: object) -> bool:
        if isinstance(state, str):
            try:
                state = json.loads(state)
            except ValueError:
                return False
        if not isinstance(state, dict) or state.get("version") != 1:
            return False
        known = [key for key, spec in self._specs.items() if spec.kind == "panel"]
        trees = {n: None for n in range(1, self._workspace_count + 1)}
        seen: set[str] = set()
        for name, payload in (state.get("workspaces") or {}).items():
            try:
                number = int(name)
            except (TypeError, ValueError):
                continue
            if number not in trees:
                continue
            node = node_from_dict(payload, set(known) - seen)
            trees[number] = node
            seen.update(leaves(node))
        if not seen:
            return False
        self._trees = trees
        focus = state.get("focus") or {}
        self._focus = {n: (focus.get(str(n)) if focus.get(str(n)) in leaves(trees[n]) else (leaves(trees[n]) or [None])[-1]) for n in trees}
        self._monocle = {int(n) for n in state.get("monocle") or [] if int(n) in trees and trees[int(n)] is not None}
        active = int(state.get("active") or 1)
        self._active = active if active in trees else 1
        self._classic = bool(state.get("classic"))
        classic_key = state.get("classic_key")
        self._classic_key = classic_key if classic_key in known else (known[0] if known else None)
        self._schedule_rebuild()
        self._set_focus_key(self.focused_key())
        self.workspaceChanged.emit(self._active)
        return True

    def set_layout(self, workspaces: dict[int, object], *, active: int = 1, focus: dict[int, str] | None = None) -> None:
        """Install a layout given as ``{workspace: tree-dict}`` (used for defaults)."""

        state = {
            "version": 1,
            "active": active,
            "workspaces": {str(n): tree for n, tree in workspaces.items()},
            "focus": {str(n): k for n, k in (focus or {}).items()},
        }
        self._classic = False
        self.restore_state(state)

    # ── rendering ─────────────────────────────────────────────────────
    def _schedule_rebuild(self, grow: str | None = None) -> None:
        self._grow_key = grow
        self._rebuild()

    def _park_all(self) -> None:
        for spec in self._specs.values():
            if spec.frame is not None and spec.frame.parent() is not self._holder:
                spec.frame.setParent(self._holder)
                spec.frame.hide()
        if self._empty.parent() is not self._holder:
            self._empty.setParent(self._holder)

    def _rebuild(self) -> None:
        # Reparenting tiles makes Qt move keyboard focus around; that churn must
        # not be mistaken for the operator focusing a different tile.
        self._rebuilding = True
        try:
            self._rebuild_tree()
        finally:
            self._rebuilding = False
        self._sync_keyboard_focus()

    def _rebuild_tree(self) -> None:
        old_root = self._root_widget
        if old_root is not None:
            self._layout.removeWidget(old_root)
        self._park_all()
        for animation in self._animations:
            animation.stop()
        self._animations = []
        for splitter in self._splitters:
            splitter.setParent(None)
            splitter.deleteLater()
        self._splitters = []
        self._split_nodes = {}

        if self._classic:
            key = self._classic_key
            root = self._specs[key].frame if key in self._specs else self._empty
        elif self._active in self._monocle and self._focus.get(self._active) in self._specs:
            root = self._specs[self._focus[self._active]].frame
        else:
            tree = self._trees[self._active]
            root = self._build_node(tree) if tree is not None else self._empty
        root.setParent(self)
        self._layout.addWidget(root)
        for splitter in self._splitters:
            splitter.show()
        for key in self.visible_keys():
            frame = self.frame(key)
            if frame is not None:
                frame.show()
        root.show()
        self._root_widget = root
        self._preview.raise_()
        grow = getattr(self, "_grow_key", None)
        self._grow_key = None
        if grow is not None:
            self._animate_entry(grow)
        self.layoutChanged.emit()

    def _build_node(self, node) -> QtWidgets.QWidget:
        if isinstance(node, Leaf):
            return self._specs[node.key].frame
        orientation = QtCore.Qt.Orientation.Horizontal if node.orientation == "h" else QtCore.Qt.Orientation.Vertical
        splitter = QtWidgets.QSplitter(orientation)
        splitter.setObjectName("tilingSplit")
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(self._gap)
        splitter.setOpaqueResize(True)
        for child in node.children:
            splitter.addWidget(self._build_node(child))
        splitter.setSizes([max(1, int(r * 10000)) for r in node.ratios])
        splitter.splitterMoved.connect(lambda *_args, s=splitter, n=node: self._capture_ratios(s, n))
        self._splitters.append(splitter)
        self._split_nodes[id(splitter)] = node
        return splitter

    def _capture_ratios(self, splitter: QtWidgets.QSplitter, node: Split) -> None:
        sizes = splitter.sizes()
        total = float(sum(sizes)) or 1.0
        node.ratios = [size / total for size in sizes]

    def _apply_ratios(self, *, animated: bool) -> None:
        for splitter in self._splitters:
            node = self._split_nodes.get(id(splitter))
            if node is None:
                continue
            total = sum(splitter.sizes()) or 10000
            target = [max(1, int(r * total)) for r in node.ratios]
            self._animate_sizes(splitter, target if animated else None, final=target)

    def _animate_sizes(self, splitter: QtWidgets.QSplitter, target: list[int] | None, *, final: list[int],
                       start: list[int] | None = None) -> None:
        if target is None or self._animation_ms <= 0 or not splitter.isVisible():
            splitter.setSizes(final)
            return
        begin = start or splitter.sizes()
        animation = QtCore.QVariantAnimation(self)
        animation.setDuration(self._animation_ms)
        animation.setEasingCurve(QtCore.QEasingCurve.Type.OutCubic)
        animation.setStartValue(0.0)
        animation.setEndValue(1.0)

        def step(value, s=splitter, a=begin, b=target):
            try:
                s.setSizes([int(x + (y - x) * float(value)) for x, y in zip(a, b)])
            except RuntimeError:
                pass

        animation.valueChanged.connect(step)
        animation.start()
        self._animations.append(animation)

    def _animate_entry(self, key: str) -> None:
        """Grow a newly tiled frame out of its neighbour instead of popping in."""

        frame = self.frame(key)
        if frame is None or self._animation_ms <= 0:
            return
        splitter = frame.parentWidget()
        if not isinstance(splitter, QtWidgets.QSplitter):
            return
        index = splitter.indexOf(frame)
        final = splitter.sizes()
        total = sum(final) or 10000
        start = [total - 1 if i != index else 1 for i in range(len(final))]
        if len(final) > 2:
            start = list(final)
            start[index] = 1
        self._animate_sizes(splitter, final, final=final, start=start)

    def _rect_in_canvas(self, widget: QtWidgets.QWidget) -> QtCore.QRect:
        top_left = widget.mapTo(self, QtCore.QPoint(0, 0))
        return QtCore.QRect(top_left, widget.size())

    # ── focus tracking ────────────────────────────────────────────────
    def _on_app_focus_changed(self, _old: QtWidgets.QWidget | None, new: QtWidgets.QWidget | None) -> None:
        if getattr(self, "_rebuilding", False) or not shiboken6.isValid(self):
            return
        widget = new
        while widget is not None:
            if isinstance(widget, TileFrame) and widget.key in self._specs:
                if widget.key != self.focused_key() and widget.isVisible():
                    if self._classic:
                        return
                    self._focus[self._active] = widget.key
                    self._set_focus_key(widget.key)
                return
            widget = widget.parentWidget()

    # ── drag and drop ─────────────────────────────────────────────────
    def _drop_target(self, pos: QtCore.QPoint) -> tuple[str | None, str, QtCore.QRect]:
        for key in self.visible_keys():
            frame = self.frame(key)
            if frame is None or not frame.isVisible():
                continue
            rect = self._rect_in_canvas(frame)
            if not rect.contains(pos):
                continue
            fx = (pos.x() - rect.left()) / max(1, rect.width())
            fy = (pos.y() - rect.top()) / max(1, rect.height())
            edges = {"left": fx, "right": 1 - fx, "up": fy, "down": 1 - fy}
            zone, distance = min(edges.items(), key=lambda item: item[1])
            if distance > 0.3:
                return key, "center", rect.adjusted(rect.width() // 6, rect.height() // 6, -rect.width() // 6, -rect.height() // 6)
            half_w, half_h = rect.width() // 2, rect.height() // 2
            preview = {
                "left": QtCore.QRect(rect.left(), rect.top(), half_w, rect.height()),
                "right": QtCore.QRect(rect.left() + half_w, rect.top(), rect.width() - half_w, rect.height()),
                "up": QtCore.QRect(rect.left(), rect.top(), rect.width(), half_h),
                "down": QtCore.QRect(rect.left(), rect.top() + half_h, rect.width(), rect.height() - half_h),
            }[zone]
            return key, zone, preview
        return None, "", QtCore.QRect()

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:  # type: ignore[override]
        if event.mimeData().hasFormat(TILE_MIME):
            event.acceptProposedAction()

    def dragMoveEvent(self, event: QtGui.QDragMoveEvent) -> None:  # type: ignore[override]
        if not event.mimeData().hasFormat(TILE_MIME):
            return
        key = bytes(event.mimeData().data(TILE_MIME)).decode("utf-8")
        target, _zone, rect = self._drop_target(event.position().toPoint())
        if target and target != key:
            self._preview.setGeometry(rect)
            self._preview.show()
            self._preview.raise_()
        else:
            self._preview.hide()
        event.acceptProposedAction()

    def dragLeaveEvent(self, event: QtGui.QDragLeaveEvent) -> None:  # type: ignore[override]
        self._preview.hide()

    def dropEvent(self, event: QtGui.QDropEvent) -> None:  # type: ignore[override]
        self._preview.hide()
        if not event.mimeData().hasFormat(TILE_MIME):
            return
        key = bytes(event.mimeData().data(TILE_MIME)).decode("utf-8")
        target, zone, _rect = self._drop_target(event.position().toPoint())
        if target and target != key and key in self._specs:
            event.acceptProposedAction()
            QtCore.QTimer.singleShot(0, lambda: self.dock(key, target, zone))

    # ── launcher / help ───────────────────────────────────────────────
    def launcher_entries(self) -> list[tuple[str, Callable[[], None]]]:
        entries: list[tuple[str, Callable[[], None]]] = []
        for key in self._order:
            spec = self._specs[key]
            where = self.workspace_of(key)
            suffix = f"  · workspace {where}" if where else "  · open as new tile"
            entries.append((f"{spec.icon}  {spec.title}{suffix}".strip(), lambda k=key: self.open(k)))
        for key, spec in self._specs.items():
            if spec.kind == "window":
                entries.append((f"{spec.icon}  {spec.title}  · tool window".strip(), lambda k=key: self.focus(k)))
        entries.extend(self._launcher_actions)
        return entries

    def show_launcher(self) -> "TileLauncher":
        launcher = TileLauncher(self.window(), self.launcher_entries())
        launcher.open()
        return launcher

    def show_keybindings(self, modifier: str = DEFAULT_MODIFIER) -> QtWidgets.QDialog:
        dialog = KeybindingHelp(self.window(), modifier)
        dialog.open()
        return dialog


def _overlap(a0: int, a1: int, b0: int, b1: int) -> int:
    return max(0, min(a1, b1) - max(a0, b0))


class TileLauncher(QtWidgets.QDialog):
    """Omarchy-style launcher: type to filter, Enter to open."""

    def __init__(self, parent: QtWidgets.QWidget | None, entries: list[tuple[str, Callable[[], None]]]) -> None:
        super().__init__(parent, QtCore.Qt.WindowType.Popup | QtCore.Qt.WindowType.FramelessWindowHint)
        self.setObjectName("tileLauncher")
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self._entries = entries
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)
        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText("Open panel or tool…")
        self.list = QtWidgets.QListWidget()
        self.list.setObjectName("tileLauncherList")
        layout.addWidget(self.search)
        layout.addWidget(self.list, 1)
        self.search.textChanged.connect(self._filter)
        self.search.returnPressed.connect(self._activate)
        self.list.itemActivated.connect(lambda _item: self._activate())
        self.search.installEventFilter(self)
        self._filter("")
        self.resize(460, 380)
        if parent is not None:
            center = parent.geometry().center()
            self.move(center.x() - self.width() // 2, center.y() - self.height() // 2)
        self.search.setFocus()

    def _filter(self, text: str) -> None:
        self.list.clear()
        needle = text.strip().lower()
        for index, (title, _callback) in enumerate(self._entries):
            if all(part in title.lower() for part in needle.split()):
                item = QtWidgets.QListWidgetItem(title)
                item.setData(QtCore.Qt.ItemDataRole.UserRole, index)
                self.list.addItem(item)
        if self.list.count():
            self.list.setCurrentRow(0)

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:  # type: ignore[override]
        if obj is self.search and event.type() == QtCore.QEvent.Type.KeyPress:
            if event.key() in (QtCore.Qt.Key.Key_Down, QtCore.Qt.Key.Key_Up):
                row = self.list.currentRow() + (1 if event.key() == QtCore.Qt.Key.Key_Down else -1)
                self.list.setCurrentRow(max(0, min(self.list.count() - 1, row)))
                return True
        return super().eventFilter(obj, event)

    def _activate(self) -> None:
        item = self.list.currentItem()
        if item is None:
            return
        _title, callback = self._entries[int(item.data(QtCore.Qt.ItemDataRole.UserRole))]
        self.accept()
        callback()


KEYBINDINGS = (
    ("Arrows", "Focus the tile in that direction"),
    ("Shift+Arrows", "Swap the focused tile with its neighbour"),
    ("= / -", "Grow / shrink the focused tile"),
    ("J", "Toggle the split direction (side by side / stacked)"),
    ("B", "Balance all splits on this workspace"),
    ("F", "Monocle: focused tile fills the canvas"),
    ("W", "Close the focused tile"),
    ("T", "Float a tool window out of the grid / tile it back"),
    ("1 … 9", "Switch workspace"),
    ("Shift+1 … 9", "Send the focused tile to a workspace"),
    ("Space", "Launcher: open a panel or tool"),
    ("C", "Classic pages (one panel at a time) on/off"),
    ("K", "Show these key bindings"),
)


class KeybindingHelp(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget | None, modifier: str) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("Tiling key bindings")
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_DeleteOnClose, True)
        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("Tiling canvas")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        grid = QtWidgets.QGridLayout()
        for row, (keys, action) in enumerate(KEYBINDINGS):
            key_label = QtWidgets.QLabel(f"{modifier}+{keys}")
            key_label.setObjectName("metaKey")
            grid.addWidget(key_label, row, 0)
            grid.addWidget(QtWidgets.QLabel(action), row, 1)
        layout.addLayout(grid)
        mouse = QtWidgets.QLabel(
            "Mouse: drag a gap to resize · drag a title bar onto another tile (edges dock, centre swaps) · "
            "double-click a title bar for monocle."
        )
        mouse.setObjectName("guidanceText")
        mouse.setWordWrap(True)
        layout.addWidget(mouse)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class WorkspaceBar(QtWidgets.QWidget):
    """Waybar-style workspace switcher: active, occupied and empty workspaces."""

    def __init__(self, canvas: TilingCanvas, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("workspaceBar")
        self._canvas = canvas
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(8, 2, 8, 2)
        layout.setSpacing(4)
        self._buttons: list[QtWidgets.QPushButton] = []
        for number in range(1, canvas.workspace_count + 1):
            button = QtWidgets.QPushButton(str(number))
            button.setObjectName("workspaceButton")
            button.setCheckable(True)
            button.setFixedSize(30, 26)
            button.setToolTip(f"Workspace {number} ({DEFAULT_MODIFIER}+{number})")
            button.setAccessibleName(f"Workspace {number}")
            button.clicked.connect(lambda _checked, n=number: canvas.switch_workspace(n))
            layout.addWidget(button)
            self._buttons.append(button)
        layout.addStretch(1)
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Fixed)
        self.setFixedHeight(34)
        canvas.workspaceChanged.connect(self.refresh)
        canvas.layoutChanged.connect(self.refresh)
        self.refresh()

    def refresh(self, *_args) -> None:
        occupied = set(self._canvas.occupied_workspaces())
        for number, button in enumerate(self._buttons, start=1):
            button.setChecked(number == self._canvas.active_workspace and not self._canvas.classic)
            button.setProperty("occupied", number in occupied)
            button.style().unpolish(button)
            button.style().polish(button)


def install_tiling_shortcuts(window: QtWidgets.QWidget, canvas: TilingCanvas, modifier: str = DEFAULT_MODIFIER) -> list[QtGui.QShortcut]:
    """Bind the Omarchy-style keys on ``window`` (see module docstring)."""

    shortcuts: list[QtGui.QShortcut] = []

    def bind(keys: str, callback: Callable[[], None]) -> None:
        shortcut = QtGui.QShortcut(QtGui.QKeySequence(f"{modifier}+{keys}"), window)
        shortcut.setContext(QtCore.Qt.ShortcutContext.WindowShortcut)
        shortcut.activated.connect(callback)
        shortcuts.append(shortcut)

    arrows = {"Left": "left", "Right": "right", "Up": "up", "Down": "down"}
    for key, direction in arrows.items():
        bind(key, lambda d=direction: canvas.focus_direction(d))
        bind(f"Shift+{key}", lambda d=direction: canvas.swap_direction(d))
    bind("=", lambda: canvas.resize_focused(0.06))
    bind("+", lambda: canvas.resize_focused(0.06))
    bind("-", lambda: canvas.resize_focused(-0.06))
    bind("J", canvas.toggle_split)
    bind("B", canvas.balance)
    bind("F", canvas.toggle_monocle)
    bind("W", lambda: canvas.close_tile(canvas.focused_key()) if canvas.focused_key() else None)
    bind("T", canvas.float_tile)
    bind("Space", canvas.show_launcher)
    bind("K", lambda: canvas.show_keybindings(modifier))
    bind("C", lambda: canvas.set_classic(not canvas.classic))
    for number in range(1, min(9, canvas.workspace_count) + 1):
        bind(str(number), lambda n=number: canvas.switch_workspace(n))
        shifted = "!@#$%^&*("[number - 1]
        bind(f"Shift+{number}", lambda n=number: canvas.move_to_workspace(n))
        bind(f"Shift+{shifted}", lambda n=number: canvas.move_to_workspace(n))
    return shortcuts


TILING_QSS = """
QWidget#tilingCanvas { background: transparent; }
QSplitter#tilingSplit { background: transparent; }
QSplitter#tilingSplit::handle { background: transparent; }
QSplitter#tilingSplit::handle:hover { background: rgba(122, 2, 25, 40); border-radius: 3px; }
QFrame#tile {
    background: rgba(255, 255, 255, 196);
    border: 1px solid rgba(255, 255, 255, 230);
    border-bottom-color: rgba(122, 2, 25, 40);
    border-radius: 14px;
}
QFrame#tile[active="true"] { background: rgba(255, 255, 255, 222); }
QFrame#tile QWidget#tileHeader {
    background: rgba(255, 255, 255, 120);
    border: none;
    border-bottom: 1px solid rgba(122, 2, 25, 28);
    border-top-left-radius: 14px;
    border-top-right-radius: 14px;
}
QFrame#tile[active="true"] QWidget#tileHeader { background: rgba(253, 181, 21, 46); }
QFrame#tile QLabel#tileTitle { color: #5F5154; font-weight: 650; background: transparent; }
QFrame#tile[active="true"] QLabel#tileTitle { color: #7A0219; font-weight: 760; }
QFrame#tile QToolButton#tileButton {
    background: transparent; border: none; border-radius: 7px;
    color: #6F6265; padding: 1px 6px; min-width: 18px;
}
QFrame#tile QToolButton#tileButton:hover { background: rgba(122, 2, 25, 30); color: #7A0219; }
QFrame#tile QScrollArea#tileScroll,
QFrame#tile QScrollArea#tileScroll > QWidget > QWidget { background: transparent; border: none; }
QLabel#tilingEmpty { color: #8a7b7e; font-size: 13px; background: transparent; }
QWidget#workspaceBar { background: transparent; }
QPushButton#workspaceButton {
    min-height: 0; padding: 0; border-radius: 9px;
    background: rgba(255, 255, 255, 110); border: 1px solid rgba(122, 2, 25, 40);
    color: #8a7b7e; font-weight: 650;
}
QPushButton#workspaceButton[occupied="true"] { color: #493B3E; background: rgba(255, 255, 255, 190); }
QPushButton#workspaceButton:checked {
    color: white; border-color: rgba(255, 255, 255, 120);
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7A0219, stop:1 #4C0010);
}
QDialog#tileLauncher {
    background: rgba(250, 245, 240, 248);
    border: 1px solid rgba(122, 2, 25, 90);
    border-radius: 16px;
}
QListWidget#tileLauncherList { background: transparent; border: none; font-size: 13px; }
QListWidget#tileLauncherList::item { padding: 7px 10px; border-radius: 9px; }
QListWidget#tileLauncherList::item:selected { background: rgba(253, 181, 21, 110); color: #261E21; }
"""


def apply_tiling_theme(app: QtWidgets.QApplication) -> None:
    if TILING_QSS not in app.styleSheet():
        app.setStyleSheet(app.styleSheet() + TILING_QSS)
