from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import json
from pathlib import Path
import shutil
import sys

from PySide6 import QtCore, QtGui, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, apply_window_bounds_guard, set_app_icon  # noqa: E402


REPO_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_INI_PATH = REPO_ROOT / "VB6" / "Defaults.ini"


@dataclass
class IniEntry:
    key: str
    value: str


@dataclass
class IniSection:
    name: str
    entries: list[IniEntry] = field(default_factory=list)


@dataclass
class IniDocument:
    sections: list[IniSection] = field(default_factory=list)


def load_ini_document(path: Path) -> IniDocument:
    sections: list[IniSection] = []
    current: IniSection | None = None
    for raw_line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        stripped = raw_line.strip()
        if not stripped or stripped.startswith((";", "#")):
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            current = IniSection(name=stripped[1:-1].strip())
            sections.append(current)
            continue
        if current is None or "=" not in raw_line:
            continue
        key, value = raw_line.split("=", 1)
        current.entries.append(IniEntry(key=key.strip(), value=value.strip()))
    return IniDocument(sections=sections)


def save_ini_document(document: IniDocument, path: Path) -> None:
    lines: list[str] = []
    for index, section in enumerate(document.sections):
        if index:
            lines.append("")
        lines.append(f"[{section.name}]")
        for entry in section.entries:
            lines.append(f"{entry.key}={entry.value}")
    payload = "\n".join(lines).rstrip() + "\n"
    path.write_text(payload, encoding="utf-8")


def document_to_json_payload(document: IniDocument) -> dict[str, object]:
    return {
        "sections": [
            {
                "name": section.name,
                "entries": [{"key": entry.key, "value": entry.value} for entry in section.entries],
            }
            for section in document.sections
        ]
    }


def document_from_json_payload(payload: object) -> IniDocument:
    if not isinstance(payload, dict):
        raise ValueError("JSON root must be an object.")

    if "sections" in payload:
        section_payloads = payload.get("sections")
        if not isinstance(section_payloads, list):
            raise ValueError("The 'sections' field must be a list.")
    else:
        section_payloads = [
            {
                "name": name,
                "entries": [{"key": key, "value": value} for key, value in values.items()],
            }
            for name, values in payload.items()
            if isinstance(values, dict)
        ]

    sections: list[IniSection] = []
    seen_sections: set[str] = set()
    for section_payload in section_payloads:
        if not isinstance(section_payload, dict):
            raise ValueError("Each section must be an object.")
        name = str(section_payload.get("name", "")).strip()
        if not name:
            raise ValueError("Each section needs a non-empty name.")
        if name in seen_sections:
            raise ValueError(f"Duplicate section name: {name}")
        seen_sections.add(name)

        entry_payloads = section_payload.get("entries", [])
        if not isinstance(entry_payloads, list):
            raise ValueError(f"Section {name} has an invalid entries list.")

        entries: list[IniEntry] = []
        seen_keys: set[str] = set()
        for entry_payload in entry_payloads:
            if not isinstance(entry_payload, dict):
                raise ValueError(f"Section {name} has a non-object entry.")
            key = str(entry_payload.get("key", "")).strip()
            if not key:
                raise ValueError(f"Section {name} contains an empty key.")
            if key in seen_keys:
                raise ValueError(f"Section {name} contains duplicate key {key}.")
            seen_keys.add(key)
            value = entry_payload.get("value", "")
            entries.append(IniEntry(key=key, value="" if value is None else str(value)))
        sections.append(IniSection(name=name, entries=entries))
    return IniDocument(sections=sections)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Settings Editor")
        self.current_path: Path | None = None
        self.document = IniDocument()
        self._dirty = False
        self._updating_sections = False
        self._updating_entries = False
        self._build_ui()
        self._apply_style_overrides()
        self.setMinimumSize(1240, 820)
        self.resize(1460, 920)
        if DEFAULT_INI_PATH.exists():
            self._load_document(DEFAULT_INI_PATH)
        else:
            self._set_dirty(False)
            self._refresh_ui()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(14)

        header_card, header_layout = self._build_card(
            "Settings Editor",
            "Load VB6-compatible INI files, edit them by section, exchange them as JSON, and keep restorable pre-save snapshots.",
        )
        badge_row = QtWidgets.QHBoxLayout()
        badge_row.setSpacing(8)
        self.file_path_label = QtWidgets.QLabel("No INI loaded")
        self.file_path_label.setObjectName("pathPill")
        self.file_path_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        self.summary_label = QtWidgets.QLabel("0 sections, 0 keys")
        self.summary_label.setObjectName("valuePill")
        self.snapshot_label = QtWidgets.QLabel("No snapshots yet")
        self.snapshot_label.setObjectName("valuePill")
        badge_row.addWidget(self.file_path_label, stretch=1)
        badge_row.addWidget(self.summary_label)
        badge_row.addWidget(self.snapshot_label)
        header_layout.addLayout(badge_row)

        action_row = QtWidgets.QHBoxLayout()
        self.load_defaults_btn = QtWidgets.QPushButton("Load Defaults.ini")
        self.open_btn = QtWidgets.QPushButton("Open INI")
        self.reload_btn = QtWidgets.QPushButton("Reload")
        self.save_btn = QtWidgets.QPushButton("Save")
        self.save_btn.setObjectName("accent")
        self.save_as_btn = QtWidgets.QPushButton("Save As")
        self.export_json_btn = QtWidgets.QPushButton("Export JSON")
        self.import_json_btn = QtWidgets.QPushButton("Import JSON")
        for button in (
            self.load_defaults_btn,
            self.open_btn,
            self.reload_btn,
            self.save_btn,
            self.save_as_btn,
            self.export_json_btn,
            self.import_json_btn,
        ):
            action_row.addWidget(button)
        action_row.addStretch(1)
        header_layout.addLayout(action_row)
        layout.addWidget(header_card)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter, stretch=1)

        sections_card, sections_layout = self._build_card(
            "Sections",
            "Browse the INI by logical section instead of scrolling a single flat file.",
        )
        self.section_search = QtWidgets.QLineEdit()
        self.section_search.setPlaceholderText("Filter sections")
        sections_layout.addWidget(self.section_search)
        self.section_list = QtWidgets.QListWidget()
        self.section_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        sections_layout.addWidget(self.section_list, stretch=1)
        section_button_row = QtWidgets.QHBoxLayout()
        self.add_section_btn = QtWidgets.QPushButton("Add Section")
        self.rename_section_btn = QtWidgets.QPushButton("Rename")
        self.remove_section_btn = QtWidgets.QPushButton("Remove")
        section_button_row.addWidget(self.add_section_btn)
        section_button_row.addWidget(self.rename_section_btn)
        section_button_row.addWidget(self.remove_section_btn)
        sections_layout.addLayout(section_button_row)
        splitter.addWidget(sections_card)

        editor_card, editor_layout = self._build_card(
            "Section Values",
            "Edit key/value pairs directly. Keys stay ordered exactly as shown in the table.",
        )
        self.section_title = QtWidgets.QLabel("Select a section")
        self.section_title.setObjectName("sectionTitle")
        self.section_help = QtWidgets.QLabel(
            "Use Add Key to insert a new setting. Edit cells in place. Remove selected rows to delete settings from the current section."
        )
        self.section_help.setObjectName("subtitle")
        self.section_help.setWordWrap(True)
        editor_layout.addWidget(self.section_title)
        editor_layout.addWidget(self.section_help)
        self.entry_table = QtWidgets.QTableWidget(0, 2)
        self.entry_table.setObjectName("entryTable")
        self.entry_table.setHorizontalHeaderLabels(["Key", "Value"])
        self.entry_table.verticalHeader().setVisible(False)
        self.entry_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.entry_table.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.entry_table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        entry_header = self.entry_table.horizontalHeader()
        entry_header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        entry_header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        editor_layout.addWidget(self.entry_table, stretch=1)
        entry_button_row = QtWidgets.QHBoxLayout()
        self.add_entry_btn = QtWidgets.QPushButton("Add Key")
        self.remove_entry_btn = QtWidgets.QPushButton("Remove Selected")
        entry_button_row.addWidget(self.add_entry_btn)
        entry_button_row.addWidget(self.remove_entry_btn)
        entry_button_row.addStretch(1)
        editor_layout.addLayout(entry_button_row)
        splitter.addWidget(editor_card)

        history_card, history_layout = self._build_card(
            "Snapshots And Exchange",
            "Every save captures the previous INI into an internal history folder so you can inspect and restore older revisions.",
        )
        self.history_list = QtWidgets.QListWidget()
        self.history_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        history_layout.addWidget(self.history_list, stretch=1)
        history_button_row = QtWidgets.QHBoxLayout()
        self.restore_snapshot_btn = QtWidgets.QPushButton("Load Snapshot Into Editor")
        self.open_history_btn = QtWidgets.QPushButton("Open History Folder")
        history_button_row.addWidget(self.restore_snapshot_btn)
        history_button_row.addWidget(self.open_history_btn)
        history_layout.addLayout(history_button_row)
        self.status_console = QtWidgets.QPlainTextEdit()
        self.status_console.setReadOnly(True)
        self.status_console.setObjectName("historyConsole")
        self.status_console.setMinimumHeight(180)
        history_layout.addWidget(self.status_console)
        splitter.addWidget(history_card)
        splitter.setSizes([250, 720, 340])

        self.load_defaults_btn.clicked.connect(self._load_defaults)
        self.open_btn.clicked.connect(self._open_ini)
        self.reload_btn.clicked.connect(self._reload_current_file)
        self.save_btn.clicked.connect(self._save_current_file)
        self.save_as_btn.clicked.connect(self._save_as)
        self.export_json_btn.clicked.connect(self._export_json)
        self.import_json_btn.clicked.connect(self._import_json)
        self.section_search.textChanged.connect(self._refresh_section_list)
        self.section_list.currentRowChanged.connect(self._on_section_selection_changed)
        self.add_section_btn.clicked.connect(self._add_section)
        self.rename_section_btn.clicked.connect(self._rename_section)
        self.remove_section_btn.clicked.connect(self._remove_section)
        self.add_entry_btn.clicked.connect(self._add_entry)
        self.remove_entry_btn.clicked.connect(self._remove_selected_entries)
        self.entry_table.itemChanged.connect(self._on_entry_item_changed)
        self.restore_snapshot_btn.clicked.connect(self._load_selected_snapshot_into_editor)
        self.open_history_btn.clicked.connect(self._open_history_folder)
        self.history_list.itemDoubleClicked.connect(lambda _item: self._load_selected_snapshot_into_editor())

    def _build_card(self, title_text: str, subtitle_text: str | None = None) -> tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)
        title = QtWidgets.QLabel(title_text)
        title.setObjectName("title")
        layout.addWidget(title)
        if subtitle_text:
            subtitle = QtWidgets.QLabel(subtitle_text)
            subtitle.setObjectName("subtitle")
            subtitle.setWordWrap(True)
            layout.addWidget(subtitle)
        apply_card_shadow(card)
        return card, layout

    def _apply_style_overrides(self) -> None:
        self.setStyleSheet(
            self.styleSheet()
            + """
            QLabel#pathPill {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(122, 2, 25, 0.16);
                border-radius: 16px;
                padding: 8px 10px;
                font-weight: 640;
            }
            QLabel#sectionTitle {
                color: #7a0219;
                font-size: 18px;
                font-weight: 720;
            }
            QListWidget,
            QTableWidget {
                background: rgba(255, 255, 255, 0.90);
                border: 1px solid rgba(122, 2, 25, 0.14);
                border-radius: 16px;
                padding: 6px;
                selection-background-color: rgba(122, 2, 25, 0.18);
            }
            QPlainTextEdit#historyConsole {
                background: rgba(31, 23, 22, 0.92);
                color: #fef0c7;
                border-radius: 16px;
                border: 1px solid rgba(255, 215, 130, 0.24);
                padding: 8px;
            }
            """
        )

    def _append_status(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_console.appendPlainText(f"[{timestamp}] {message}")

    def _history_dir_for(self, target: Path) -> Path:
        return target.parent / ".rapidpy_history" / target.stem

    def _history_files_for_current_path(self) -> list[Path]:
        if self.current_path is None:
            return []
        history_dir = self._history_dir_for(self.current_path)
        if not history_dir.exists():
            return []
        return sorted(history_dir.glob(f"*{self.current_path.suffix}"), reverse=True)

    def _create_snapshot(self, target: Path) -> Path | None:
        if not target.exists():
            return None
        history_dir = self._history_dir_for(target)
        history_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        suffix = target.suffix or ".ini"
        snapshot = history_dir / f"{timestamp}{suffix}"
        counter = 1
        while snapshot.exists():
            snapshot = history_dir / f"{timestamp}_{counter}{suffix}"
            counter += 1
        shutil.copy2(target, snapshot)
        return snapshot

    def _current_section(self) -> IniSection | None:
        item = self.section_list.currentItem()
        if item is None:
            return None
        section_name = item.data(QtCore.Qt.UserRole)
        for section in self.document.sections:
            if section.name == section_name:
                return section
        return None

    def _selected_section_name(self) -> str | None:
        section = self._current_section()
        return None if section is None else section.name

    def _set_dirty(self, dirty: bool) -> None:
        self._dirty = dirty
        suffix = " *" if dirty else ""
        self.setWindowTitle(f"RapidPy Settings Editor{suffix}")

    def _section_and_key_count(self) -> tuple[int, int]:
        section_count = len(self.document.sections)
        key_count = sum(len(section.entries) for section in self.document.sections)
        return section_count, key_count

    def _refresh_ui(self, preferred_section: str | None = None) -> None:
        self._refresh_summary_labels()
        self._refresh_section_list(preferred_section)
        self._refresh_history_list()

    def _refresh_summary_labels(self) -> None:
        self.file_path_label.setText(str(self.current_path) if self.current_path else "Editor content is not attached to an INI yet")
        section_count, key_count = self._section_and_key_count()
        self.summary_label.setText(f"{section_count} sections, {key_count} keys")
        snapshot_count = len(self._history_files_for_current_path())
        self.snapshot_label.setText("No snapshots yet" if snapshot_count == 0 else f"{snapshot_count} snapshots")
        self.reload_btn.setEnabled(self.current_path is not None)
        self.save_btn.setEnabled(bool(self.document.sections) or self.current_path is not None)
        self.save_as_btn.setEnabled(bool(self.document.sections))
        self.export_json_btn.setEnabled(bool(self.document.sections))
        has_history = snapshot_count > 0
        self.restore_snapshot_btn.setEnabled(has_history)
        self.open_history_btn.setEnabled(self.current_path is not None)

    def _refresh_section_list(self, preferred_section: str | None = None) -> None:
        if isinstance(preferred_section, str) and not preferred_section.strip():
            preferred_section = None
        if preferred_section is None:
            preferred_section = self._selected_section_name()
        filter_text = self.section_search.text().strip().lower()
        self._updating_sections = True
        self.section_list.clear()
        for section in self.document.sections:
            if filter_text and filter_text not in section.name.lower():
                continue
            item = QtWidgets.QListWidgetItem(section.name)
            item.setData(QtCore.Qt.UserRole, section.name)
            item.setToolTip(f"{len(section.entries)} keys")
            self.section_list.addItem(item)
        self._updating_sections = False

        if self.section_list.count() == 0:
            self.section_title.setText("No section selected")
            self._refresh_entry_table(None)
            return

        target_name = preferred_section
        if target_name:
            for index in range(self.section_list.count()):
                item = self.section_list.item(index)
                if item.data(QtCore.Qt.UserRole) == target_name:
                    self.section_list.setCurrentRow(index)
                    break
            else:
                self.section_list.setCurrentRow(0)
        else:
            self.section_list.setCurrentRow(0)

    def _refresh_entry_table(self, section: IniSection | None, preferred_row: int | None = None) -> None:
        self._updating_entries = True
        self.entry_table.setRowCount(0)
        if section is None:
            self.section_title.setText("No section selected")
            self.add_entry_btn.setEnabled(False)
            self.remove_entry_btn.setEnabled(False)
            self._updating_entries = False
            return

        self.section_title.setText(section.name)
        self.add_entry_btn.setEnabled(True)
        self.remove_entry_btn.setEnabled(bool(section.entries))
        self.entry_table.setRowCount(len(section.entries))
        for row, entry in enumerate(section.entries):
            key_item = QtWidgets.QTableWidgetItem(entry.key)
            value_item = QtWidgets.QTableWidgetItem(entry.value)
            self.entry_table.setItem(row, 0, key_item)
            self.entry_table.setItem(row, 1, value_item)
        self._updating_entries = False
        if section.entries:
            self.entry_table.selectRow(0 if preferred_row is None else max(0, min(preferred_row, len(section.entries) - 1)))

    def _refresh_history_list(self) -> None:
        self.history_list.clear()
        for snapshot in self._history_files_for_current_path():
            item = QtWidgets.QListWidgetItem(snapshot.stem)
            item.setData(QtCore.Qt.UserRole, str(snapshot))
            item.setToolTip(str(snapshot))
            self.history_list.addItem(item)
        self.restore_snapshot_btn.setEnabled(self.history_list.count() > 0)

    def _load_document(self, path: Path) -> None:
        self.document = load_ini_document(path)
        self.current_path = path
        self._set_dirty(False)
        self._refresh_ui(preferred_section=self.document.sections[0].name if self.document.sections else None)
        self._append_status(f"Loaded {path}")

    def _maybe_save_before_destructive_action(self) -> bool:
        if not self._dirty:
            return True
        response = QtWidgets.QMessageBox.question(
            self,
            "Unsaved Changes",
            "The editor has unsaved changes. Save them first?",
            QtWidgets.QMessageBox.StandardButton.Save
            | QtWidgets.QMessageBox.StandardButton.Discard
            | QtWidgets.QMessageBox.StandardButton.Cancel,
            QtWidgets.QMessageBox.StandardButton.Save,
        )
        if response == QtWidgets.QMessageBox.StandardButton.Cancel:
            return False
        if response == QtWidgets.QMessageBox.StandardButton.Discard:
            return True
        return self._save_current_file()

    def _save_document_to(self, target: Path) -> bool:
        try:
            snapshot = self._create_snapshot(target)
            save_ini_document(self.document, target)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save Failed", str(exc))
            self._append_status(f"Save failed: {exc}")
            return False
        self.current_path = target
        self._set_dirty(False)
        self._refresh_ui(preferred_section=self._selected_section_name())
        if snapshot is None:
            self._append_status(f"Saved {target}")
        else:
            self._append_status(f"Saved {target} after snapshotting {snapshot.name}")
        return True

    def _load_defaults(self) -> None:
        if not DEFAULT_INI_PATH.exists():
            QtWidgets.QMessageBox.warning(self, "Defaults Missing", f"Could not find {DEFAULT_INI_PATH}")
            return
        if not self._maybe_save_before_destructive_action():
            return
        self._load_document(DEFAULT_INI_PATH)

    def _open_ini(self) -> None:
        if not self._maybe_save_before_destructive_action():
            return
        selected, _filter = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open INI File",
            str((self.current_path or DEFAULT_INI_PATH).parent if (self.current_path or DEFAULT_INI_PATH) else REPO_ROOT),
            "INI files (*.ini);;All files (*.*)",
        )
        if not selected:
            return
        self._load_document(Path(selected))

    def _reload_current_file(self) -> None:
        if self.current_path is None:
            return
        if not self._maybe_save_before_destructive_action():
            return
        self._load_document(self.current_path)

    def _save_current_file(self) -> bool:
        if self.current_path is None:
            return self._save_as()
        return self._save_document_to(self.current_path)

    def _save_as(self) -> bool:
        selected, _filter = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save INI File",
            str(self.current_path or DEFAULT_INI_PATH),
            "INI files (*.ini);;All files (*.*)",
        )
        if not selected:
            return False
        return self._save_document_to(Path(selected))

    def _export_json(self) -> None:
        selected, _filter = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export JSON",
            str((self.current_path or DEFAULT_INI_PATH).with_suffix(".json")),
            "JSON files (*.json)",
        )
        if not selected:
            return
        target = Path(selected)
        try:
            target.write_text(json.dumps(document_to_json_payload(self.document), indent=2), encoding="utf-8")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Export Failed", str(exc))
            self._append_status(f"JSON export failed: {exc}")
            return
        self._append_status(f"Exported JSON to {target}")

    def _import_json(self) -> None:
        if not self._maybe_save_before_destructive_action():
            return
        selected, _filter = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Import JSON",
            str((self.current_path or DEFAULT_INI_PATH).with_suffix(".json")),
            "JSON files (*.json)",
        )
        if not selected:
            return
        try:
            payload = json.loads(Path(selected).read_text(encoding="utf-8"))
            self.document = document_from_json_payload(payload)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Import Failed", str(exc))
            self._append_status(f"JSON import failed: {exc}")
            return
        self._set_dirty(True)
        self._refresh_ui(preferred_section=self.document.sections[0].name if self.document.sections else None)
        self._append_status(f"Loaded JSON into editor from {selected}")

    def _on_section_selection_changed(self, _row: int) -> None:
        if self._updating_sections:
            return
        self._refresh_entry_table(self._current_section())

    def _add_section(self) -> None:
        name, accepted = QtWidgets.QInputDialog.getText(self, "Add Section", "Section name:")
        name = name.strip()
        if not accepted or not name:
            return
        if any(section.name == name for section in self.document.sections):
            QtWidgets.QMessageBox.warning(self, "Duplicate Section", f"Section {name} already exists.")
            return
        self.document.sections.append(IniSection(name=name))
        self._set_dirty(True)
        self._refresh_ui(preferred_section=name)
        self._append_status(f"Added section {name}")

    def _rename_section(self) -> None:
        section = self._current_section()
        if section is None:
            return
        name, accepted = QtWidgets.QInputDialog.getText(self, "Rename Section", "Section name:", text=section.name)
        name = name.strip()
        if not accepted or not name or name == section.name:
            return
        if any(other.name == name for other in self.document.sections if other is not section):
            QtWidgets.QMessageBox.warning(self, "Duplicate Section", f"Section {name} already exists.")
            return
        old_name = section.name
        section.name = name
        self._set_dirty(True)
        self._refresh_ui(preferred_section=name)
        self._append_status(f"Renamed section {old_name} to {name}")

    def _remove_section(self) -> None:
        section = self._current_section()
        if section is None:
            return
        response = QtWidgets.QMessageBox.question(
            self,
            "Remove Section",
            f"Remove section {section.name} and all of its keys?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No,
        )
        if response != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        self.document.sections = [item for item in self.document.sections if item is not section]
        self._set_dirty(True)
        next_section = self.document.sections[0].name if self.document.sections else None
        self._refresh_ui(preferred_section=next_section)
        self._append_status(f"Removed section {section.name}")

    def _add_entry(self) -> None:
        section = self._current_section()
        if section is None:
            return
        key, accepted = QtWidgets.QInputDialog.getText(self, "Add Key", "Key name:")
        key = key.strip()
        if not accepted or not key:
            return
        if any(entry.key == key for entry in section.entries):
            QtWidgets.QMessageBox.warning(self, "Duplicate Key", f"Key {key} already exists in {section.name}.")
            return
        value, accepted = QtWidgets.QInputDialog.getText(self, "Add Key", f"Value for {key}:")
        if not accepted:
            return
        section.entries.append(IniEntry(key=key, value=value.strip()))
        self._set_dirty(True)
        self._refresh_entry_table(section, preferred_row=len(section.entries) - 1)
        self._refresh_summary_labels()
        self._append_status(f"Added key {key} to {section.name}")

    def _remove_selected_entries(self) -> None:
        section = self._current_section()
        if section is None:
            return
        rows = sorted({index.row() for index in self.entry_table.selectionModel().selectedRows()}, reverse=True)
        if not rows:
            return
        removed_keys = [section.entries[row].key for row in rows]
        for row in rows:
            del section.entries[row]
        self._set_dirty(True)
        self._refresh_entry_table(section, preferred_row=min(rows[-1], max(0, len(section.entries) - 1)) if section.entries else None)
        self._refresh_summary_labels()
        self._append_status(f"Removed {', '.join(removed_keys)} from {section.name}")

    def _on_entry_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        if self._updating_entries:
            return
        section = self._current_section()
        if section is None:
            return
        row = item.row()
        if row >= len(section.entries):
            return
        entry = section.entries[row]
        if item.column() == 0:
            new_key = item.text().strip()
            if not new_key:
                QtWidgets.QMessageBox.warning(self, "Invalid Key", "Key names cannot be empty.")
                self._refresh_entry_table(section, preferred_row=row)
                return
            if any(other.key == new_key for index, other in enumerate(section.entries) if index != row):
                QtWidgets.QMessageBox.warning(self, "Duplicate Key", f"Key {new_key} already exists in {section.name}.")
                self._refresh_entry_table(section, preferred_row=row)
                return
            if entry.key != new_key:
                old_key = entry.key
                entry.key = new_key
                self._set_dirty(True)
                self._refresh_summary_labels()
                self._append_status(f"Renamed key {old_key} to {new_key} in {section.name}")
            return
        new_value = item.text().strip()
        if entry.value != new_value:
            entry.value = new_value
            self._set_dirty(True)
            self._append_status(f"Updated {section.name}.{entry.key}")

    def _load_selected_snapshot_into_editor(self) -> None:
        item = self.history_list.currentItem()
        if item is None:
            return
        snapshot_path = Path(str(item.data(QtCore.Qt.UserRole)))
        try:
            self.document = load_ini_document(snapshot_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Snapshot Load Failed", str(exc))
            self._append_status(f"Snapshot load failed: {exc}")
            return
        self._set_dirty(True)
        self._refresh_ui(preferred_section=self.document.sections[0].name if self.document.sections else None)
        self._append_status(f"Loaded snapshot {snapshot_path.name} into the editor. Save to write it back to the active INI.")

    def _open_history_folder(self) -> None:
        if self.current_path is None:
            return
        history_dir = self._history_dir_for(self.current_path)
        history_dir.mkdir(parents=True, exist_ok=True)
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(history_dir)))

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        if self._maybe_save_before_destructive_action():
            event.accept()
            return
        event.ignore()


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "settings_editor_icon.png", assets_dir)
    window = MainWindow()
    set_app_icon(window, "settings_editor_icon.png", assets_dir)
    window.show()
    return app.exec()
