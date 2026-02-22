from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolTip,
    QVBoxLayout,
    QWidget,
)


CHECK_COLUMN_CANDIDATES = [
    "check number",
    "checkno",
    "ref no",
    "num",
    "document no",
    "check #",
    "cheque #",
    "check",
    "cheque",
]
VENDOR_COLUMN_CANDIDATES = ["vendor", "payee", "name", "vendor name"]
CHECK_TEXT_PREFIXES = ("check", "cheque", "chk")


def normalize_check_number(
    value: object,
    normalize_mode: bool = True,
    extract_from_text_mode: bool = True,
) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    if not normalize_mode:
        return text

    cleaned = text.strip()
    if cleaned.endswith(".0"):
        cleaned = cleaned[:-2].strip()

    if extract_from_text_mode:
        prefix_pattern = r"^\s*(?:" + "|".join(CHECK_TEXT_PREFIXES) + r")\s*[-:#]*\s*(\d+)\s*$"
        pref_match = re.match(prefix_pattern, cleaned, flags=re.IGNORECASE)
        if pref_match:
            return pref_match.group(1)

        number_match = re.search(r"(\d+)", cleaned)
        if number_match:
            return number_match.group(1)

    return cleaned


class SplashScreen(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedSize(QSize(900, 500))
        self._background = self._build_background()

    def _build_background(self) -> QPixmap:
        splash_path = Path(__file__).parent / "assets" / "splash_bg.jpg"
        canvas = QPixmap(self.size())
        painter = QPainter(canvas)

        source = QPixmap(str(splash_path)) if splash_path.exists() else QPixmap()
        if not source.isNull():
            scaled = source.scaled(self.size(), Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            x = (scaled.width() - self.width()) // 2
            y = (scaled.height() - self.height()) // 2
            painter.drawPixmap(0, 0, scaled, x, y, self.width(), self.height())
        else:
            gradient = QLinearGradient(0, 0, self.width(), self.height())
            gradient.setColorAt(0.0, QColor("#10243f"))
            gradient.setColorAt(0.5, QColor("#1d3f66"))
            gradient.setColorAt(1.0, QColor("#11253b"))
            painter.fillRect(canvas.rect(), gradient)

        overlay = QLinearGradient(0, 0, 0, self.height())
        overlay.setColorAt(0.0, QColor(0, 0, 0, 110))
        overlay.setColorAt(1.0, QColor(0, 0, 0, 185))
        painter.fillRect(canvas.rect(), overlay)
        painter.end()
        return canvas

    def paintEvent(self, event) -> None:  # type: ignore[override]
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.drawPixmap(self.rect(), self._background)

        painter.setPen(QPen(QColor("#ffffff")))
        painter.setFont(QFont("Segoe UI", 44, QFont.Bold))
        painter.drawText(self.rect().adjusted(40, -30, -40, -20), Qt.AlignCenter, "Jose's CSV Check Merger")

        painter.setPen(QPen(QColor("#dce6f4")))
        painter.setFont(QFont("Segoe UI", 15))
        painter.drawText(self.rect().adjusted(40, 170, -40, -20), Qt.AlignHCenter, "QuickBooks Vendor Name Updater")


class LoginDialog(QDialog):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Login")
        self.setModal(True)
        self.setFixedSize(380, 220)

        layout = QVBoxLayout(self)

        title = QLabel("Sign in")
        title.setStyleSheet("font-size: 20px; font-weight: 700; color: #f2c14e;")

        form = QFormLayout()
        self.user_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)

        form.addRow("User:", self.user_input)
        form.addRow("Password:", self.password_input)

        login_btn = QPushButton("Login")
        login_btn.clicked.connect(self.try_login)

        self.message_label = QLabel("")
        self.message_label.setStyleSheet("color: #ffb3b3;")

        layout.addWidget(title)
        layout.addLayout(form)
        layout.addWidget(self.message_label)
        layout.addWidget(login_btn)

        self.setStyleSheet(
            "QDialog { background: #111111; color: #f5f5f5; }"
            "QLineEdit { background: #0f0f10; color: #f5f5f5; border: 1px solid #c29b2d; border-radius: 6px; padding: 6px; }"
            "QPushButton { background: #1c1c1d; color: #fff8e6; border: 1px solid #d4af37; border-radius: 6px; min-height: 30px; font-weight: 600; }"
            "QPushButton:hover { background: #2a2a2c; }"
        )

    def try_login(self) -> None:
        if self.user_input.text().strip() == "Kiri" and self.password_input.text() == "Jcr16331878":
            self.accept()
            return

        self.message_label.setText("Invalid username or password.")

        self._build_ui()
        self._apply_styles()

    def _build_ui(self) -> None:
        container = QWidget(self)
        self.setCentralWidget(container)
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(20, 18, 20, 18)
        root_layout.setSpacing(12)

        title = QLabel("QuickBooks Check Vendor Updater")
        title.setObjectName("mainTitle")
        subtitle = QLabel(
            "Load both CSVs, confirm column mapping, then update vendor/payee names by check number."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("subTitle")
        root_layout.addWidget(title)
        root_layout.addWidget(subtitle)

        file_group = QGroupBox("Step 1 — Select Files")
        file_group.setObjectName("panel")
        file_layout = QGridLayout(file_group)

        self.quickbooks_path = QLineEdit()
        self.quickbooks_path.setReadOnly(True)
        self.quickbooks_path.setPlaceholderText("Choose the QuickBooks Upload file (CSV/Excel)...")
        qb_btn = QPushButton("Browse QuickBooks File")
        qb_btn.clicked.connect(self.load_quickbooks_csv)

        self.reference_path = QLineEdit()
        self.reference_path.setReadOnly(True)
        self.reference_path.setPlaceholderText("Choose the Check Reference file (CSV/Excel)...")
        ref_btn = QPushButton("Browse Reference File")
        ref_btn.clicked.connect(self.load_reference_csv)

        file_layout.addWidget(QLabel("QuickBooks Upload file (target):"), 0, 0)
        file_layout.addWidget(self.quickbooks_path, 0, 1)
        file_layout.addWidget(qb_btn, 0, 2)

        file_layout.addWidget(QLabel("Check Reference file (lookup):"), 1, 0)
        file_layout.addWidget(self.reference_path, 1, 1)
        file_layout.addWidget(ref_btn, 1, 2)

        mapping_group = QGroupBox("Step 2 — Confirm Column Mapping")
        mapping_group.setObjectName("panel")
        mapping_layout = QFormLayout(mapping_group)

        self.qb_check_combo = QComboBox()
        self.qb_vendor_combo = QComboBox()
        self.ref_check_combo = QComboBox()
        self.ref_vendor_combo = QComboBox()

        mapping_layout.addRow("QuickBooks Check Number column:", self.qb_check_combo)
        mapping_layout.addRow("QuickBooks Vendor/Payee to update:", self.qb_vendor_combo)
        mapping_layout.addRow("Reference Check Number column:", self.ref_check_combo)
        mapping_layout.addRow("Reference Vendor Name column:", self.ref_vendor_combo)

        mapping_hint = QLabel(
            "Tip: Common check columns include Check Number, Num, Ref No, Document No, or values like 'Check 101'."
        )
        mapping_hint.setWordWrap(True)

        options_group = QGroupBox("Step 3 — Matching Options")
        options_group.setObjectName("panel")
        options_layout = QHBoxLayout(options_group)
        self.normalize_checkbox = QCheckBox("Normalize check values (trim + remove .0)")
        self.normalize_checkbox.setChecked(True)
        self.extract_checkbox = QCheckBox("Extract number from text (e.g., 'Check 101')")
        self.extract_checkbox.setChecked(True)
        self.unmatched_checkbox = QCheckBox("Export unmatched rows report")
        self.unmatched_checkbox.setChecked(True)
        self.backup_checkbox = QCheckBox("Create backup of original QuickBooks file on save")
        self.backup_checkbox.setChecked(True)

        options_layout.addWidget(self.normalize_checkbox)
        options_layout.addWidget(self.extract_checkbox)
        options_layout.addWidget(self.unmatched_checkbox)
        options_layout.addWidget(self.backup_checkbox)
        options_layout.addStretch()

        actions_layout = QHBoxLayout()
        self.process_btn = QPushButton("Step 4 — Update Vendor Names")
        self.process_btn.clicked.connect(self.process_updates)
        self.save_btn = QPushButton("Step 5 — Save Updated CSV")
        self.save_btn.clicked.connect(self.save_updated_csv)
        self.save_btn.setEnabled(False)
        reset_btn = QPushButton("Reset")
        reset_btn.clicked.connect(self.reset_app)

        help_btn = QPushButton("How matching works")
        help_btn.clicked.connect(self.show_matching_help)

        actions_layout.addWidget(self.process_btn)
        actions_layout.addWidget(self.save_btn)
        actions_layout.addWidget(help_btn)
        actions_layout.addWidget(reset_btn)
        actions_layout.addStretch()

        self.status_label = QLabel("Status: Ready")
        self.status_label.setObjectName("statusLabel")

        self.summary_box = QTextEdit()
        self.summary_box.setObjectName("summaryBox")
        self.summary_box.setReadOnly(True)
        self.summary_box.setFixedHeight(130)
        self.summary_box.setPlaceholderText("Summary will appear here after processing...")

        self.preview_table = QTableWidget()
        self.preview_table.setObjectName("previewTable")
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        root_layout.addWidget(file_group)
        root_layout.addWidget(mapping_group)
        root_layout.addWidget(mapping_hint)
        root_layout.addWidget(options_group)
        root_layout.addLayout(actions_layout)
        root_layout.addWidget(self.status_label)
        root_layout.addWidget(QLabel("Summary"))
        root_layout.addWidget(self.summary_box)
        root_layout.addWidget(QLabel("Preview (first 100 rows)"))
        root_layout.addWidget(self.preview_table)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #0f0f10; color: #f5f5f5; }
            QLabel { color: #f5f5f5; }
            #mainTitle { font-size: 30px; font-weight: 800; color: #f2c14e; letter-spacing: 0.5px; }
            #subTitle { color: #f3e8ca; margin-bottom: 8px; font-size: 13px; }
            #statusLabel {
                font-weight: 700;
                color: #fff8e6;
                background: #1f1f20;
                border: 1px solid #d4af37;
                border-radius: 8px;
                padding: 8px 10px;
            }
            QGroupBox#panel {
                font-weight: 700;
                border: 1px solid #d4af37;
                border-radius: 12px;
                margin-top: 10px;
                background: #151516;
                padding: 10px;
                color: #f5f5f5;
            }
            QGroupBox#panel::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
                color: #f2c14e;
            }
            QPushButton {
                min-height: 34px;
                padding: 6px 12px;
                border-radius: 8px;
                border: 1px solid #d4af37;
                background: #1c1c1d;
                color: #fff8e6;
                font-weight: 600;
            }
            QPushButton:hover { background: #2a2a2c; }
            QPushButton:pressed { background: #111111; }
            QPushButton:disabled {
                color: #8b8b8b;
                border: 1px solid #545454;
                background: #1a1a1a;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #0f0f10;
                color: #f5f5f5;
                border: 1px solid #c29b2d;
                border-radius: 8px;
                padding: 6px;
                selection-background-color: #d4af37;
                selection-color: #111111;
            }
            QCheckBox { color: #f5f5f5; spacing: 8px; }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #c29b2d;
                border-radius: 4px;
                background: #111111;
            }
            QCheckBox::indicator:checked { background: #d4af37; }
            #summaryBox { background: #101011; color: #f5f5f5; }
            #previewTable {
                gridline-color: #3c3c3f;
                alternate-background-color: #171718;
                background: #101011;
                color: #f5f5f5;
                border: 1px solid #c29b2d;
                border-radius: 8px;
            }
            QHeaderView::section {
                background: #d4af37;
                color: #151516;
                padding: 6px;
                border: 0;
                border-right: 1px solid #8a6b1d;
                border-bottom: 1px solid #8a6b1d;
                font-weight: 700;
            }
            QToolTip {
                background-color: #111111;
                color: #f5f5f5;
                border: 1px solid #d4af37;
                padding: 6px;
            }
            """
        )


class CheckVendorUpdater(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("QuickBooks Check Vendor Updater")
        self.resize(1240, 800)

        self.quickbooks_df: Optional[pd.DataFrame] = None
        self.reference_df: Optional[pd.DataFrame] = None
        self.updated_df: Optional[pd.DataFrame] = None
        self.unmatched_df: Optional[pd.DataFrame] = None
        self.duplicates: Dict[str, int] = {}

        self._build_ui()
        self._apply_styles()

    def _build_ui(self) -> None:
        container = QWidget(self)
        self.setCentralWidget(container)
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(20, 18, 20, 18)
        root_layout.setSpacing(12)

        title = QLabel("QuickBooks Check Vendor Updater")
        title.setObjectName("mainTitle")
        subtitle = QLabel(
            "Load both CSVs, confirm column mapping, then update vendor/payee names by check number."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("subTitle")
        root_layout.addWidget(title)
        root_layout.addWidget(subtitle)

        file_group = QGroupBox("Step 1 — Select Files")
        file_group.setObjectName("panel")
        file_layout = QGridLayout(file_group)

        self.quickbooks_path = QLineEdit()
        self.quickbooks_path.setReadOnly(True)
        self.quickbooks_path.setPlaceholderText("Choose the QuickBooks Upload file (CSV/Excel)...")
        qb_btn = QPushButton("Browse QuickBooks File")
        qb_btn.clicked.connect(self.load_quickbooks_csv)

        self.reference_path = QLineEdit()
        self.reference_path.setReadOnly(True)
        self.reference_path.setPlaceholderText("Choose the Check Reference file (CSV/Excel)...")
        ref_btn = QPushButton("Browse Reference File")
        ref_btn.clicked.connect(self.load_reference_csv)

        file_layout.addWidget(QLabel("QuickBooks Upload file (target):"), 0, 0)
        file_layout.addWidget(self.quickbooks_path, 0, 1)
        file_layout.addWidget(qb_btn, 0, 2)

        file_layout.addWidget(QLabel("Check Reference file (lookup):"), 1, 0)
        file_layout.addWidget(self.reference_path, 1, 1)
        file_layout.addWidget(ref_btn, 1, 2)

        mapping_group = QGroupBox("Step 2 — Confirm Column Mapping")
        mapping_group.setObjectName("panel")
        mapping_layout = QFormLayout(mapping_group)

        self.qb_check_combo = QComboBox()
        self.qb_vendor_combo = QComboBox()
        self.ref_check_combo = QComboBox()
        self.ref_vendor_combo = QComboBox()

        mapping_layout.addRow("QuickBooks Check Number column:", self.qb_check_combo)
        mapping_layout.addRow("QuickBooks Vendor/Payee to update:", self.qb_vendor_combo)
        mapping_layout.addRow("Reference Check Number column:", self.ref_check_combo)
        mapping_layout.addRow("Reference Vendor Name column:", self.ref_vendor_combo)

        mapping_hint = QLabel(
            "Tip: Common check columns include Check Number, Num, Ref No, Document No, or values like 'Check 101'."
        )
        mapping_hint.setWordWrap(True)

        options_group = QGroupBox("Step 3 — Matching Options")
        options_group.setObjectName("panel")
        options_layout = QHBoxLayout(options_group)
        self.normalize_checkbox = QCheckBox("Normalize check values (trim + remove .0)")
        self.normalize_checkbox.setChecked(True)
        self.extract_checkbox = QCheckBox("Extract number from text (e.g., 'Check 101')")
        self.extract_checkbox.setChecked(True)
        self.unmatched_checkbox = QCheckBox("Export unmatched rows report")
        self.unmatched_checkbox.setChecked(True)
        self.backup_checkbox = QCheckBox("Create backup of original QuickBooks file on save")
        self.backup_checkbox.setChecked(True)

        options_layout.addWidget(self.normalize_checkbox)
        options_layout.addWidget(self.extract_checkbox)
        options_layout.addWidget(self.unmatched_checkbox)
        options_layout.addWidget(self.backup_checkbox)
        options_layout.addStretch()

        actions_layout = QHBoxLayout()
        self.process_btn = QPushButton("Step 4 — Update Vendor Names")
        self.process_btn.clicked.connect(self.process_updates)
        self.save_btn = QPushButton("Step 5 — Save Updated CSV")
        self.save_btn.clicked.connect(self.save_updated_csv)
        self.save_btn.setEnabled(False)
        reset_btn = QPushButton("Reset")
        reset_btn.clicked.connect(self.reset_app)

        help_btn = QPushButton("How matching works")
        help_btn.clicked.connect(self.show_matching_help)

        actions_layout.addWidget(self.process_btn)
        actions_layout.addWidget(self.save_btn)
        actions_layout.addWidget(help_btn)
        actions_layout.addWidget(reset_btn)
        actions_layout.addStretch()

        self.status_label = QLabel("Status: Ready")
        self.status_label.setObjectName("statusLabel")

        self.summary_box = QTextEdit()
        self.summary_box.setObjectName("summaryBox")
        self.summary_box.setReadOnly(True)
        self.summary_box.setFixedHeight(130)
        self.summary_box.setPlaceholderText("Summary will appear here after processing...")

        self.preview_table = QTableWidget()
        self.preview_table.setObjectName("previewTable")
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        root_layout.addWidget(file_group)
        root_layout.addWidget(mapping_group)
        root_layout.addWidget(mapping_hint)
        root_layout.addWidget(options_group)
        root_layout.addLayout(actions_layout)
        root_layout.addWidget(self.status_label)
        root_layout.addWidget(QLabel("Summary"))
        root_layout.addWidget(self.summary_box)
        root_layout.addWidget(QLabel("Preview (first 100 rows)"))
        root_layout.addWidget(self.preview_table)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #0f0f10; color: #f5f5f5; }
            QLabel { color: #f5f5f5; }
            #mainTitle { font-size: 30px; font-weight: 800; color: #f2c14e; letter-spacing: 0.5px; }
            #subTitle { color: #f3e8ca; margin-bottom: 8px; font-size: 13px; }
            #statusLabel {
                font-weight: 700;
                color: #fff8e6;
                background: #1f1f20;
                border: 1px solid #d4af37;
                border-radius: 8px;
                padding: 8px 10px;
            }
            QGroupBox#panel {
                font-weight: 700;
                border: 1px solid #d4af37;
                border-radius: 12px;
                margin-top: 10px;
                background: #151516;
                padding: 10px;
                color: #f5f5f5;
            }
            QGroupBox#panel::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
                color: #f2c14e;
            }
            QPushButton {
                min-height: 34px;
                padding: 6px 12px;
                border-radius: 8px;
                border: 1px solid #d4af37;
                background: #1c1c1d;
                color: #fff8e6;
                font-weight: 600;
            }
            QPushButton:hover { background: #2a2a2c; }
            QPushButton:pressed { background: #111111; }
            QPushButton:disabled {
                color: #8b8b8b;
                border: 1px solid #545454;
                background: #1a1a1a;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #0f0f10;
                color: #f5f5f5;
                border: 1px solid #c29b2d;
                border-radius: 8px;
                padding: 6px;
                selection-background-color: #d4af37;
                selection-color: #111111;
            }
            QCheckBox { color: #f5f5f5; spacing: 8px; }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #c29b2d;
                border-radius: 4px;
                background: #111111;
            }
            QCheckBox::indicator:checked { background: #d4af37; }
            #summaryBox { background: #101011; color: #f5f5f5; }
            #previewTable {
                gridline-color: #3c3c3f;
                alternate-background-color: #171718;
                background: #101011;
                color: #f5f5f5;
                border: 1px solid #c29b2d;
                border-radius: 8px;
            }
            QHeaderView::section {
                background: #d4af37;
                color: #151516;
                padding: 6px;
                border: 0;
                border-right: 1px solid #8a6b1d;
                border-bottom: 1px solid #8a6b1d;
                font-weight: 700;
            }
            QToolTip {
                background-color: #111111;
                color: #f5f5f5;
                border: 1px solid #d4af37;
                padding: 6px;
            }
            """
        )

    def _set_status(self, text: str) -> None:
        self.status_label.setText(f"Status: {text}")

    def reset_app(self) -> None:
        self.quickbooks_df = None
        self.reference_df = None
        self.updated_df = None
        self.unmatched_df = None
        self.duplicates = {}

        self.quickbooks_path.clear()
        self.reference_path.clear()
        self.qb_check_combo.clear()
        self.qb_vendor_combo.clear()
        self.ref_check_combo.clear()
        self.ref_vendor_combo.clear()
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.summary_box.clear()
        self.save_btn.setEnabled(False)
        self._set_status("Ready")



    @staticmethod
    def _read_table(path: str) -> pd.DataFrame:
        suffix = Path(path).suffix.lower()
        if suffix == ".csv":
            return pd.read_csv(path, dtype=object)
        if suffix == ".xlsx":
            return pd.read_excel(path, dtype=object)
        raise ValueError("Unsupported file type. Please select CSV or Excel (.xlsx).")

    @staticmethod
    def _write_table(df: pd.DataFrame, path: str) -> None:
        suffix = Path(path).suffix.lower()
        if suffix == ".csv":
            df.to_csv(path, index=False)
            return
        if suffix == ".xlsx":
            df.to_excel(path, index=False)
            return
        raise ValueError("Unsupported output format. Please save as .csv or .xlsx.")

    def load_quickbooks_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select QuickBooks Upload File", "", "Data Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx)")
        if not path:
            return
        try:
            self.quickbooks_df = self._read_table(path)
            self.quickbooks_path.setText(path)
            self._populate_combo(self.qb_check_combo, list(self.quickbooks_df.columns), CHECK_COLUMN_CANDIDATES)
            self._populate_combo(self.qb_vendor_combo, list(self.quickbooks_df.columns), VENDOR_COLUMN_CANDIDATES)
            self._update_summary("Loaded QuickBooks file.")
            self._set_status("QuickBooks file loaded")
        except Exception as exc:
            self._error(f"Could not read QuickBooks file: {exc}")

    def load_reference_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Check Reference File", "", "Data Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx)")
        if not path:
            return
        try:
            self.reference_df = self._read_table(path)
            self.reference_path.setText(path)
            self._populate_combo(self.ref_check_combo, list(self.reference_df.columns), CHECK_COLUMN_CANDIDATES)
            self._populate_combo(self.ref_vendor_combo, list(self.reference_df.columns), VENDOR_COLUMN_CANDIDATES)
            self._update_summary("Loaded reference file.")
            self._set_status("Reference file loaded")
        except Exception as exc:
            self._error(f"Could not read reference file: {exc}")

    def _populate_combo(self, combo: QComboBox, columns: List[str], candidates: List[str]) -> None:
        combo.clear()
        combo.addItems(columns)
        lower = {c.lower().strip(): c for c in columns}
        for candidate in candidates:
            for key, original in lower.items():
                if candidate in key:
                    combo.setCurrentText(original)
                    return

    def _required_mapping(self) -> Tuple[str, str, str, str]:
        if self.quickbooks_df is None or self.reference_df is None:
            raise ValueError("Please load both CSV files first.")

        qb_check = self.qb_check_combo.currentText().strip()
        qb_vendor = self.qb_vendor_combo.currentText().strip()
        ref_check = self.ref_check_combo.currentText().strip()
        ref_vendor = self.ref_vendor_combo.currentText().strip()

        if not all([qb_check, qb_vendor, ref_check, ref_vendor]):
            raise ValueError("Please map all required columns before processing.")

        return qb_check, qb_vendor, ref_check, ref_vendor

    def process_updates(self) -> None:
        try:
            self._set_status("Processing...")
            qb_check, qb_vendor, ref_check, ref_vendor = self._required_mapping()

            normalize_mode = self.normalize_checkbox.isChecked()
            extract_from_text_mode = self.extract_checkbox.isChecked()
            qb_df = self.quickbooks_df.copy() if self.quickbooks_df is not None else None
            ref_df = self.reference_df.copy() if self.reference_df is not None else None
            if qb_df is None or ref_df is None:
                raise ValueError("Missing files.")

            ref_df["_check_key"] = ref_df[ref_check].apply(
                lambda v: normalize_check_number(v, normalize_mode, extract_from_text_mode)
            )
            ref_df["_vendor_value"] = ref_df[ref_vendor].fillna("").astype(str).str.strip()

            duplicate_counts = ref_df[ref_df["_check_key"] != ""]["_check_key"].value_counts()
            self.duplicates = {k: int(v) for k, v in duplicate_counts[duplicate_counts > 1].to_dict().items()}

            lookup_series = (
                ref_df[ref_df["_check_key"] != ""]
                .drop_duplicates(subset=["_check_key"], keep="first")
                .set_index("_check_key")["_vendor_value"]
            )

            qb_df["_check_key"] = qb_df[qb_check].apply(
                lambda v: normalize_check_number(v, normalize_mode, extract_from_text_mode)
            )
            qb_df["_existing_vendor"] = qb_df[qb_vendor].fillna("").astype(str)
            qb_df["_matched_vendor"] = qb_df["_check_key"].map(lookup_series)
            qb_df["_is_match"] = qb_df["_matched_vendor"].notna()

            qb_df.loc[qb_df["_is_match"], qb_vendor] = qb_df.loc[qb_df["_is_match"], "_matched_vendor"]

            total_rows = len(qb_df)
            matched_rows = int(qb_df["_is_match"].sum())
            unmatched_rows = total_rows - matched_rows
            replaced_rows = int((qb_df["_is_match"] & (qb_df["_existing_vendor"] != qb_df[qb_vendor].astype(str))).sum())
            empty_check_values = int((qb_df["_check_key"] == "").sum())

            self.unmatched_df = qb_df.loc[~qb_df["_is_match"], [qb_check, qb_vendor]].copy()
            self.updated_df = qb_df.drop(columns=["_check_key", "_existing_vendor", "_matched_vendor", "_is_match"])
            self.save_btn.setEnabled(True)
            self._render_preview(self.updated_df)

            msg_lines = [
                "Update complete.",
                f"Total rows in QuickBooks file: {total_rows}",
                f"Total check rows matched: {matched_rows}",
                f"Total unmatched rows: {unmatched_rows}",
                f"Total vendor names replaced: {replaced_rows}",
                f"Rows skipped due to blank/invalid check number: {empty_check_values}",
            ]
            if self.duplicates:
                msg_lines.append(
                    f"Warning: {len(self.duplicates)} duplicate check number(s) in reference file. First match was used."
                )

            self.summary_box.setPlainText("\n".join(msg_lines))
            self._set_status("Update complete")

            if self.duplicates:
                QMessageBox.warning(
                    self,
                    "Duplicate Check Numbers",
                    "Duplicate check numbers were found in the reference file. The first occurrence was used.",
                )

        except Exception as exc:
            self._error(str(exc))

    def save_updated_csv(self) -> None:
        if self.updated_df is None:
            self._error("No updated data to save. Run the update first.")
            return

        qb_text = self.quickbooks_path.text().strip()
        if qb_text:
            default_dir = str(Path(qb_text).parent)
        else:
            default_dir = ""

        default_name = str(Path(default_dir) / "QuickBooks_Upload_Updated.csv")
        path, _ = QFileDialog.getSaveFileName(self, "Save Updated QuickBooks File", default_name, "CSV Files (*.csv);;Excel Files (*.xlsx)")
        if not path:
            return

        try:
            backup_path: Optional[Path] = None
            if self.backup_checkbox.isChecked() and qb_text:
                source_path = Path(qb_text)
                if source_path.exists():
                    backup_path = source_path.with_name(source_path.stem + "_Backup" + source_path.suffix)
                    shutil.copy2(source_path, backup_path)

            self._write_table(self.updated_df, path)
            info = [f"Saved updated file: {path}"]

            if backup_path is not None and backup_path.exists():
                info.append(f"Saved backup of original QuickBooks file: {backup_path}")

            if self.unmatched_checkbox.isChecked() and self.unmatched_df is not None and not self.unmatched_df.empty:
                unmatched_ext = Path(path).suffix if Path(path).suffix.lower() in {".csv", ".xlsx"} else ".csv"
                unmatched_path = str(Path(path).with_name(Path(path).stem + "_Unmatched" + unmatched_ext))
                self._write_table(self.unmatched_df, unmatched_path)
                info.append(f"Saved unmatched report: {unmatched_path}")

            QMessageBox.information(self, "Saved", "\n".join(info))
            self._update_summary("\n".join(info))
            self._set_status("Files saved")
        except Exception as exc:
            self._error(f"Failed to save file: {exc}")

    def show_matching_help(self) -> None:
        QToolTip.showText(
            self.mapToGlobal(self.rect().center()),
            (
                "Matching uses the selected check-number columns from both files.\n"
                "- Normalize: trims spaces and removes trailing '.0'.\n"
                "- Extract: reads numbers from values like 'Check 101' or 'CHK-101'.\n"
                "Turn these off for strict exact text matching."
            ),
            self,
            self.rect(),
            8000,
        )

    def _render_preview(self, df: pd.DataFrame) -> None:
        preview = df.head(100)
        self.preview_table.setRowCount(len(preview))
        self.preview_table.setColumnCount(len(preview.columns))
        self.preview_table.setHorizontalHeaderLabels([str(c) for c in preview.columns])

        for row_idx in range(len(preview)):
            for col_idx in range(len(preview.columns)):
                value = "" if pd.isna(preview.iloc[row_idx, col_idx]) else str(preview.iloc[row_idx, col_idx])
                self.preview_table.setItem(row_idx, col_idx, QTableWidgetItem(value))

        self.preview_table.resizeColumnsToContents()

    def _error(self, message: str) -> None:
        QMessageBox.critical(self, "Error", message)
        self._update_summary(f"Error: {message}")
        self._set_status("Error")

    def _update_summary(self, text: str) -> None:
        current = self.summary_box.toPlainText().strip()
        if current:
            self.summary_box.setPlainText(current + "\n" + text)
        else:
            self.summary_box.setPlainText(text)


def main() -> None:
    app = QApplication(sys.argv)

    splash = SplashScreen()
    window = CheckVendorUpdater()

    def show_main_window() -> None:
        splash.close()
        login = LoginDialog()
        if login.exec() == QDialog.Accepted:
            window.show()
            return

        app.quit()

    splash.show()
    QTimer.singleShot(2000, show_main_window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
