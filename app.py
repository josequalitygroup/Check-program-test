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
    QDialog,
    QFileDialog,
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

TRANSLATIONS: Dict[str, Dict[str, str]] = {
    "en": {
        "lang_english": "English",
        "lang_spanish": "Spanish",
        "login_window": "Login",
        "login_sign_in": "Sign in",
        "login_user": "User:",
        "login_password": "Password:",
        "login_button": "Login",
        "login_invalid": "Invalid username or password.",
        "language": "Language:",
        "main_window": "QuickBooks Check Vendor Updater",
        "subtitle": "Load both files, confirm column mapping, then update vendor/payee names by check number.",
        "step1": "Step 1 — Select Files",
        "qb_placeholder": "Choose the QuickBooks Upload file (CSV/Excel)...",
        "ref_placeholder": "Choose one or more Check Reference files (CSV/Excel)...",
        "qb_browse": "Browse QuickBooks File",
        "ref_browse": "Browse Reference File(s)",
        "qb_target": "QuickBooks Upload file (target):",
        "ref_lookup": "Check Reference file(s) (lookup):",
        "step2": "Step 2 — Confirm Column Mapping",
        "map_qb_check": "QuickBooks Check Number column:",
        "map_qb_vendor": "QuickBooks Vendor/Payee to update:",
        "map_ref_check": "Reference Check Number column:",
        "map_ref_vendor": "Reference Vendor Name column:",
        "mapping_hint": "Tip: Common check columns include Check Number, Num, Ref No, Document No, or values like 'Check 101'.",
        "step3": "Step 3 — Matching Options",
        "normalize": "Normalize check values (trim + remove .0)",
        "extract": "Extract number from text (e.g., 'Check 101')",
        "unmatched": "Export unmatched rows report",
        "backup": "Create backup of original QuickBooks file on save",
        "step4": "Step 4 — Update Vendor Names",
        "step5": "Step 5 — Save Updated File",
        "reset": "Reset",
        "how_matching": "How matching works",
        "summary": "Summary",
        "preview": "Preview (first 100 rows)",
        "summary_placeholder": "Summary will appear here after processing...",
        "status_prefix": "Status",
        "ready": "Ready",
        "processing": "Processing...",
        "files_saved": "Files saved",
        "update_complete": "Update complete",
        "loaded_qb": "Loaded QuickBooks file.",
        "loaded_ref": "Loaded reference file(s): {count}.",
        "qb_loaded_status": "QuickBooks file loaded",
        "ref_loaded_status": "Reference file loaded",
        "select_qb": "Select QuickBooks Upload File",
        "select_ref": "Select Check Reference File(s)",
        "read_qb_err": "Could not read QuickBooks file: {error}",
        "read_ref_err": "Could not read reference file: {error}",
        "missing_files": "Please load both files first.",
        "missing_mapping": "Please map all required columns before processing.",
        "missing_files_runtime": "Missing files.",
        "saved_title": "Saved",
        "save_dialog": "Save Updated QuickBooks File",
        "saved_file": "Saved updated file: {path}",
        "saved_backup": "Saved backup of original QuickBooks file: {path}",
        "saved_unmatched": "Saved unmatched report: {path}",
        "save_err": "Failed to save file: {error}",
        "no_updated": "No updated data to save. Run the update first.",
        "error_title": "Error",
        "duplicate_title": "Duplicate Check Numbers",
        "duplicate_msg": "Duplicate check numbers were found in the reference file. The first occurrence was used.",
        "summary_done": "Update complete.",
        "summary_total": "Total rows in QuickBooks file: {value}",
        "summary_matched": "Total check rows matched: {value}",
        "summary_unmatched": "Total unmatched rows: {value}",
        "summary_replaced": "Total vendor names replaced: {value}",
        "summary_skipped": "Rows skipped due to blank/invalid check number: {value}",
        "summary_duplicate_warn": "Warning: {value} duplicate check number(s) in reference file. First match was used.",
        "help_text": "Matching uses the selected check-number columns from both files.\n- Normalize: trims spaces and removes trailing '.0'.\n- Extract: reads numbers from values like 'Check 101' or 'CHK-101'.\nTurn these off for strict exact text matching.",
        "unsupported_input": "Unsupported file type. Please select CSV or Excel (.xlsx).",
        "unsupported_output": "Unsupported output format. Please save as .csv or .xlsx.",
    },
    "es": {
        "lang_english": "Inglés",
        "lang_spanish": "Español",
        "login_window": "Inicio de sesión",
        "login_sign_in": "Iniciar sesión",
        "login_user": "Usuario:",
        "login_password": "Contraseña:",
        "login_button": "Entrar",
        "login_invalid": "Usuario o contraseña inválidos.",
        "language": "Idioma:",
        "main_window": "Actualizador de Proveedores de Cheques QuickBooks",
        "subtitle": "Cargue ambos archivos, confirme el mapeo de columnas y actualice proveedor/beneficiario por número de cheque.",
        "step1": "Paso 1 — Seleccionar archivos",
        "qb_placeholder": "Elija el archivo de QuickBooks (CSV/Excel)...",
        "ref_placeholder": "Elija uno o más archivos de referencia de cheques (CSV/Excel)...",
        "qb_browse": "Buscar archivo QuickBooks",
        "ref_browse": "Buscar archivo(s) de referencia",
        "qb_target": "Archivo QuickBooks (objetivo):",
        "ref_lookup": "Archivo(s) de referencia (búsqueda):",
        "step2": "Paso 2 — Confirmar mapeo de columnas",
        "map_qb_check": "Columna de número de cheque en QuickBooks:",
        "map_qb_vendor": "Columna de proveedor/beneficiario a actualizar:",
        "map_ref_check": "Columna de número de cheque en referencia:",
        "map_ref_vendor": "Columna de nombre de proveedor en referencia:",
        "mapping_hint": "Sugerencia: columnas comunes incluyen Check Number, Num, Ref No, Document No, o valores como 'Check 101'.",
        "step3": "Paso 3 — Opciones de coincidencia",
        "normalize": "Normalizar valores de cheque (recortar + quitar .0)",
        "extract": "Extraer número desde texto (ej., 'Check 101')",
        "unmatched": "Exportar reporte de no coincidentes",
        "backup": "Crear copia de respaldo del archivo QuickBooks al guardar",
        "step4": "Paso 4 — Actualizar nombres de proveedor",
        "step5": "Paso 5 — Guardar archivo actualizado",
        "reset": "Restablecer",
        "how_matching": "Cómo funciona la coincidencia",
        "summary": "Resumen",
        "preview": "Vista previa (primeras 100 filas)",
        "summary_placeholder": "El resumen aparecerá aquí después de procesar...",
        "status_prefix": "Estado",
        "ready": "Listo",
        "processing": "Procesando...",
        "files_saved": "Archivos guardados",
        "update_complete": "Actualización completada",
        "loaded_qb": "Archivo QuickBooks cargado.",
        "loaded_ref": "Archivo(s) de referencia cargado(s): {count}.",
        "qb_loaded_status": "Archivo QuickBooks cargado",
        "ref_loaded_status": "Archivo de referencia cargado",
        "select_qb": "Seleccionar archivo de QuickBooks",
        "select_ref": "Seleccionar archivo(s) de referencia",
        "read_qb_err": "No se pudo leer el archivo de QuickBooks: {error}",
        "read_ref_err": "No se pudo leer el archivo de referencia: {error}",
        "missing_files": "Por favor cargue ambos archivos primero.",
        "missing_mapping": "Por favor asigne todas las columnas requeridas antes de procesar.",
        "missing_files_runtime": "Faltan archivos.",
        "saved_title": "Guardado",
        "save_dialog": "Guardar archivo actualizado de QuickBooks",
        "saved_file": "Archivo actualizado guardado: {path}",
        "saved_backup": "Respaldo del archivo QuickBooks guardado: {path}",
        "saved_unmatched": "Reporte de no coincidentes guardado: {path}",
        "save_err": "Error al guardar archivo: {error}",
        "no_updated": "No hay datos actualizados para guardar. Ejecute la actualización primero.",
        "error_title": "Error",
        "duplicate_title": "Números de cheque duplicados",
        "duplicate_msg": "Se encontraron números de cheque duplicados en el archivo de referencia. Se usó la primera ocurrencia.",
        "summary_done": "Actualización completada.",
        "summary_total": "Total de filas en QuickBooks: {value}",
        "summary_matched": "Total de filas coincidentes: {value}",
        "summary_unmatched": "Total de filas sin coincidencia: {value}",
        "summary_replaced": "Total de nombres reemplazados: {value}",
        "summary_skipped": "Filas omitidas por número de cheque vacío/inválido: {value}",
        "summary_duplicate_warn": "Advertencia: {value} número(s) de cheque duplicado(s) en referencia. Se usó la primera coincidencia.",
        "help_text": "La coincidencia usa las columnas de número de cheque seleccionadas en ambos archivos.\n- Normalizar: recorta espacios y quita '.0'.\n- Extraer: toma números de valores como 'Check 101' o 'CHK-101'.\nDesactive estas opciones para coincidencia estricta por texto.",
        "unsupported_input": "Tipo de archivo no compatible. Seleccione CSV o Excel (.xlsx).",
        "unsupported_output": "Formato de salida no compatible. Guarde como .csv o .xlsx.",
    },
}


def normalize_check_number(value: object, normalize_mode: bool = True, extract_from_text_mode: bool = True) -> str:
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
    def __init__(self, language: str = "en") -> None:
        super().__init__()
        self.language = language if language in TRANSLATIONS else "en"
        self.setModal(True)
        self.setFixedSize(390, 260)

        layout = QVBoxLayout(self)
        self.title_label = QLabel()
        self.title_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #f2c14e;")

        form = QFormLayout()
        self.language_label = QLabel()
        self.language_combo = QComboBox()
        self.language_combo.addItem(TRANSLATIONS[self.language]["lang_english"], "en")
        self.language_combo.addItem(TRANSLATIONS[self.language]["lang_spanish"], "es")
        self.language_combo.setCurrentIndex(0 if self.language == "en" else 1)
        self.language_combo.currentIndexChanged.connect(self._on_language_changed)

        self.user_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)

        self.user_label = QLabel()
        self.password_label = QLabel()
        form.addRow(self.language_label, self.language_combo)
        form.addRow(self.user_label, self.user_input)
        form.addRow(self.password_label, self.password_input)

        self.login_btn = QPushButton()
        self.login_btn.clicked.connect(self.try_login)

        self.message_label = QLabel("")
        self.message_label.setStyleSheet("color: #ffb3b3;")

        layout.addWidget(self.title_label)
        layout.addLayout(form)
        layout.addWidget(self.message_label)
        layout.addWidget(self.login_btn)

        self.setStyleSheet(
            "QDialog { background: #111111; color: #f5f5f5; }"
            "QLineEdit, QComboBox { background: #0f0f10; color: #f5f5f5; border: 1px solid #c29b2d; border-radius: 6px; padding: 6px; }"
            "QPushButton { background: #1c1c1d; color: #fff8e6; border: 1px solid #d4af37; border-radius: 6px; min-height: 30px; font-weight: 600; }"
            "QPushButton:hover { background: #2a2a2c; }"
        )
        self._apply_texts()

    def tr(self, key: str, **kwargs: object) -> str:
        text = TRANSLATIONS[self.language][key]
        return text.format(**kwargs) if kwargs else text

    def _on_language_changed(self) -> None:
        self.language = str(self.language_combo.currentData())
        self._apply_texts()

    def _apply_texts(self) -> None:
        self.setWindowTitle(self.tr("login_window"))
        self.title_label.setText(self.tr("login_sign_in"))
        self.language_label.setText(self.tr("language"))
        self.user_label.setText(self.tr("login_user"))
        self.password_label.setText(self.tr("login_password"))
        self.login_btn.setText(self.tr("login_button"))

    def try_login(self) -> None:
        if self.user_input.text().strip() == "Kiri" and self.password_input.text() == "Jcr16331878":
            self.accept()
            return
        self.message_label.setText(self.tr("login_invalid"))


class CheckVendorUpdater(QMainWindow):
    def __init__(self, language: str = "en") -> None:
        super().__init__()
        self.language = language if language in TRANSLATIONS else "en"
        self.resize(1240, 800)

        self.quickbooks_df: Optional[pd.DataFrame] = None
        self.reference_df: Optional[pd.DataFrame] = None
        self.reference_files: List[str] = []
        self.updated_df: Optional[pd.DataFrame] = None
        self.unmatched_df: Optional[pd.DataFrame] = None
        self.duplicates: Dict[str, int] = {}

        self._build_ui()
        self._apply_styles()
        self._apply_texts()

    def tr(self, key: str, **kwargs: object) -> str:
        text = TRANSLATIONS[self.language][key]
        return text.format(**kwargs) if kwargs else text

    def _build_ui(self) -> None:
        container = QWidget(self)
        self.setCentralWidget(container)
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(20, 18, 20, 18)
        root_layout.setSpacing(12)

        top_row = QHBoxLayout()
        self.title_label = QLabel()
        self.title_label.setObjectName("mainTitle")
        top_row.addWidget(self.title_label)
        top_row.addStretch()
        self.language_label = QLabel()
        self.language_combo = QComboBox()
        self.language_combo.addItem("English", "en")
        self.language_combo.addItem("Español", "es")
        self.language_combo.setCurrentIndex(0 if self.language == "en" else 1)
        self.language_combo.currentIndexChanged.connect(self._on_language_changed)
        top_row.addWidget(self.language_label)
        top_row.addWidget(self.language_combo)

        self.subtitle_label = QLabel()
        self.subtitle_label.setWordWrap(True)
        self.subtitle_label.setObjectName("subTitle")
        root_layout.addLayout(top_row)
        root_layout.addWidget(self.subtitle_label)

        self.file_group = QGroupBox()
        self.file_group.setObjectName("panel")
        file_layout = QGridLayout(self.file_group)

        self.quickbooks_path = QLineEdit()
        self.quickbooks_path.setReadOnly(True)
        self.qb_btn = QPushButton()
        self.qb_btn.clicked.connect(self.load_quickbooks_csv)

        self.reference_path = QLineEdit()
        self.reference_path.setReadOnly(True)
        self.ref_btn = QPushButton()
        self.ref_btn.clicked.connect(self.load_reference_csv)

        self.qb_target_label = QLabel()
        self.ref_lookup_label = QLabel()

        file_layout.addWidget(self.qb_target_label, 0, 0)
        file_layout.addWidget(self.quickbooks_path, 0, 1)
        file_layout.addWidget(self.qb_btn, 0, 2)

        file_layout.addWidget(self.ref_lookup_label, 1, 0)
        file_layout.addWidget(self.reference_path, 1, 1)
        file_layout.addWidget(self.ref_btn, 1, 2)

        self.mapping_group = QGroupBox()
        self.mapping_group.setObjectName("panel")
        mapping_layout = QFormLayout(self.mapping_group)

        self.qb_check_combo = QComboBox()
        self.qb_vendor_combo = QComboBox()
        self.ref_check_combo = QComboBox()
        self.ref_vendor_combo = QComboBox()

        self.map_qb_check_label = QLabel()
        self.map_qb_vendor_label = QLabel()
        self.map_ref_check_label = QLabel()
        self.map_ref_vendor_label = QLabel()

        mapping_layout.addRow(self.map_qb_check_label, self.qb_check_combo)
        mapping_layout.addRow(self.map_qb_vendor_label, self.qb_vendor_combo)
        mapping_layout.addRow(self.map_ref_check_label, self.ref_check_combo)
        mapping_layout.addRow(self.map_ref_vendor_label, self.ref_vendor_combo)

        self.mapping_hint = QLabel()
        self.mapping_hint.setWordWrap(True)

        self.options_group = QGroupBox()
        self.options_group.setObjectName("panel")
        options_layout = QHBoxLayout(self.options_group)
        self.normalize_checkbox = QCheckBox()
        self.normalize_checkbox.setChecked(True)
        self.extract_checkbox = QCheckBox()
        self.extract_checkbox.setChecked(True)
        self.unmatched_checkbox = QCheckBox()
        self.unmatched_checkbox.setChecked(True)
        self.backup_checkbox = QCheckBox()
        self.backup_checkbox.setChecked(True)

        options_layout.addWidget(self.normalize_checkbox)
        options_layout.addWidget(self.extract_checkbox)
        options_layout.addWidget(self.unmatched_checkbox)
        options_layout.addWidget(self.backup_checkbox)
        options_layout.addStretch()

        actions_layout = QHBoxLayout()
        self.process_btn = QPushButton()
        self.process_btn.clicked.connect(self.process_updates)
        self.save_btn = QPushButton()
        self.save_btn.clicked.connect(self.save_updated_csv)
        self.save_btn.setEnabled(False)
        self.reset_btn = QPushButton()
        self.reset_btn.clicked.connect(self.reset_app)
        self.help_btn = QPushButton()
        self.help_btn.clicked.connect(self.show_matching_help)

        actions_layout.addWidget(self.process_btn)
        actions_layout.addWidget(self.save_btn)
        actions_layout.addWidget(self.help_btn)
        actions_layout.addWidget(self.reset_btn)
        actions_layout.addStretch()

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")

        self.summary_title_label = QLabel()
        self.summary_box = QTextEdit()
        self.summary_box.setObjectName("summaryBox")
        self.summary_box.setReadOnly(True)
        self.summary_box.setFixedHeight(130)

        self.preview_title_label = QLabel()
        self.preview_table = QTableWidget()
        self.preview_table.setObjectName("previewTable")
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

        root_layout.addWidget(self.file_group)
        root_layout.addWidget(self.mapping_group)
        root_layout.addWidget(self.mapping_hint)
        root_layout.addWidget(self.options_group)
        root_layout.addLayout(actions_layout)
        root_layout.addWidget(self.status_label)
        root_layout.addWidget(self.summary_title_label)
        root_layout.addWidget(self.summary_box)
        root_layout.addWidget(self.preview_title_label)
        root_layout.addWidget(self.preview_table)

    def _apply_texts(self) -> None:
        self.setWindowTitle(self.tr("main_window"))
        self.title_label.setText(self.tr("main_window"))
        self.subtitle_label.setText(self.tr("subtitle"))
        self.language_label.setText(self.tr("language"))
        self.language_combo.setItemText(0, self.tr("lang_english"))
        self.language_combo.setItemText(1, self.tr("lang_spanish"))

        self.file_group.setTitle(self.tr("step1"))
        self.quickbooks_path.setPlaceholderText(self.tr("qb_placeholder"))
        self.reference_path.setPlaceholderText(self.tr("ref_placeholder"))
        self.qb_btn.setText(self.tr("qb_browse"))
        self.ref_btn.setText(self.tr("ref_browse"))
        self.qb_target_label.setText(self.tr("qb_target"))
        self.ref_lookup_label.setText(self.tr("ref_lookup"))

        self.mapping_group.setTitle(self.tr("step2"))
        self.map_qb_check_label.setText(self.tr("map_qb_check"))
        self.map_qb_vendor_label.setText(self.tr("map_qb_vendor"))
        self.map_ref_check_label.setText(self.tr("map_ref_check"))
        self.map_ref_vendor_label.setText(self.tr("map_ref_vendor"))
        self.mapping_hint.setText(self.tr("mapping_hint"))

        self.options_group.setTitle(self.tr("step3"))
        self.normalize_checkbox.setText(self.tr("normalize"))
        self.extract_checkbox.setText(self.tr("extract"))
        self.unmatched_checkbox.setText(self.tr("unmatched"))
        self.backup_checkbox.setText(self.tr("backup"))

        self.process_btn.setText(self.tr("step4"))
        self.save_btn.setText(self.tr("step5"))
        self.help_btn.setText(self.tr("how_matching"))
        self.reset_btn.setText(self.tr("reset"))
        self.summary_title_label.setText(self.tr("summary"))
        self.preview_title_label.setText(self.tr("preview"))
        self.summary_box.setPlaceholderText(self.tr("summary_placeholder"))
        self._set_status(self.tr("ready"))

    def _on_language_changed(self) -> None:
        self.language = str(self.language_combo.currentData())
        self._apply_texts()

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #0f0f10; color: #f5f5f5; }
            QLabel { color: #f5f5f5; }
            #mainTitle { font-size: 30px; font-weight: 800; color: #f2c14e; letter-spacing: 0.5px; }
            #subTitle { color: #f3e8ca; margin-bottom: 8px; font-size: 13px; }
            #statusLabel { font-weight: 700; color: #fff8e6; background: #1f1f20; border: 1px solid #d4af37; border-radius: 8px; padding: 8px 10px; }
            QGroupBox#panel { font-weight: 700; border: 1px solid #d4af37; border-radius: 12px; margin-top: 10px; background: #151516; padding: 10px; color: #f5f5f5; }
            QGroupBox#panel::title { subcontrol-origin: margin; left: 10px; padding: 0 6px; color: #f2c14e; }
            QPushButton { min-height: 34px; padding: 6px 12px; border-radius: 8px; border: 1px solid #d4af37; background: #1c1c1d; color: #fff8e6; font-weight: 600; }
            QPushButton:hover { background: #2a2a2c; }
            QPushButton:pressed { background: #111111; }
            QPushButton:disabled { color: #8b8b8b; border: 1px solid #545454; background: #1a1a1a; }
            QLineEdit, QComboBox, QTextEdit { background: #0f0f10; color: #f5f5f5; border: 1px solid #c29b2d; border-radius: 8px; padding: 6px; selection-background-color: #d4af37; selection-color: #111111; }
            QCheckBox { color: #f5f5f5; spacing: 8px; }
            QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #c29b2d; border-radius: 4px; background: #111111; }
            QCheckBox::indicator:checked { background: #d4af37; }
            #summaryBox { background: #101011; color: #f5f5f5; }
            #previewTable { gridline-color: #3c3c3f; alternate-background-color: #171718; background: #101011; color: #f5f5f5; border: 1px solid #c29b2d; border-radius: 8px; }
            QHeaderView::section { background: #d4af37; color: #151516; padding: 6px; border: 0; border-right: 1px solid #8a6b1d; border-bottom: 1px solid #8a6b1d; font-weight: 700; }
            QToolTip { background-color: #111111; color: #f5f5f5; border: 1px solid #d4af37; padding: 6px; }
            """
        )

    def _set_status(self, text: str) -> None:
        self.status_label.setText(f"{self.tr('status_prefix')}: {text}")

    def reset_app(self) -> None:
        self.quickbooks_df = None
        self.reference_df = None
        self.reference_files = []
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
        self._set_status(self.tr("ready"))

    @staticmethod
    def _read_table(path: str, language: str) -> pd.DataFrame:
        suffix = Path(path).suffix.lower()
        if suffix == ".csv":
            return pd.read_csv(path, dtype=object)
        if suffix == ".xlsx":
            return pd.read_excel(path, dtype=object)
        raise ValueError(TRANSLATIONS[language]["unsupported_input"])

    @staticmethod
    def _write_table(df: pd.DataFrame, path: str, language: str) -> None:
        suffix = Path(path).suffix.lower()
        if suffix == ".csv":
            df.to_csv(path, index=False)
            return
        if suffix == ".xlsx":
            df.to_excel(path, index=False)
            return
        raise ValueError(TRANSLATIONS[language]["unsupported_output"])

    def load_quickbooks_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, self.tr("select_qb"), "", "Data Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx)")
        if not path:
            return
        try:
            self.quickbooks_df = self._read_table(path, self.language)
            self.quickbooks_path.setText(path)
            self._populate_combo(self.qb_check_combo, list(self.quickbooks_df.columns), CHECK_COLUMN_CANDIDATES)
            self._populate_combo(self.qb_vendor_combo, list(self.quickbooks_df.columns), VENDOR_COLUMN_CANDIDATES)
            self._update_summary(self.tr("loaded_qb"))
            self._set_status(self.tr("qb_loaded_status"))
        except Exception as exc:
            self._error(self.tr("read_qb_err", error=exc))

    def load_reference_csv(self) -> None:
        dialog = QFileDialog(self, self.tr("select_ref"))
        dialog.setFileMode(QFileDialog.ExistingFiles)
        dialog.setNameFilter("Data Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx)")
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)

        if not dialog.exec():
            return

        paths = dialog.selectedFiles()
        if not paths:
            return

        frames: List[pd.DataFrame] = []
        errors: List[str] = []
        for path in paths:
            try:
                frames.append(self._read_table(path, self.language))
            except Exception as exc:
                errors.append(f"{Path(path).name}: {exc}")

        if not frames:
            self._error(self.tr("read_ref_err", error="; ".join(errors)))
            return

        self.reference_df = pd.concat(frames, ignore_index=True, sort=False)
        self.reference_files = list(paths)
        self.reference_path.setText("; ".join(self.reference_files))
        self._populate_combo(self.ref_check_combo, list(self.reference_df.columns), CHECK_COLUMN_CANDIDATES)
        self._populate_combo(self.ref_vendor_combo, list(self.reference_df.columns), VENDOR_COLUMN_CANDIDATES)
        self._update_summary(self.tr("loaded_ref", count=len(self.reference_files)))
        self._set_status(self.tr("ref_loaded_status"))

        if errors:
            self._update_summary(self.tr("error_title") + ": " + "; ".join(errors))

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
            raise ValueError(self.tr("missing_files"))

        qb_check = self.qb_check_combo.currentText().strip()
        qb_vendor = self.qb_vendor_combo.currentText().strip()
        ref_check = self.ref_check_combo.currentText().strip()
        ref_vendor = self.ref_vendor_combo.currentText().strip()

        if not all([qb_check, qb_vendor, ref_check, ref_vendor]):
            raise ValueError(self.tr("missing_mapping"))
        return qb_check, qb_vendor, ref_check, ref_vendor

    def process_updates(self) -> None:
        try:
            self._set_status(self.tr("processing"))
            qb_check, qb_vendor, ref_check, ref_vendor = self._required_mapping()

            normalize_mode = self.normalize_checkbox.isChecked()
            extract_from_text_mode = self.extract_checkbox.isChecked()
            qb_df = self.quickbooks_df.copy() if self.quickbooks_df is not None else None
            ref_df = self.reference_df.copy() if self.reference_df is not None else None
            if qb_df is None or ref_df is None:
                raise ValueError(self.tr("missing_files_runtime"))

            ref_df["_check_key"] = ref_df[ref_check].apply(lambda v: normalize_check_number(v, normalize_mode, extract_from_text_mode))
            ref_df["_vendor_value"] = ref_df[ref_vendor].fillna("").astype(str).str.strip()

            duplicate_counts = ref_df[ref_df["_check_key"] != ""]["_check_key"].value_counts()
            self.duplicates = {k: int(v) for k, v in duplicate_counts[duplicate_counts > 1].to_dict().items()}

            lookup_series = ref_df[ref_df["_check_key"] != ""].drop_duplicates(subset=["_check_key"], keep="first").set_index("_check_key")["_vendor_value"]

            qb_df["_check_key"] = qb_df[qb_check].apply(lambda v: normalize_check_number(v, normalize_mode, extract_from_text_mode))
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
                self.tr("summary_done"),
                self.tr("summary_total", value=total_rows),
                self.tr("summary_matched", value=matched_rows),
                self.tr("summary_unmatched", value=unmatched_rows),
                self.tr("summary_replaced", value=replaced_rows),
                self.tr("summary_skipped", value=empty_check_values),
            ]
            if self.duplicates:
                msg_lines.append(self.tr("summary_duplicate_warn", value=len(self.duplicates)))

            self.summary_box.setPlainText("\n".join(msg_lines))
            self._set_status(self.tr("update_complete"))

            if self.duplicates:
                QMessageBox.warning(self, self.tr("duplicate_title"), self.tr("duplicate_msg"))

        except Exception as exc:
            self._error(str(exc))

    def save_updated_csv(self) -> None:
        if self.updated_df is None:
            self._error(self.tr("no_updated"))
            return

        qb_text = self.quickbooks_path.text().strip()
        default_dir = str(Path(qb_text).parent) if qb_text else ""
        default_name = str(Path(default_dir) / "QuickBooks_Upload_Updated.csv")
        path, _ = QFileDialog.getSaveFileName(self, self.tr("save_dialog"), default_name, "CSV Files (*.csv);;Excel Files (*.xlsx)")
        if not path:
            return

        try:
            backup_path: Optional[Path] = None
            if self.backup_checkbox.isChecked() and qb_text:
                source_path = Path(qb_text)
                if source_path.exists():
                    backup_path = source_path.with_name(source_path.stem + "_Backup" + source_path.suffix)
                    shutil.copy2(source_path, backup_path)

            self._write_table(self.updated_df, path, self.language)
            info = [self.tr("saved_file", path=path)]

            if backup_path is not None and backup_path.exists():
                info.append(self.tr("saved_backup", path=backup_path))

            if self.unmatched_checkbox.isChecked() and self.unmatched_df is not None and not self.unmatched_df.empty:
                unmatched_ext = Path(path).suffix if Path(path).suffix.lower() in {".csv", ".xlsx"} else ".csv"
                unmatched_path = str(Path(path).with_name(Path(path).stem + "_Unmatched" + unmatched_ext))
                self._write_table(self.unmatched_df, unmatched_path, self.language)
                info.append(self.tr("saved_unmatched", path=unmatched_path))

            QMessageBox.information(self, self.tr("saved_title"), "\n".join(info))
            self._update_summary("\n".join(info))
            self._set_status(self.tr("files_saved"))
        except Exception as exc:
            self._error(self.tr("save_err", error=exc))

    def show_matching_help(self) -> None:
        QToolTip.showText(self.mapToGlobal(self.rect().center()), self.tr("help_text"), self, self.rect(), 8000)

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
        QMessageBox.critical(self, self.tr("error_title"), message)
        self._update_summary(f"{self.tr('error_title')}: {message}")
        self._set_status(self.tr("error_title"))

    def _update_summary(self, text: str) -> None:
        current = self.summary_box.toPlainText().strip()
        self.summary_box.setPlainText(current + "\n" + text if current else text)


def main() -> None:
    app = QApplication(sys.argv)

    splash = SplashScreen()
    window = CheckVendorUpdater(language="en")

    def show_main_window() -> None:
        splash.close()
        login = LoginDialog(language="en")
        if login.exec() == QDialog.Accepted:
            window.language = login.language
            window.language_combo.setCurrentIndex(0 if window.language == "en" else 1)
            window._apply_texts()
            window.show()
            return
        app.quit()

    splash.show()
    QTimer.singleShot(2000, show_main_window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
