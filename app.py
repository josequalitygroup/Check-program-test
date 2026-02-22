from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
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
    "check",
]

VENDOR_COLUMN_CANDIDATES = ["vendor", "payee", "name", "vendor name"]


def normalize_check_number(value: object, normalize_mode: bool = True) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not normalize_mode:
        return text

    if text.endswith(".0"):
        text = text[:-2]
    return text.strip()


class CheckVendorUpdater(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("QuickBooks Check Vendor Updater")
        self.resize(1200, 760)

        self.quickbooks_df: Optional[pd.DataFrame] = None
        self.reference_df: Optional[pd.DataFrame] = None
        self.updated_df: Optional[pd.DataFrame] = None
        self.unmatched_df: Optional[pd.DataFrame] = None
        self.duplicates: Dict[str, int] = {}

        self._build_ui()

    def _build_ui(self) -> None:
        container = QWidget(self)
        self.setCentralWidget(container)
        root_layout = QVBoxLayout(container)

        file_group = QGroupBox("1) Select Files")
        file_layout = QGridLayout(file_group)

        self.quickbooks_path = QLineEdit()
        self.quickbooks_path.setReadOnly(True)
        qb_btn = QPushButton("Browse QuickBooks Upload CSV")
        qb_btn.clicked.connect(self.load_quickbooks_csv)

        self.reference_path = QLineEdit()
        self.reference_path.setReadOnly(True)
        ref_btn = QPushButton("Browse Check Reference CSV")
        ref_btn.clicked.connect(self.load_reference_csv)

        file_layout.addWidget(QLabel("QuickBooks Upload CSV (target):"), 0, 0)
        file_layout.addWidget(self.quickbooks_path, 0, 1)
        file_layout.addWidget(qb_btn, 0, 2)

        file_layout.addWidget(QLabel("Check Reference CSV (lookup):"), 1, 0)
        file_layout.addWidget(self.reference_path, 1, 1)
        file_layout.addWidget(ref_btn, 1, 2)

        mapping_group = QGroupBox("2) Column Mapping")
        mapping_layout = QFormLayout(mapping_group)

        self.qb_check_combo = QComboBox()
        self.qb_vendor_combo = QComboBox()
        self.ref_check_combo = QComboBox()
        self.ref_vendor_combo = QComboBox()

        mapping_layout.addRow("QuickBooks Check Number Column:", self.qb_check_combo)
        mapping_layout.addRow("QuickBooks Vendor/Payee Column to Update:", self.qb_vendor_combo)
        mapping_layout.addRow("Reference Check Number Column:", self.ref_check_combo)
        mapping_layout.addRow("Reference Vendor Name Column:", self.ref_vendor_combo)

        options_layout = QHBoxLayout()
        self.normalize_checkbox = QCheckBox("Use normalized check number matching")
        self.normalize_checkbox.setChecked(True)
        self.unmatched_checkbox = QCheckBox("Export unmatched checks CSV")
        self.unmatched_checkbox.setChecked(True)
        options_layout.addWidget(self.normalize_checkbox)
        options_layout.addWidget(self.unmatched_checkbox)
        options_layout.addStretch()

        button_layout = QHBoxLayout()
        self.process_btn = QPushButton("3) Update Vendor Names")
        self.process_btn.clicked.connect(self.process_updates)
        self.save_btn = QPushButton("4) Save Updated CSV")
        self.save_btn.clicked.connect(self.save_updated_csv)
        self.save_btn.setEnabled(False)
        button_layout.addWidget(self.process_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addStretch()

        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setFixedHeight(130)

        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)

        root_layout.addWidget(file_group)
        root_layout.addWidget(mapping_group)
        root_layout.addLayout(options_layout)
        root_layout.addLayout(button_layout)
        root_layout.addWidget(QLabel("Summary"))
        root_layout.addWidget(self.summary_box)
        root_layout.addWidget(QLabel("Preview (first 100 rows)"))
        root_layout.addWidget(self.preview_table)

    def load_quickbooks_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select QuickBooks Upload CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        try:
            self.quickbooks_df = pd.read_csv(path, dtype=object)
            self.quickbooks_path.setText(path)
            self._populate_combo(self.qb_check_combo, list(self.quickbooks_df.columns), CHECK_COLUMN_CANDIDATES)
            self._populate_combo(self.qb_vendor_combo, list(self.quickbooks_df.columns), VENDOR_COLUMN_CANDIDATES)
            self._update_summary("Loaded QuickBooks file.")
        except Exception as exc:
            self._error(f"Could not read QuickBooks CSV: {exc}")

    def load_reference_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Select Check Reference CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        try:
            self.reference_df = pd.read_csv(path, dtype=object)
            self.reference_path.setText(path)
            self._populate_combo(self.ref_check_combo, list(self.reference_df.columns), CHECK_COLUMN_CANDIDATES)
            self._populate_combo(self.ref_vendor_combo, list(self.reference_df.columns), VENDOR_COLUMN_CANDIDATES)
            self._update_summary("Loaded reference file.")
        except Exception as exc:
            self._error(f"Could not read reference CSV: {exc}")

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
            qb_check, qb_vendor, ref_check, ref_vendor = self._required_mapping()

            normalize_mode = self.normalize_checkbox.isChecked()
            qb_df = self.quickbooks_df.copy() if self.quickbooks_df is not None else None
            ref_df = self.reference_df.copy() if self.reference_df is not None else None
            if qb_df is None or ref_df is None:
                raise ValueError("Missing files.")

            ref_df["_check_key"] = ref_df[ref_check].apply(lambda v: normalize_check_number(v, normalize_mode))
            ref_df["_vendor_value"] = ref_df[ref_vendor].fillna("").astype(str).str.strip()

            duplicate_counts = ref_df[ref_df["_check_key"] != ""]["_check_key"].value_counts()
            self.duplicates = {k: int(v) for k, v in duplicate_counts[duplicate_counts > 1].to_dict().items()}

            lookup_series = (
                ref_df[ref_df["_check_key"] != ""]
                .drop_duplicates(subset=["_check_key"], keep="first")
                .set_index("_check_key")["_vendor_value"]
            )

            qb_df["_check_key"] = qb_df[qb_check].apply(lambda v: normalize_check_number(v, normalize_mode))
            qb_df["_existing_vendor"] = qb_df[qb_vendor].fillna("").astype(str)
            qb_df["_matched_vendor"] = qb_df["_check_key"].map(lookup_series)
            qb_df["_is_match"] = qb_df["_matched_vendor"].notna()

            qb_df.loc[qb_df["_is_match"], qb_vendor] = qb_df.loc[qb_df["_is_match"], "_matched_vendor"]

            total_rows = len(qb_df)
            matched_rows = int(qb_df["_is_match"].sum())
            unmatched_rows = total_rows - matched_rows
            replaced_rows = int((qb_df["_is_match"] & (qb_df["_existing_vendor"] != qb_df[qb_vendor].astype(str))).sum())

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
            ]
            if self.duplicates:
                msg_lines.append(
                    f"Warning: {len(self.duplicates)} duplicate check number(s) in reference file. First match was used."
                )

            self.summary_box.setPlainText("\n".join(msg_lines))

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

        default_dir = str(Path(self.quickbooks_path.text()).parent) if self.quickbooks_path.text() else ""
        default_name = str(Path(default_dir) / "QuickBooks_Upload_Updated.csv")
        path, _ = QFileDialog.getSaveFileName(self, "Save Updated QuickBooks CSV", default_name, "CSV Files (*.csv)")
        if not path:
            return

        try:
            self.updated_df.to_csv(path, index=False)
            info = [f"Saved updated CSV: {path}"]

            if self.unmatched_checkbox.isChecked() and self.unmatched_df is not None and not self.unmatched_df.empty:
                unmatched_path = str(Path(path).with_name(Path(path).stem + "_Unmatched.csv"))
                self.unmatched_df.to_csv(unmatched_path, index=False)
                info.append(f"Saved unmatched report: {unmatched_path}")

            QMessageBox.information(self, "Saved", "\n".join(info))
            self._update_summary("\n".join(info))
        except Exception as exc:
            self._error(f"Failed to save CSV: {exc}")

    def _render_preview(self, df: pd.DataFrame) -> None:
        preview = df.head(100)
        self.preview_table.setRowCount(len(preview))
        self.preview_table.setColumnCount(len(preview.columns))
        self.preview_table.setHorizontalHeaderLabels([str(c) for c in preview.columns])

        for row_idx in range(len(preview)):
            for col_idx, col_name in enumerate(preview.columns):
                value = "" if pd.isna(preview.iloc[row_idx, col_idx]) else str(preview.iloc[row_idx, col_idx])
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                self.preview_table.setItem(row_idx, col_idx, item)

        self.preview_table.resizeColumnsToContents()

    def _error(self, message: str) -> None:
        QMessageBox.critical(self, "Error", message)
        self._update_summary(f"Error: {message}")

    def _update_summary(self, text: str) -> None:
        current = self.summary_box.toPlainText().strip()
        if current:
            self.summary_box.setPlainText(current + "\n" + text)
        else:
            self.summary_box.setPlainText(text)


def main() -> None:
    app = QApplication(sys.argv)
    window = CheckVendorUpdater()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
