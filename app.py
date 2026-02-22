import base64
import calendar
import io
import json
import os
import threading
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
import pypdfium2 as pdfium
import tkinter as tk
from openai import OpenAI
from PIL import Image
from tkinter import filedialog, messagebox, ttk

ALLOWED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png"}


@dataclass
class CheckRow:
    check_number: str
    vendor_name: str
    month: int
    year: int
    source_file: str
    confidence: str = ""
    notes: str = ""


class CheckExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Check Extractor (OpenAI)")
        self.root.geometry("1150x680")

        self.client = self._build_client()
        self.rows: List[CheckRow] = []
        self.busy = False

        now = datetime.now()
        self.month_var = tk.StringVar(value=now.strftime("%B"))
        self.year_var = tk.StringVar(value=str(now.year))
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_client(self):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            return None
        return OpenAI(api_key=api_key)

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        self.upload_btn = ttk.Button(top, text="Upload Files", command=self.upload_files)
        self.upload_btn.pack(side="left", padx=5)
        ttk.Button(top, text="Apply Filter", command=self.refresh_table).pack(side="left", padx=5)
        ttk.Button(top, text="Export Filtered to Excel", command=self.export_excel).pack(side="left", padx=5)

        ttk.Label(top, text="Month:").pack(side="left", padx=(20, 5))
        months = list(calendar.month_name)[1:]
        self.month_combo = ttk.Combobox(top, textvariable=self.month_var, values=months, width=12, state="readonly")
        self.month_combo.pack(side="left")

        ttk.Label(top, text="Year:").pack(side="left", padx=(20, 5))
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 5, current_year + 6)]
        self.year_combo = ttk.Combobox(top, textvariable=self.year_var, values=years, width=8)
        self.year_combo.pack(side="left")

        middle = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        middle.pack(fill="both", expand=True)

        columns = ("check_number", "vendor_name", "month", "year", "source_file", "confidence", "notes")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings")
        self.tree.heading("check_number", text="Check Number")
        self.tree.heading("vendor_name", text="Vendor Name")
        self.tree.heading("month", text="Month")
        self.tree.heading("year", text="Year")
        self.tree.heading("source_file", text="Source File")
        self.tree.heading("confidence", text="Confidence")
        self.tree.heading("notes", text="Notes")

        self.tree.column("check_number", width=120)
        self.tree.column("vendor_name", width=220)
        self.tree.column("month", width=90)
        self.tree.column("year", width=80)
        self.tree.column("source_file", width=240)
        self.tree.column("confidence", width=110)
        self.tree.column("notes", width=260)

        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.on_double_click_edit)

        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")

        if not self.client:
            self.status_var.set("Set OPENAI_API_KEY to start processing.")

    def _set_busy(self, value: bool):
        self.busy = value
        state = "disabled" if value else "normal"
        self.upload_btn.configure(state=state)
        self.month_combo.configure(state="disabled" if value else "readonly")
        self.year_combo.configure(state=state)

    def _selected_month_year(self):
        try:
            selected_month = list(calendar.month_name).index(self.month_var.get())
            selected_year = int(self.year_var.get())
            if selected_month == 0:
                raise ValueError
            return selected_month, selected_year
        except Exception:
            raise ValueError("Please provide a valid month and year.")

    def upload_files(self):
        if self.busy:
            return

        if not self.client:
            self.client = self._build_client()
            if not self.client:
                messagebox.showwarning("Missing API Key", "OPENAI_API_KEY is not set.")
                return

        try:
            self._selected_month_year()
        except ValueError as exc:
            messagebox.showerror("Invalid Filter", str(exc))
            return

        paths = filedialog.askopenfilenames(
            title="Select check PDF/image files",
            filetypes=[("Supported", "*.pdf *.jpg *.jpeg *.png"), ("PDF", "*.pdf"), ("Images", "*.jpg *.jpeg *.png")],
        )
        if not paths:
            return

        bad = [p for p in paths if Path(p).suffix.lower() not in ALLOWED_EXTENSIONS]
        if bad:
            messagebox.showerror("Invalid Files", "Unsupported file type(s):\n" + "\n".join(bad))
            return

        self._set_busy(True)
        threading.Thread(target=self._process_files, args=(paths,), daemon=True).start()

    def _process_files(self, paths):
        self.root.after(0, lambda: self.progress.start(8))
        self.root.after(0, lambda: self.status_var.set("Processing files..."))

        selected_month, selected_year = self._selected_month_year()
        new_rows: List[CheckRow] = []

        for file_path in paths:
            try:
                pages = self._file_to_images(file_path)
                for page_num, image in enumerate(pages, start=1):
                    checks = self.extract_checks(image)
                    source_name = f"{Path(file_path).name} (page {page_num})" if len(pages) > 1 else Path(file_path).name

                    if not checks:
                        checks = [{"check_number": "", "vendor_name": "", "confidence": "low", "notes": "No check detected"}]

                    for check in checks:
                        new_rows.append(
                            CheckRow(
                                check_number=str(check.get("check_number", "")).strip(),
                                vendor_name=str(check.get("vendor_name", "")).strip(),
                                month=selected_month,
                                year=selected_year,
                                source_file=source_name,
                                confidence=str(check.get("confidence", "")).strip(),
                                notes=str(check.get("notes", "")).strip(),
                            )
                        )
            except Exception as exc:
                new_rows.append(
                    CheckRow(
                        check_number="",
                        vendor_name="",
                        month=selected_month,
                        year=selected_year,
                        source_file=Path(file_path).name,
                        confidence="error",
                        notes=f"Failed: {exc}",
                    )
                )

        self.rows.extend(new_rows)
        self.root.after(0, self._finish_processing, len(new_rows))

    def _finish_processing(self, added_count: int):
        self.progress.stop()
        self._set_busy(False)
        self.refresh_table()
        self.status_var.set(f"Done. Added {added_count} row(s).")

    def _file_to_images(self, file_path: str) -> List[Image.Image]:
        ext = Path(file_path).suffix.lower()
        if ext in {".jpg", ".jpeg", ".png"}:
            return [Image.open(file_path).convert("RGB")]

        if ext == ".pdf":
            doc = pdfium.PdfDocument(file_path)
            images = []
            for i in range(len(doc)):
                page = doc[i]
                bitmap = page.render(scale=2).to_pil()
                images.append(bitmap.convert("RGB"))
            return images

        raise ValueError("Unsupported file type")

    def extract_checks(self, image: Image.Image) -> List[Dict[str, str]]:
        if not self.client:
            raise RuntimeError("OpenAI client not initialized")

        image_bytes = io.BytesIO()
        image.save(image_bytes, format="PNG")
        encoded = base64.b64encode(image_bytes.getvalue()).decode("utf-8")

        prompt = (
            "Extract checks from this image. Return STRICT JSON only with this exact schema: "
            "{\"checks\":[{\"check_number\":\"\",\"vendor_name\":\"\",\"confidence\":\"high|medium|low\",\"notes\":\"\"}]}. "
            "If no check is present, return {\"checks\":[]}. "
            "Use the payee/vendor as vendor_name."
        )

        response = self.client.responses.create(
            model="gpt-4.1-mini",
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": f"data:image/png;base64,{encoded}"},
                    ],
                }
            ],
            text={"format": {"type": "json_object"}},
        )

        return self._parse_checks_json(response.output_text)

    @staticmethod
    def _parse_checks_json(raw: str) -> List[Dict[str, str]]:
        payload = raw.strip()
        if payload.startswith("```"):
            payload = payload.strip("`")
            payload = payload.replace("json", "", 1).strip()

        data = json.loads(payload)
        checks = data.get("checks", []) if isinstance(data, dict) else []
        normalized = []
        for item in checks:
            if not isinstance(item, dict):
                continue
            normalized.append(
                {
                    "check_number": str(item.get("check_number", "")),
                    "vendor_name": str(item.get("vendor_name", "")),
                    "confidence": str(item.get("confidence", "")),
                    "notes": str(item.get("notes", "")),
                }
            )
        return normalized

    def filtered_rows(self):
        try:
            selected_month, selected_year = self._selected_month_year()
        except ValueError as exc:
            messagebox.showerror("Invalid Filter", str(exc))
            return []

        return [r for r in self.rows if r.month == selected_month and r.year == selected_year]

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, row in enumerate(self.filtered_rows()):
            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    row.check_number,
                    row.vendor_name,
                    calendar.month_name[row.month],
                    row.year,
                    row.source_file,
                    row.confidence,
                    row.notes,
                ),
            )

    def on_double_click_edit(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return

        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace("#", "")) - 1
        if col_idx < 0:
            return

        x, y, width, height = self.tree.bbox(item_id, col)
        old_val = self.tree.item(item_id, "values")[col_idx]

        entry = ttk.Entry(self.tree)
        entry.insert(0, old_val)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()

        def save_edit(_):
            vals = list(self.tree.item(item_id, "values"))
            vals[col_idx] = entry.get()
            self.tree.item(item_id, values=vals)
            entry.destroy()
            self._sync_tree_to_rows()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda _: entry.destroy())

    def _sync_tree_to_rows(self):
        filtered = self.filtered_rows()
        for i, item_id in enumerate(self.tree.get_children()):
            vals = self.tree.item(item_id, "values")
            if i >= len(filtered):
                continue
            row = filtered[i]
            row.check_number = vals[0]
            row.vendor_name = vals[1]
            row.confidence = vals[5]
            row.notes = vals[6]

    def export_excel(self):
        rows = self.filtered_rows()
        if not rows:
            messagebox.showinfo("No Data", "No rows for selected month/year.")
            return

        month = rows[0].month
        year = rows[0].year
        save_path = filedialog.asksaveasfilename(
            title="Save Excel",
            defaultextension=".xlsx",
            initialfile=f"Checks_{year}_{month:02d}.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not save_path:
            return

        data = [asdict(r) for r in rows]
        for item in data:
            item["month"] = calendar.month_name[item["month"]]

        df = pd.DataFrame(data)
        df = df[["check_number", "vendor_name", "month", "year", "source_file", "confidence", "notes"]]
        df.columns = ["Check Number", "Vendor Name", "Month", "Year", "Source File", "Confidence", "Notes"]
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Exported", f"Saved: {save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    CheckExtractorApp(root)
    root.mainloop()
