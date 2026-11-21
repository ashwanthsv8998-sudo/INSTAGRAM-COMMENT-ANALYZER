import csv
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(SCRIPT_DIR, "comments_entry.xlsx")
DEFAULT_CSV = os.path.join(SCRIPT_DIR, "comments_entry.csv")

# Ensure Start-Menu IDLE finds openpyxl offline from your local venv
VENV_SITE_PACKAGES = os.path.join(SCRIPT_DIR, "venv", "Lib", "site-packages")
if os.path.exists(VENV_SITE_PACKAGES) and VENV_SITE_PACKAGES not in sys.path:
    sys.path.append(VENV_SITE_PACKAGES)

try:
    import openpyxl

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def detect_default_file() -> str:
    if HAS_OPENPYXL and os.path.exists(DEFAULT_XLSX):
        return DEFAULT_XLSX
    if os.path.exists(DEFAULT_CSV):
        return DEFAULT_CSV
    return DEFAULT_XLSX if HAS_OPENPYXL else DEFAULT_CSV


def load_records(file_path: str) -> tuple[list[str], list[dict]]:
    """Loads headers and all rows as dictionaries from .xlsx or .csv."""
    if not os.path.exists(file_path):
        return [], []

    ext = os.path.splitext(file_path)[1].lower()
    raw_rows = []

    if ext == ".xlsx":
        if not HAS_OPENPYXL:
            raise RuntimeError(
                "openpyxl is required to read .xlsx files. Select a .csv file"
                " or ensure venv is present."
            )
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        raw_rows = list(ws.iter_rows(values_only=True))
    elif ext == ".csv":
        with open(file_path, mode="r", encoding="utf-8-sig") as f:
            raw_rows = list(csv.reader(f))

    if len(raw_rows) < 2:
        return [], []

    headers = [
        str(h).strip() if h is not None and str(h).strip() else f"Column_{i+1}"
        for i, h in enumerate(raw_rows[0])
    ]

    records = []
    for row in raw_rows[1:]:
        if not any(cell is not None and str(cell).strip() for cell in row):
            continue
        record = {}
        for i, h in enumerate(headers):
            val = row[i] if i < len(row) and row[i] is not None else ""
            record[h] = str(val).strip()
        records.append(record)

    return headers, records


class CommentSearchInterface(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Comment Keyword Search & User Details")
        self.geometry("760x560")
        self.configure(bg="#f8fafc")

        self.current_file = detect_default_file()
        self.headers: list[str] = []
        self.all_records: list[dict] = []
        self.filtered_records: list[dict] = []

        self._build_ui()
        self.reload_data()

    def _build_ui(self):
        # --- Top File Connection Bar ---
        file_bar = ttk.Frame(self)
        file_bar.pack(fill="x", padx=15, pady=(12, 4))

        self.file_label_var = tk.StringVar()
        ttk.Label(
            file_bar,
            textvariable=self.file_label_var,
            foreground="#0f766e",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")

        ttk.Button(
            file_bar, text="Refresh Data", command=self.reload_data
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            file_bar, text="Browse File...", command=self.browse_file
        ).pack(side="right")

        # --- Search Bar Section ---
        search_frame = ttk.LabelFrame(
            self, text="Keyword Search & Analysis", padding=12
        )
        search_frame.pack(fill="x", padx=15, pady=6)

        ttk.Label(
            search_frame, text="Enter Keyword:", font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=(0, 8))

        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            search_frame, textvariable=self.search_var, font=("Segoe UI", 11)
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.search_entry.bind("<Return>", lambda e: self.perform_search())
        self.search_entry.focus()

        ttk.Button(
            search_frame, text="Search", command=self.perform_search
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            search_frame, text="Show All", command=self.show_all
        ).pack(side="left")

        # --- Keyword Analysis Summary Label ---
        self.stats_var = tk.StringVar(
            value="Type a keyword and press Enter to analyze matching comments."
        )
        ttk.Label(
            self,
            textvariable=self.stats_var,
            foreground="#1e293b",
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w", padx=17, pady=(2, 6))

        # --- Results Table ---
        table_frame = ttk.LabelFrame(
            self, text="Matching Comments & User Details", padding=8
        )
        table_frame.pack(fill="both", expand=True, padx=15, pady=(0, 8))

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical")
        self.tree = ttk.Treeview(
            table_frame,
            show="headings",
            height=9,
            yscrollcommand=y_scroll.set,
        )
        y_scroll.config(command=self.tree.yview)
        y_scroll.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        # --- Selected User & Highlighted Comment Detail Pane ---
        detail_frame = ttk.LabelFrame(
            self,
            text="Selected User Details & Highlighted Comment",
            padding=10,
        )
        detail_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.detail_text = tk.Text(
            detail_frame,
            height=6,
            font=("Segoe UI", 10),
            wrap="word",
            bg="#ffffff",
            relief="flat",
        )
        self.detail_text.pack(fill="x")
        self.detail_text.tag_configure(
            "bold_label", font=("Segoe UI", 10, "bold"), foreground="#0f172a"
        )
        self.detail_text.tag_configure(
            "highlight",
            background="#fef08a",
            foreground="#854d0e",
            font=("Segoe UI", 10, "bold"),
        )
        self.detail_text.config(state="disabled")

    def browse_file(self):
        chosen = filedialog.askopenfilename(
            initialdir=SCRIPT_DIR,
            title="Select Excel or CSV File",
            filetypes=[
                ("Excel & CSV Files", "*.xlsx *.csv"),
                ("Excel Workbook", "*.xlsx"),
                ("CSV File", "*.csv"),
            ],
        )
        if chosen:
            self.current_file = chosen
            self.reload_data()

    def reload_data(self):
        try:
            self.headers, self.all_records = load_records(self.current_file)
        except Exception as exc:
            messagebox.showerror("Load Error", str(exc))
            return

        if os.path.exists(self.current_file):
            self.file_label_var.set(
                f"Connected: {self.current_file} ({len(self.all_records)} total"
                " rows)"
            )
        else:
            self.file_label_var.set(
                f"Waiting for file: {self.current_file} (Add entries first, then"
                " click Refresh)"
            )

        self._configure_columns()
        self.perform_search()

    def _configure_columns(self):
        cols = self.headers if self.headers else ["Username", "Comment"]
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            width = (
                360
                if ("comment" in col.lower() or "text" in col.lower())
                else 150
            )
            self.tree.column(col, width=width, anchor="w")

    def _find_comment_column(self) -> str | None:
        if not self.headers:
            return None
        for h in self.headers:
            if "comment" in h.lower() or "text" in h.lower():
                return h
        return self.headers[1] if len(self.headers) > 1 else self.headers[0]

    def show_all(self):
        self.search_var.set("")
        self.perform_search()

    def perform_search(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        keyword = self.search_var.get().strip()
        comment_col = self._find_comment_column()

        if not self.all_records:
            self.filtered_records = []
            self.stats_var.set("No records loaded yet.")
            self._clear_details()
            return

        if not keyword:
            self.filtered_records = list(self.all_records)
            self.stats_var.set(
                f"Showing all {len(self.filtered_records)} record(s). Enter a"
                " keyword above to filter."
            )
        else:
            kw_lower = keyword.lower()
            matched = []
            total_mentions = 0
            for rec in self.all_records:
                text_to_search = (
                    rec.get(comment_col, "")
                    if comment_col
                    else " ".join(rec.values())
                )
                count = text_to_search.lower().count(kw_lower)
                if count > 0:
                    matched.append(rec)
                    total_mentions += count

            self.filtered_records = matched
            self.stats_var.set(
                f"Keyword '{keyword}': found in {len(matched)} comment(s)"
                f" ({total_mentions} total occurrence(s))."
            )

        cols = self.tree["columns"]
        for idx, rec in enumerate(self.filtered_records):
            values = [rec.get(c, "") for c in cols]
            self.tree.insert("", "end", iid=str(idx), values=values)

        if self.filtered_records:
            self.tree.selection_set("0")
            self.on_row_select(None)
        else:
            self._clear_details()

    def _clear_details(self):
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.config(state="disabled")

    def on_row_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        record = self.filtered_records[idx]
        keyword = self.search_var.get().strip()
        comment_col = self._find_comment_column()

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")

        for key, val in record.items():
            self.detail_text.insert("end", f"{key}: ", "bold_label")
            if key == comment_col and keyword:
                pattern = re.compile(re.escape(keyword), re.IGNORECASE)
                last_end = 0
                for match in pattern.finditer(val):
                    self.detail_text.insert("end", val[last_end : match.start()])
                    self.detail_text.insert("end", match.group(0), "highlight")
                    last_end = match.end()
                self.detail_text.insert("end", val[last_end:] + "\n")
            else:
                self.detail_text.insert("end", f"{val}\n")

        self.detail_text.config(state="disabled")


if __name__ == "__main__":
    app = CommentSearchInterface()
    app.mainloop()
