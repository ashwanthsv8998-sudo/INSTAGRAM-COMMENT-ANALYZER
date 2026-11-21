import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

# Ensure Start-Menu IDLE finds openpyxl offline from your local venv
VENV_SITE_PACKAGES = r"D:\ASHWA\intern\venv\Lib\site-packages"
if (
    os.path.exists(VENV_SITE_PACKAGES)
    and VENV_SITE_PACKAGES not in sys.path
):
    sys.path.append(VENV_SITE_PACKAGES)

import openpyxl

EXCEL_PATH = r"D:\ASHWA\intern\comments_entry.xlsx"


def create_or_append_excel(username: str, comment: str) -> str:
    """
    Creates the Excel sheet with ['Username', 'Comment'] on the first entry,
    then connects and appends every subsequent entry to the bottom.
    """
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)

    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Comments"
        ws.append(["Username", "Comment"])
        status = "Created new Excel sheet and saved first entry."
    else:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        status = "Appended entry to existing Excel sheet."

    ws.append([username.strip(), comment.strip()])
    wb.save(EXCEL_PATH)
    return status


def read_all_rows() -> list[tuple[str, str]]:
    """Reads existing rows from the connected Excel sheet for live display."""
    if not os.path.exists(EXCEL_PATH):
        return []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) <= 1:
        return []
    return [
        (str(r[0] or ""), str(r[1] or ""))
        for r in rows[1:]
        if any(cell is not None for cell in r)
    ]


class ExcelEntryInterface(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Comment Data Entry -> Excel")
        self.geometry("620x480")
        self.configure(bg="#f8fafc")

        # --- Entry Form Section ---
        form_frame = ttk.LabelFrame(
            self, text="Enter Comment Details", padding=15
        )
        form_frame.pack(fill="x", padx=15, pady=(15, 8))

        ttk.Label(
            form_frame, text="Username:", font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, sticky="w", pady=6)
        self.username_entry = ttk.Entry(
            form_frame, width=45, font=("Segoe UI", 10)
        )
        self.username_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=6)

        ttk.Label(
            form_frame, text="Comment:", font=("Segoe UI", 10, "bold")
        ).grid(row=1, column=0, sticky="nw", pady=6)
        self.comment_text = tk.Text(
            form_frame, width=45, height=4, font=("Segoe UI", 10)
        )
        self.comment_text.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=6)

        form_frame.columnconfigure(1, weight=1)

        save_btn = ttk.Button(
            form_frame, text="Save to Excel Sheet", command=self.on_save
        )
        save_btn.grid(row=2, column=1, sticky="e", pady=(8, 0))

        # --- Connection Status Bar ---
        self.status_var = tk.StringVar()
        self.update_connection_status()
        status_label = ttk.Label(
            self, textvariable=self.status_var, foreground="#0f766e"
        )
        status_label.pack(anchor="w", padx=17, pady=(0, 6))

        # --- Live Two-Column Table Connected to Excel ---
        table_frame = ttk.LabelFrame(
            self, text="Connected Excel Sheet Preview (Username | Comment)", padding=10
        )
        table_frame.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        self.tree = ttk.Treeview(
            table_frame, columns=("Username", "Comment"), show="headings"
        )
        self.tree.heading("Username", text="Username")
        self.tree.heading("Comment", text="Comment")
        self.tree.column("Username", width=160, anchor="w")
        self.tree.column("Comment", width=400, anchor="w")
        self.tree.pack(fill="both", expand=True)

        self.refresh_table()

    def update_connection_status(self):
        if os.path.exists(EXCEL_PATH):
            self.status_var.set(f"Connected to Excel: {EXCEL_PATH}")
        else:
            self.status_var.set(
                f"Ready: First entry will automatically create {EXCEL_PATH}"
            )

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for user, comment in read_all_rows():
            self.tree.insert("", "end", values=(user, comment))

    def on_save(self):
        username = self.username_entry.get().strip()
        comment = self.comment_text.get("1.0", "end").strip()

        if not username or not comment:
            messagebox.showwarning(
                "Required Fields", "Please enter both Username and Comment."
            )
            return

        try:
            msg = create_or_append_excel(username, comment)
            self.username_entry.delete(0, "end")
            self.comment_text.delete("1.0", "end")
            self.username_entry.focus()
            self.update_connection_status()
            self.refresh_table()
            messagebox.showinfo("Saved", msg)
        except PermissionError:
            messagebox.showerror(
                "Excel File Open",
                "Please close 'comments_entry.xlsx' in Microsoft Excel before saving.",
            )


if __name__ == "__main__":
    app = ExcelEntryInterface()
    app.mainloop()
