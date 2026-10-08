import csv
import json
import logging
import os
import queue
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import requests
except ImportError:
    requests = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

APP_TITLE = "AKASH TECH SOLUTIONS - Project API Automation & Data Integration"
APP_VERSION = "1.2.2"
BASE_DIR = Path(__file__).resolve().parents[1]
LOG_DIR = BASE_DIR / "logs"
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_DIR = BASE_DIR / "config"
LOG_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
CONFIG_DIR.mkdir(exist_ok=True)

GREEN = "#08783F"
GREEN_DARK = "#075A30"
GREEN_LIGHT = "#EAF7EF"
GREEN_PALE = "#F4FBF6"
BLACK = "#0B1711"
TEXT = "#17221C"
MUTED = "#66736B"
BORDER = "#D5E2D9"
WHITE = "#FFFFFF"
RED = "#C62828"
AMBER = "#9A6700"

logging.basicConfig(
    filename=LOG_DIR / "api_automation.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


class DemoAPI:
    """Small deterministic offline API for demonstrations and testing."""

    def __init__(self):
        self.records = [
            {"id": 101, "name": "Acme Motors", "email": "contact@acmemotors.example", "status": "Active", "country": "India"},
            {"id": 102, "name": "Northwind Digital", "email": "hello@northwind.example", "status": "Active", "country": "UK"},
            {"id": 103, "name": "BlueSky Retail", "email": "sales@bluesky.example", "status": "Pending", "country": "UAE"},
            {"id": 104, "name": "GreenLeaf Media", "email": "info@greenleaf.example", "status": "Inactive", "country": "Australia"},
            {"id": 105, "name": "Vertex Labs", "email": "team@vertex.example", "status": "Active", "country": "USA"},
            {"id": 106, "name": "PixelWorks", "email": "hello@pixelworks.example", "status": "Pending", "country": "Canada"},
        ]
        self.next_id = 107

    def request(self, method, payload=None, record_id=None):
        method = method.upper()
        if method == "GET":
            if record_id:
                for r in self.records:
                    if str(r["id"]) == str(record_id):
                        return {"status": 200, "data": r}
                return {"status": 404, "data": {"error": "Record not found"}}
            return {"status": 200, "data": list(self.records)}
        if method == "POST":
            data = dict(payload or {})
            data["id"] = self.next_id
            self.next_id += 1
            self.records.append(data)
            return {"status": 201, "data": data}
        if method == "PUT":
            for i, r in enumerate(self.records):
                if str(r["id"]) == str(record_id):
                    updated = dict(r)
                    updated.update(payload or {})
                    updated["id"] = r["id"]
                    self.records[i] = updated
                    return {"status": 200, "data": updated}
            return {"status": 404, "data": {"error": "Record not found"}}
        if method == "DELETE":
            before = len(self.records)
            self.records = [r for r in self.records if str(r["id"]) != str(record_id)]
            if len(self.records) < before:
                return {"status": 200, "data": {"deleted": int(record_id) if str(record_id).isdigit() else record_id}}
            return {"status": 404, "data": {"error": "Record not found"}}
        return {"status": 400, "data": {"error": "Unsupported method"}}


class CompletionDialog(tk.Toplevel):
    def __init__(self, parent, status, rows, elapsed, output_path=None):
        super().__init__(parent)
        ok = int(status or 500) < 400
        self.title("Request Complete" if ok else "Request Failed")
        self.geometry("520x340")
        self.resizable(False, False)
        self.configure(bg=WHITE)
        self.transient(parent)
        self.grab_set()

        header_color = GREEN_DARK if ok else RED
        header = tk.Frame(self, bg=header_color, height=66)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="✓" if ok else "!", bg=header_color, fg=WHITE, font=("Segoe UI", 22, "bold")).pack(side="left", padx=22)
        tk.Label(header, text="Request Complete" if ok else "Request Failed", bg=header_color, fg=WHITE, font=("Segoe UI", 15, "bold")).pack(side="left")

        body = tk.Frame(self, bg=WHITE)
        body.pack(fill="both", expand=True, padx=28, pady=20)
        message = "The API request completed successfully." if ok else "The API request returned an error status."
        tk.Label(body, text=message, bg=WHITE, fg=TEXT, font=("Segoe UI", 10)).pack(anchor="w")

        stats = tk.Frame(body, bg=GREEN_PALE, highlightbackground=BORDER, highlightthickness=1)
        stats.pack(fill="x", pady=16)
        for label, value in [("HTTP", status), ("Rows", rows), ("Time", f"{elapsed:.2f}s")]:
            cell = tk.Frame(stats, bg=GREEN_PALE)
            cell.pack(side="left", expand=True, fill="x", padx=8, pady=12)
            tk.Label(cell, text=str(value), bg=GREEN_PALE, fg=GREEN_DARK, font=("Segoe UI", 14, "bold")).pack()
            tk.Label(cell, text=label, bg=GREEN_PALE, fg=MUTED, font=("Segoe UI", 8)).pack()

        if output_path:
            tk.Label(body, text="Output location", bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(body, text=str(output_path), bg=WHITE, fg=TEXT, font=("Segoe UI", 9), wraplength=450, justify="left").pack(anchor="w", pady=(2, 10))

        tk.Button(body, text="Close", command=self.destroy, bg=GREEN, fg=WHITE,
                  activebackground=GREEN_DARK, activeforeground=WHITE, relief="flat", bd=0,
                  padx=22, pady=9, font=("Segoe UI", 9, "bold")).pack(side="right")


class APIAutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        try:
            self.state("zoomed")
        except tk.TclError:
            self.geometry("1366x900")
        self.minsize(1180, 760)
        self.configure(bg=GREEN_PALE)
        icon_path = BASE_DIR / "assets" / "akash_api.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        self.demo_api = DemoAPI()
        self.last_json = None
        self.last_rows = []
        self.last_result = None
        self.running = False
        self.request_started = None
        self.request_start_label = "--"
        self.request_end_label = "--"
        self.msg_queue = queue.Queue()
        self.log_lines = []
        self.current_output = None

        self._setup_style()
        self._build_ui()
        self.after(100, self._poll_queue)
        self._status("Project 03 | API Automation & Data Integration - Ready - Start: -- - End: --")

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 9), background=WHITE, fieldbackground=WHITE, foreground=TEXT, borderwidth=0)
        style.configure("Treeview.Heading", background=GREEN_DARK, foreground=WHITE, font=("Segoe UI", 9, "bold"), padding=8)
        style.map("Treeview", background=[("selected", "#D9F1E2")], foreground=[("selected", TEXT)])
        style.configure("TCombobox", padding=7, font=("Segoe UI", 9))
        style.configure("TEntry", padding=7)
        style.configure("TNotebook", background=GREEN_PALE, borderwidth=0)
        style.configure("TNotebook.Tab", background="#DCE9E0", foreground=TEXT, padding=(14, 8), font=("Segoe UI", 9, "bold"))
        style.map("TNotebook.Tab", background=[("selected", GREEN)], foreground=[("selected", WHITE)])
        style.configure("Horizontal.TProgressbar", troughcolor="#DCE9E0", background=GREEN, bordercolor="#DCE9E0", lightcolor=GREEN, darkcolor=GREEN)

    def _build_ui(self):
        self._build_header()
        main = tk.Frame(self, bg=GREEN_PALE)
        main.pack(fill="both", expand=True, padx=18, pady=14)
        self._build_toolbar(main)
        self._build_request_panel(main)
        self._build_result_panel(main)
        self._build_progress_area(main)
        self.status = tk.Label(self, bg=GREEN_DARK, fg=WHITE, anchor="w", padx=18, font=("Segoe UI", 8, "bold"), height=2)
        self.status.pack(fill="x", side="bottom")

    def _build_header(self):
        header = tk.Frame(self, bg=BLACK, height=96)
        header.pack(fill="x")
        header.pack_propagate(False)

        logo = tk.Canvas(header, width=66, height=66, bg=BLACK, highlightthickness=0)
        logo.pack(side="left", padx=(22, 12), pady=15)
        logo.create_polygon(33, 5, 59, 57, 45, 57, 38, 43, 28, 43, 21, 57, 7, 57, fill="#00B85A", outline="")
        logo.create_polygon(31, 18, 22, 40, 40, 40, fill="#8EF0B8", outline="")

        title = tk.Frame(header, bg=BLACK)
        title.pack(side="left", fill="y", pady=13)
        tk.Label(title, text="AKASH", bg=BLACK, fg=WHITE, font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(title, text="TECH SOLUTIONS", bg=BLACK, fg="#9CE8BB", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Frame(header, bg="#435249", width=1).pack(side="left", fill="y", pady=18, padx=22)

        project = tk.Frame(header, bg=BLACK)
        project.pack(side="left", pady=18)
        tk.Label(project, text="Project - API Automation & Data Integration", bg=BLACK, fg=WHITE, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(project, text="API REQUESTS   |   AUTH   |   JSON   |   PAGINATION   |   TRANSFORM   |   EXPORT   |   LOGGING", bg=BLACK, fg="#B6C6BC", font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(4, 0))

        actions = tk.Frame(header, bg=BLACK)
        actions.pack(side="right", padx=22)
        self._header_button(actions, "⚙", "Settings", self.show_settings).pack(side="left", padx=4)
        self._header_button(actions, "?", "Help", self.show_help).pack(side="left", padx=4)
        self._header_button(actions, "i", "About", self.show_about).pack(side="left", padx=4)

    def _header_button(self, parent, symbol, text, command):
        f = tk.Frame(parent, bg="#102019", highlightbackground="#284337", highlightthickness=1)
        tk.Button(f, text=symbol, command=command, bg="#102019", fg=WHITE, activebackground=GREEN, activeforeground=WHITE, relief="flat", bd=0, font=("Segoe UI", 13, "bold"), width=4).pack()
        tk.Label(f, text=text, bg="#102019", fg="#D8E7DE", font=("Segoe UI", 7, "bold")).pack(padx=5, pady=(0, 5))
        return f

    def _button(self, parent, text, command, primary=False, width=None):
        return tk.Button(parent, text=text, command=command, bg=GREEN if primary else WHITE, fg=WHITE if primary else TEXT,
                          activebackground=GREEN_DARK if primary else "#E8F2EB", activeforeground=WHITE if primary else TEXT,
                          relief="flat", bd=0, padx=15, pady=7, font=("Segoe UI", 9, "bold"), width=width or 0, cursor="hand2")

    def _build_toolbar(self, parent):
        bar = tk.Frame(parent, bg=GREEN_PALE)
        bar.pack(fill="x")
        self._button(bar, "▶  Send Request", self.send_request, True).pack(side="left", padx=(0, 7))
        self._button(bar, "▣  Load Config", self.load_config).pack(side="left", padx=7)
        self._button(bar, "＋  New Request", self.clear_request).pack(side="left", padx=7)
        self._button(bar, "▤  Open Output", self.open_output).pack(side="left", padx=7)
        self._button(bar, "✕  Clear Log", self.clear_log).pack(side="left", padx=7)
        self.mode_label = tk.Label(bar, text="● DEMO MODE", bg=GREEN_PALE, fg=GREEN_DARK, font=("Segoe UI", 10, "bold"))
        self.mode_label.pack(side="right", padx=10)

    def _section_title(self, parent, title, subtitle=None):
        head = tk.Frame(parent, bg=GREEN_LIGHT, height=38)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(head, text=title, bg=GREEN_LIGHT, fg=GREEN_DARK, font=("Segoe UI", 10, "bold")).pack(side="left", padx=15, pady=13)
        if subtitle:
            tk.Label(head, text=subtitle, bg=GREEN_LIGHT, fg=MUTED, font=("Segoe UI", 8)).pack(side="right", padx=15)
        return head

    def _build_request_panel(self, parent):
        card = tk.Frame(parent, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", pady=(8, 6))
        self._section_title(card, "REQUEST BUILDER", "REST API • JSON • AUTHENTICATION")
        body = tk.Frame(card, bg=WHITE)
        body.pack(fill="x", padx=12, pady=5)

        self.method_var = tk.StringVar(value="GET")
        self.url_var = tk.StringVar(value="https://jsonplaceholder.typicode.com/users")
        self.auth_type_var = tk.StringVar(value="None")
        self.token_var = tk.StringVar()
        self.record_id_var = tk.StringVar()
        self.page_var = tk.StringVar(value="1")
        self.per_page_var = tk.StringVar(value="10")
        self.retries_var = tk.StringVar(value="3")
        self.timeout_var = tk.StringVar(value="20")
        self.demo_var = tk.BooleanVar(value=True)

        self.method_combo = self._label_combo(body, "METHOD", self.method_var, ["GET", "POST", "PUT", "DELETE"], 0, 0, 12)
        self.method_combo.bind("<<ComboboxSelected>>", lambda _e: self._method_changed())
        self._label_entry(body, "API ENDPOINT", self.url_var, 0, 1, 42)
        self._label_combo(body, "AUTH", self.auth_type_var, ["None", "Bearer Token", "API Key"], 0, 2, 16)
        self._label_entry(body, "TOKEN / API KEY", self.token_var, 0, 3, 25, show="*")
        self._label_entry(body, "RECORD ID", self.record_id_var, 0, 4, 14)

        self._label_entry(body, "PAGE", self.page_var, 1, 0, 10)
        self._label_entry(body, "PER PAGE", self.per_page_var, 1, 1, 10)
        self._label_entry(body, "RETRIES", self.retries_var, 1, 2, 10)
        self._label_entry(body, "TIMEOUT (SEC)", self.timeout_var, 1, 3, 12)
        demo_box = tk.Frame(body, bg=WHITE)
        demo_box.grid(row=1, column=4, sticky="ew", padx=7, pady=(5, 2))
        tk.Checkbutton(demo_box, text="DEMO MODE", variable=self.demo_var, command=self._mode_changed, bg=WHITE, fg=GREEN_DARK,
                       selectcolor=GREEN_LIGHT, activebackground=WHITE, activeforeground=GREEN_DARK, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 1))

        tk.Label(body, text="JSON BODY (POST / PUT)", bg=WHITE, fg=TEXT, font=("Segoe UI", 8, "bold")).grid(row=2, column=0, columnspan=5, sticky="w", padx=7, pady=(6, 3))
        self.body_text = tk.Text(body, height=5, wrap="none", bg="#F8FBF9", fg=TEXT, insertbackground=TEXT, relief="solid", bd=1, highlightthickness=0, font=("Consolas", 9))
        self.body_text.grid(row=3, column=0, columnspan=5, sticky="nsew", padx=7, pady=(0, 8))
        self.body_text.insert("1.0", '{"name":"New Client","email":"newclient@example.com","status":"Active","country":"India"}')
        for c in range(5):
            body.grid_columnconfigure(c, weight=1)
        body.grid_rowconfigure(3, weight=1)
        self._method_changed()

    def _method_changed(self):
        method = self.method_var.get().upper()
        body_methods = method in {"POST", "PUT"}
        if body_methods:
            self.body_text.configure(state="normal", bg="#F8FBF9", fg=TEXT)
            if not self.body_text.get("1.0", "end").strip():
                self.body_text.insert("1.0", '{"name":"New Client","email":"newclient@example.com","status":"Active","country":"India"}')
        else:
            self.body_text.configure(state="normal")
            self.body_text.delete("1.0", "end")
            self.body_text.configure(state="disabled", bg="#EEF3EF", fg=MUTED)

    def _label_entry(self, parent, label, variable, row, col, width=20, show=None):
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=col, sticky="ew", padx=7, pady=1)
        tk.Label(box, text=label, bg=WHITE, fg=TEXT, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        entry = tk.Entry(box, textvariable=variable, width=width, show=show, bg="#F8FBF9", fg=TEXT, insertbackground=TEXT, relief="solid", bd=1)
        entry.pack(fill="x", pady=(2, 0), ipady=3)
        return entry

    def _label_combo(self, parent, label, variable, values, row, col, width):
        box = tk.Frame(parent, bg=WHITE)
        box.grid(row=row, column=col, sticky="ew", padx=7, pady=1)
        tk.Label(box, text=label, bg=WHITE, fg=TEXT, font=("Segoe UI", 8, "bold")).pack(anchor="w")
        combo = ttk.Combobox(box, textvariable=variable, values=values, state="readonly", width=width)
        combo.pack(fill="x", pady=(3, 0))
        return combo

    def _build_result_panel(self, parent):
        card = tk.Frame(parent, bg=WHITE, highlightbackground=BORDER, highlightthickness=1, height=360)
        card.pack(fill="both", expand=True, pady=(0, 0))
        card.pack_propagate(False)
        self._section_title(card, "RESULTS & WORKSPACE", "RESPONSE • TRANSFORM • LOG • CONFIG")
        notebook = ttk.Notebook(card)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.notebook = notebook

        response_tab = tk.Frame(notebook, bg=WHITE)
        notebook.add(response_tab, text=" RESPONSE ")
        self.response_text = tk.Text(response_tab, bg="#0B110D", fg="#D9EEE0", insertbackground=WHITE, relief="flat", font=("Consolas", 9))
        self.response_text.pack(fill="both", expand=True, padx=8, pady=8)

        transform_tab = tk.Frame(notebook, bg=WHITE)
        notebook.add(transform_tab, text=" TRANSFORM & EXPORT ")
        top = tk.Frame(transform_tab, bg=WHITE)
        top.pack(fill="x", padx=10, pady=(4, 3))
        tk.Label(top, text="JSON ARRAY / OBJECT → TABULAR DATA", bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(top, text="Optional: enter the array field name when the API wraps records inside data/results/items/users.", bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 2))
        controls = tk.Frame(top, bg=WHITE)
        controls.pack(fill="x")
        self.flatten_var = tk.StringVar()
        tk.Entry(controls, textvariable=self.flatten_var, bg="#F8FBF9", fg=TEXT, relief="solid", bd=1, width=28).pack(side="left", ipady=3)
        self._button(controls, "PREVIEW TABLE", self.preview_table, True).pack(side="left", padx=8)
        self._button(controls, "EXPORT CSV", self.export_csv).pack(side="left", padx=4)
        self._button(controls, "EXPORT EXCEL", self.export_excel).pack(side="left", padx=4)
        table_frame = tk.Frame(transform_tab, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_frame, show="headings", selectmode="extended", height=8)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y.grid(row=0, column=1, sticky="ns")
        x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)

        log_tab = tk.Frame(notebook, bg=WHITE)
        notebook.add(log_tab, text=" REQUEST LOG ")
        log_tools = tk.Frame(log_tab, bg=WHITE)
        log_tools.pack(fill="x", padx=10, pady=8)
        self._button(log_tools, "CLEAR LOG", self.clear_log).pack(side="left")
        self.log_text = tk.Text(log_tab, bg="#0B110D", fg="#BCD4C2", insertbackground=WHITE, relief="flat", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        config_tab = tk.Frame(notebook, bg=WHITE)
        notebook.add(config_tab, text=" CONFIGURATION ")
        tk.Label(config_tab, text="Save reusable API endpoint settings for repeatable business workflows.", bg=WHITE, fg=TEXT, font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(16, 8))
        cfg = tk.Frame(config_tab, bg=WHITE)
        cfg.pack(anchor="w", padx=14)
        self._button(cfg, "SAVE CONFIG", self.save_config, True).pack(side="left", padx=(0, 8))
        self._button(cfg, "LOAD CONFIG", self.load_config).pack(side="left")
        tk.Label(config_tab, text="Security note: token/API-key values are saved only if you explicitly choose to save this configuration.", bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=12)

    def _build_progress_area(self, parent):
        area = tk.Frame(parent, bg=GREEN_PALE, height=48)
        area.pack(fill="x", pady=(10, 0))
        area.pack_propagate(False)
        left = tk.Frame(area, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        tk.Label(left, text="REQUEST PROGRESS", bg=WHITE, fg=GREEN_DARK, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(8, 2))
        self.progress = ttk.Progressbar(left, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=12, pady=(0, 10))
        self.progress["value"] = 0

        right = tk.Frame(area, bg=WHITE, highlightbackground=BORDER, highlightthickness=1, width=310)
        right.pack(side="right", fill="y", padx=(6, 0))
        right.pack_propagate(False)
        self.summary_var = tk.StringVar(value="Requests: 0   •   Success: 0   •   Errors: 0")
        tk.Label(right, textvariable=self.summary_var, bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(expand=True)
        self.total_requests = 0
        self.success_requests = 0
        self.error_requests = 0

    def _mode_changed(self):
        if self.demo_var.get():
            self.mode_label.config(text="● DEMO MODE", fg=GREEN_DARK)
            self._status("Demo Mode enabled | Requests run locally without credentials")
        else:
            self.mode_label.config(text="● LIVE API", fg=AMBER)
            self._status("Live API mode | Verify endpoint, authentication and payload before sending")

    def _load_sample_request(self):
        self._log("Application started")
        self._log("Demo API ready: GET, POST, PUT and DELETE")
        self._mode_changed()

    def _status(self, text):
        if hasattr(self, "status"):
            self.status.config(text=text)

    def _log(self, message, level="INFO"):
        line = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {level:<5} | {message}"
        self.log_lines.append(line)
        if hasattr(self, "log_text"):
            self.log_text.insert("end", line + "\n")
            self.log_text.see("end")
        getattr(logging, level.lower(), logging.info)(message)

    def clear_log(self):
        self.log_lines.clear()
        if hasattr(self, "log_text"):
            self.log_text.delete("1.0", "end")
        self._log("Log cleared")

    def clear_request(self):
        self.method_var.set("GET")
        self.url_var.set("https://jsonplaceholder.typicode.com/users")
        self.auth_type_var.set("None")
        self.token_var.set("")
        self.record_id_var.set("")
        self.page_var.set("1")
        self.per_page_var.set("10")
        self.retries_var.set("3")
        self.timeout_var.set("20")
        self.demo_var.set(True)
        self.body_text.configure(state="normal")
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", '{"name":"New Client","email":"newclient@example.com","status":"Active","country":"India"}')
        self._method_changed()
        self.response_text.delete("1.0", "end")
        self.last_json = None
        self.last_rows = []
        self.last_result = None
        self.progress["value"] = 0
        self._clear_tree()
        self._mode_changed()
        self._status("Ready | New request")
        self._log("Request form reset")

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()

    def _read_request_values(self):
        method = self.method_var.get().upper()
        url = self.url_var.get().strip()
        auth = self.auth_type_var.get()
        token = self.token_var.get().strip()
        record_id = self.record_id_var.get().strip()
        page = self.page_var.get().strip()
        per_page = self.per_page_var.get().strip()
        retries = max(1, int(self.retries_var.get()))
        timeout = max(1, int(self.timeout_var.get()))
        demo = bool(self.demo_var.get())
        raw_body = self.body_text.get("1.0", "end").strip() if method in {"POST", "PUT"} else ""
        payload = None
        if raw_body:
            try:
                payload = json.loads(raw_body)
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON body: {exc}")
        return {"method": method, "url": url, "auth": auth, "token": token, "record_id": record_id,
                "page": page, "per_page": per_page, "retries": retries, "timeout": timeout, "demo": demo, "payload": payload}

    def send_request(self):
        if self.running:
            return
        try:
            values = self._read_request_values()
        except Exception as exc:
            messagebox.showerror("Request Validation", str(exc))
            return
        if not values["demo"] and not values["url"]:
            messagebox.showwarning("Request Validation", "Please enter an API endpoint.")
            return
        self.running = True
        self.request_started = time.perf_counter()
        self.progress["value"] = 10
        self.total_requests += 1
        self._update_summary()
        self.request_start_label = datetime.now().strftime("%H:%M:%S")
        self.request_end_label = "--"
        self._status(f"Project 03 | {values['method']} Request - Running - Start: {self.request_start_label} - End: --")
        self._log(f"Request started: {values['method']} {values['url']} | {'Demo' if values['demo'] else 'Live'}")
        threading.Thread(target=self._request_worker, args=(values,), daemon=True).start()

    def _headers(self, auth, token):
        headers = {"Accept": "application/json", "Content-Type": "application/json"}
        if token and auth == "Bearer Token":
            headers["Authorization"] = f"Bearer {token}"
        elif token and auth == "API Key":
            headers["X-API-Key"] = token
        return headers

    def _request_worker(self, values):
        try:
            method = values["method"]
            record_id = values["record_id"]
            if values["demo"]:
                time.sleep(0.35)
                result = self.demo_api.request(method, values["payload"], record_id or None)
                result["demo"] = True
            else:
                if requests is None:
                    raise RuntimeError("The requests package is not installed. Run: pip install -r requirements.txt")
                url = values["url"]
                if record_id and method in {"GET", "PUT", "DELETE"} and not url.rstrip("/").endswith(record_id):
                    url = url.rstrip("/") + "/" + record_id
                params = {"page": values["page"], "limit": values["per_page"]} if method == "GET" and values["page"] else None
                result = None
                last_error = None
                for attempt in range(1, values["retries"] + 1):
                    self.msg_queue.put(("progress", min(90, 20 + int(attempt / values["retries"] * 60))))
                    try:
                        resp = requests.request(method, url, headers=self._headers(values["auth"], values["token"]),
                                                json=values["payload"] if method in {"POST", "PUT"} else None,
                                                params=params, timeout=values["timeout"])
                        try:
                            data = resp.json()
                        except ValueError:
                            data = {"raw_response": resp.text}
                        result = {"status": resp.status_code, "data": data, "demo": False, "attempts": attempt}
                        if resp.status_code < 500:
                            break
                        last_error = RuntimeError(f"Server returned HTTP {resp.status_code}")
                    except Exception as exc:
                        last_error = exc
                    if attempt < values["retries"]:
                        time.sleep(1.0 * attempt)
                if result is None:
                    raise last_error or RuntimeError("Request failed")
            self.msg_queue.put(("complete", result))
        except Exception as exc:
            self.msg_queue.put(("error", str(exc)))

    def _poll_queue(self):
        try:
            while True:
                event, payload = self.msg_queue.get_nowait()
                if event == "progress":
                    self.progress["value"] = payload
                elif event == "complete":
                    self._request_complete(payload)
                elif event == "error":
                    self._request_failed(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _request_complete(self, result):
        self.running = False
        self.progress["value"] = 100
        self.last_result = result
        self.last_json = result.get("data")
        self.success_requests += 1 if int(result.get("status", 500)) < 400 else 0
        if int(result.get("status", 500)) >= 400:
            self.error_requests += 1
        self._update_summary()
        self.response_text.delete("1.0", "end")
        self.response_text.insert("1.0", json.dumps(result, indent=2, ensure_ascii=False))
        elapsed = time.perf_counter() - self.request_started if self.request_started else 0
        self.request_end_label = datetime.now().strftime("%H:%M:%S")
        self._status(f"Project 03 | {self.method_var.get()} Request - Completed HTTP {result.get('status')} ({elapsed:.2f}s) - Start: {self.request_start_label} - End: {self.request_end_label}")
        self._log(f"Request completed: HTTP {result.get('status')} in {elapsed:.2f}s")
        self.preview_table(silent=True)
        CompletionDialog(self, result.get("status"), len(self.last_rows), elapsed)

    def _request_failed(self, error):
        self.running = False
        self.progress["value"] = 0
        self.error_requests += 1
        self._update_summary()
        self.request_end_label = datetime.now().strftime("%H:%M:%S")
        self._status(f"Project 03 | API Request - Failed - Start: {self.request_start_label} - End: {self.request_end_label}")
        self.response_text.delete("1.0", "end")
        self.response_text.insert("1.0", json.dumps({"error": error}, indent=2))
        self._log(error, "ERROR")
        messagebox.showerror("Request Error", error)

    def _update_summary(self):
        self.summary_var.set(f"Requests: {self.total_requests}   •   Success: {self.success_requests}   •   Errors: {self.error_requests}")

    def _flatten(self, value, prefix=""):
        result = {}
        if isinstance(value, dict):
            for key, val in value.items():
                new_key = f"{prefix}.{key}" if prefix else str(key)
                if isinstance(val, (dict, list)):
                    result.update(self._flatten(val, new_key))
                else:
                    result[new_key] = val
        elif isinstance(value, list):
            for idx, val in enumerate(value):
                new_key = f"{prefix}[{idx}]" if prefix else str(idx)
                if isinstance(val, (dict, list)):
                    result.update(self._flatten(val, new_key))
                else:
                    result[new_key] = val
        else:
            result[prefix or "value"] = value
        return result

    def _rows_from_json(self):
        data = self.last_json
        if data is None:
            return []
        if isinstance(data, dict):
            requested = self.flatten_var.get().strip()
            keys = [requested, "data", "results", "items", "users"] if requested else ["data", "results", "items", "users"]
            for key in keys:
                if key and isinstance(data.get(key), list):
                    data = data[key]
                    break
            else:
                data = [data]
        if isinstance(data, list):
            return [self._flatten(item) if isinstance(item, (dict, list)) else {"value": item} for item in data]
        return [{"value": data}]

    def preview_table(self, silent=False):
        rows = self._rows_from_json()
        self.last_rows = rows
        self._clear_tree()
        if not rows:
            if not silent:
                messagebox.showinfo("Preview", "No tabular data found in the response.")
            return
        columns = []
        for row in rows:
            for key in row:
                if key not in columns:
                    columns.append(key)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            width = max(120, min(280, 20 + max(len(str(col)), *(len(str(row.get(col, ""))) for row in rows)) * 8))
            self.tree.column(col, width=width, minwidth=100, anchor="w", stretch=False)
        self.tree.tag_configure("data", foreground=TEXT, background=WHITE)
        self.tree.tag_configure("alt", foreground=TEXT, background="#F4FBF6")
        for idx, row in enumerate(rows):
            self.tree.insert("", "end", values=[row.get(c, "") for c in columns], tags=("alt" if idx % 2 else "data",))
        self._log(f"Preview generated: {len(rows)} rows / {len(columns)} columns")

    def export_csv(self):
        if not self.last_rows:
            self.preview_table(silent=True)
        if not self.last_rows:
            messagebox.showwarning("Export", "No data available to export.")
            return
        path = filedialog.asksaveasfilename(initialdir=OUTPUT_DIR, initialfile="api_export.csv", defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        columns = list(dict.fromkeys(k for row in self.last_rows for k in row))
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=columns)
            writer.writeheader()
            writer.writerows(self.last_rows)
        self.current_output = Path(path)
        self._log(f"CSV exported: {path}")
        self._status(f"CSV exported | {Path(path).name}")
        messagebox.showinfo("Export Complete", f"CSV saved successfully.\n\n{path}")

    def export_excel(self):
        if Workbook is None:
            messagebox.showerror("Excel Export", "openpyxl is not installed. Run: pip install -r requirements.txt")
            return
        if not self.last_rows:
            self.preview_table(silent=True)
        if not self.last_rows:
            messagebox.showwarning("Export", "No data available to export.")
            return
        path = filedialog.asksaveasfilename(initialdir=OUTPUT_DIR, initialfile="api_export.xlsx", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        columns = list(dict.fromkeys(k for row in self.last_rows for k in row))
        wb = Workbook()
        ws = wb.active
        ws.title = "API Data"
        ws.append(columns)
        for row in self.last_rows:
            ws.append([row.get(c, "") for c in columns])
        summary = wb.create_sheet("Export Summary")
        summary_rows = [
            ("Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Rows", len(self.last_rows)),
            ("Columns", len(columns)),
            ("Endpoint", self.url_var.get()),
            ("Method", self.method_var.get()),
            ("Demo Mode", "Yes" if self.demo_var.get() else "No"),
        ]
        for row in summary_rows:
            summary.append(list(row))
        wb.save(path)
        self.current_output = Path(path)
        self._log(f"Excel exported: {path}")
        self._status(f"Excel exported | {Path(path).name}")
        messagebox.showinfo("Export Complete", f"Excel file saved successfully.\n\n{path}")

    def save_config(self):
        path = filedialog.asksaveasfilename(initialdir=CONFIG_DIR, initialfile="api_endpoint_config.json", defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not path:
            return
        c = self._read_request_values()
        config = {k: c[k] for k in ["method", "url", "auth", "token", "record_id", "page", "per_page", "retries", "timeout", "demo"]}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        self._log(f"Configuration saved: {path}")
        messagebox.showinfo("Configuration", "Configuration saved successfully.")

    def load_config(self):
        path = filedialog.askopenfilename(initialdir=CONFIG_DIR, filetypes=[("JSON files", "*.json")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                c = json.load(f)
            self.method_var.set(c.get("method", "GET"))
            self.url_var.set(c.get("url", ""))
            self.auth_type_var.set(c.get("auth", c.get("auth_type", "None")))
            self.token_var.set(c.get("token", ""))
            self.record_id_var.set(c.get("record_id", ""))
            self.page_var.set(str(c.get("page", 1)))
            self.per_page_var.set(str(c.get("per_page", 10)))
            self.retries_var.set(str(c.get("retries", 3)))
            self.timeout_var.set(str(c.get("timeout", 20)))
            self.demo_var.set(bool(c.get("demo", c.get("demo_mode", False))))
            # Apply the loaded HTTP method to the request-body control.
            # GET/DELETE must never retain the previous POST/PUT JSON body.
            self._method_changed()
            self._mode_changed()
            self._log(f"Configuration loaded: {path}")
            self._status("Configuration loaded")
        except Exception as exc:
            messagebox.showerror("Configuration", f"Could not load configuration.\n\n{exc}")

    def open_output(self):
        OUTPUT_DIR.mkdir(exist_ok=True)
        try:
            os.startfile(str(self.current_output.parent if self.current_output else OUTPUT_DIR))
        except AttributeError:
            webbrowser.open(OUTPUT_DIR.as_uri())

    def show_about(self):
        messagebox.showinfo("About", f"AKASH TECH SOLUTIONS\n\nProject 03 — API Automation & Data Integration\nVersion {APP_VERSION}\n\nPython desktop automation tool for REST APIs, JSON transformation, exports and repeatable workflows.")

    def show_help(self):
        messagebox.showinfo("Help", "1. Start in Demo Mode.\n2. Select GET/POST/PUT/DELETE.\n3. Click Send Request.\n4. Review Response.\n5. Use Transform & Export for CSV/Excel.\n6. Save configurations for repeat use.\n7. Turn Demo Mode off only when testing a real API.")

    def show_settings(self):
        messagebox.showinfo("Settings", f"Output folder:\n{OUTPUT_DIR}\n\nLog folder:\n{LOG_DIR}\n\nConfiguration folder:\n{CONFIG_DIR}")


if __name__ == "__main__":
    app = APIAutomationApp()
    app.mainloop()
