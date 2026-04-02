"""
ZiiPOS Menu Converter V2.0
- Tab 1: Excel Import (reuses V1 logic from Menu_Converter.py)
- Tab 2: AI Import (OCR images/PDF/Word, parse, translate, export)
"""

import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, askopenfilenames, askdirectory
import sys
import os
import json
import threading
import traceback
import pandas as pd
from datetime import datetime

from i18n import t, LANG_CODES, lang_display_name
from Menu_Converter import (
    DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_FILE,
    ensure_template, normalize_source_columns,
    processMenuGroup, processCategory, processItem,
    _safe, _num, _bool,
)

SUPPORTED_IMG_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp")
SUPPORTED_DOC_EXT = (".pdf", ".docx", ".doc")
AI_FILE_TYPES = [
    ("All Supported", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp *.pdf *.docx *.xlsx *.xls"),
    ("Images", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
    ("PDF", "*.pdf"),
    ("Word", "*.docx"),
    ("Excel", "*.xlsx *.xls"),
]

DESC_SLOTS = ["Description1", "Description2", "Description3", "Description4"]
LANG_OPTIONS = ["en", "cn", "jp", "kr", "vi", "th"]
LANG_NONE = ""

CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ZiiPOSMenuConverter")
CONFIG_FILE = os.path.join(CONFIG_DIR, "settings.json")

def _load_config() -> dict:
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def _save_config(cfg: dict):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[WARN] Failed to save config: {e}")


class App:
    def __init__(self, root):
        self.root = root
        self._config = _load_config()
        self.ui_lang = self._config.get("ui_lang", "en")
        self._build_ui()

    def _build_ui(self):
        root = self.root
        root.title(t("title", self.ui_lang))
        w, h = 780, 680
        sx = root.winfo_screenwidth()
        sy = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sx-w)//2}+{(sy-h)//2}")
        root.resizable(True, True)

        for child in root.winfo_children():
            child.destroy()

        self.ft = tkFont.Font(family="Segoe UI", size=10)
        self.ft_bold = tkFont.Font(family="Segoe UI", size=10, weight="bold")

        top_frame = tk.Frame(root)
        top_frame.pack(fill="x", padx=10, pady=(8, 0))

        tk.Label(top_frame, text="UI:", font=self.ft).pack(side="left")
        for code in ("en", "cn", "jp"):
            btn = tk.Button(
                top_frame,
                text=code.upper(),
                font=self.ft,
                width=4,
                relief="sunken" if code == self.ui_lang else "raised",
                command=lambda c=code: self._switch_lang(c),
            )
            btn.pack(side="left", padx=2)

        tk.Label(
            top_frame,
            text=t("title", self.ui_lang),
            font=tkFont.Font(family="Segoe UI", size=12, weight="bold"),
        ).pack(side="right")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab1 = tk.Frame(self.notebook)
        self.tab2 = tk.Frame(self.notebook)
        self.notebook.add(self.tab1, text=t("tab_excel", self.ui_lang))
        self.notebook.add(self.tab2, text=t("tab_ai", self.ui_lang))

        self._build_tab1()
        self._build_tab2()

        self.notebook.select(self.tab2)

    def _switch_lang(self, lang):
        self.ui_lang = lang
        self._config["ui_lang"] = lang
        _save_config(self._config)
        self._build_ui()

    # ──────────────────────────────────────────────────────────
    # Tab 1: Excel Import (V1 logic)
    # ──────────────────────────────────────────────────────────
    def _build_tab1(self):
        frame = self.tab1
        ft = self.ft
        y = 20

        def _label(text, row):
            tk.Label(frame, text=text, font=ft, anchor="w").place(
                x=20, y=row, width=120, height=30
            )

        def _entry(row, default=""):
            e = tk.Entry(frame, font=ft, borderwidth=1)
            if default:
                e.insert(0, default)
            e.place(x=150, y=row, width=380, height=30)
            return e

        def _btn(text, row, cmd):
            tk.Button(frame, text=text, font=ft, command=cmd).place(
                x=540, y=row, width=80, height=30
            )

        _label(t("lbl_menu_file", self.ui_lang), y)
        self.t1_ent_menu = _entry(y)
        _btn(t("btn_select", self.ui_lang), y, self._t1_sel_menu)

        y += 42
        _label(t("lbl_output_path", self.ui_lang), y)
        self.t1_ent_out = _entry(y, DEFAULT_OUTPUT_DIR)
        _btn(t("btn_browse", self.ui_lang), y, self._t1_sel_out)

        y += 60
        tk.Button(
            frame, text=t("btn_convert", self.ui_lang), font=ft,
            bg="#4CAF50", fg="white", width=14,
            command=self._t1_convert,
        ).place(x=200, y=y, width=140, height=42)

        tk.Button(
            frame, text=t("btn_close", self.ui_lang), font=ft,
            width=14, command=sys.exit,
        ).place(x=360, y=y, width=140, height=42)

    def _t1_sel_menu(self):
        f = askopenfilename(filetypes=[
            (t("file_types_excel", self.ui_lang), "*.xlsx *.xls")
        ])
        if f:
            self.t1_ent_menu.delete(0, tk.END)
            self.t1_ent_menu.insert(0, f)

    def _t1_sel_out(self):
        d = askdirectory(initialdir=self.t1_ent_out.get())
        if d:
            self.t1_ent_out.delete(0, tk.END)
            self.t1_ent_out.insert(0, d)

    def _t1_convert(self):
        source = self.t1_ent_menu.get().strip()
        output = self.t1_ent_out.get().strip() or DEFAULT_OUTPUT_DIR
        if not source:
            messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
            return
        if not ensure_template(DEFAULT_TEMPLATE_FILE):
            return
        try:
            from Menu_Converter import processMenu
            out = processMenu(source, DEFAULT_TEMPLATE_FILE, output)
            messagebox.showinfo("Done", t("msg_done", self.ui_lang).format(out))
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", t("msg_err", self.ui_lang).format(e))

    # ──────────────────────────────────────────────────────────
    # Tab 2: AI Import
    # ──────────────────────────────────────────────────────────
    def _build_tab2(self):
        frame = self.tab2
        ft = self.ft

        top = tk.Frame(frame)
        top.pack(fill="x", padx=10, pady=(10, 0))

        # Row 1: Source file
        r1 = tk.Frame(top)
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text=t("lbl_source_file", self.ui_lang), font=ft, width=14, anchor="w").pack(side="left")
        self.t2_ent_file = tk.Entry(r1, font=ft, borderwidth=1)
        self.t2_ent_file.pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r1, text=t("btn_select", self.ui_lang), font=ft, command=self._t2_sel_file).pack(side="left")

        # Row 2: Source language + Translation mode
        r2 = tk.Frame(top)
        r2.pack(fill="x", pady=2)
        tk.Label(r2, text=t("lbl_source_lang", self.ui_lang), font=ft, width=14, anchor="w").pack(side="left")
        self.t2_src_lang = ttk.Combobox(r2, font=ft, width=12, state="readonly")
        src_vals = ["auto"] + LANG_OPTIONS
        self.t2_src_lang["values"] = [
            lang_display_name(c, self.ui_lang) if c != "auto" else t("lang_auto", self.ui_lang)
            for c in src_vals
        ]
        self.t2_src_lang.current(0)
        self.t2_src_lang.pack(side="left", padx=(4, 16))
        self._src_lang_codes = src_vals

        tk.Label(r2, text=t("lbl_trans_mode", self.ui_lang), font=ft, anchor="w").pack(side="left")
        self.t2_trans_mode = ttk.Combobox(r2, font=ft, width=14, state="readonly")
        self.t2_trans_mode["values"] = [
            t("trans_offline", self.ui_lang),
            t("trans_online", self.ui_lang),
        ]
        self.t2_trans_mode.current(0)
        self.t2_trans_mode.pack(side="left", padx=(4, 8))

        tk.Label(r2, text=t("lbl_api_key", self.ui_lang), font=ft, anchor="w").pack(side="left")
        self.t2_api_provider = ttk.Combobox(r2, font=ft, width=10, state="readonly")
        self.t2_api_provider["values"] = ["OpenAI", "DeepSeek", "Doubao", "Claude"]
        saved_provider = self._config.get("api_provider", "OpenAI")
        provider_list = list(self.t2_api_provider["values"])
        if saved_provider in provider_list:
            self.t2_api_provider.current(provider_list.index(saved_provider))
        else:
            self.t2_api_provider.current(0)
        self.t2_api_provider.bind("<<ComboboxSelected>>", self._on_provider_change)
        self.t2_api_provider.pack(side="left", padx=(4, 4))

        self.t2_api_key = tk.Entry(r2, font=ft, width=16, show="*", borderwidth=1)
        saved_key = self._config.get("api_keys", {}).get(
            self.t2_api_provider.get(), ""
        )
        if saved_key:
            self.t2_api_key.insert(0, saved_key)
        self.t2_api_key.pack(side="left", padx=(0, 2))

        tk.Button(r2, text="💾", font=ft, width=2,
                  command=self._save_api_key).pack(side="left", padx=(0, 0))

        # Row 3: Output path
        r3 = tk.Frame(top)
        r3.pack(fill="x", pady=2)
        tk.Label(r3, text=t("lbl_output_path", self.ui_lang), font=ft, width=14, anchor="w").pack(side="left")
        self.t2_ent_out = tk.Entry(r3, font=ft, borderwidth=1)
        self.t2_ent_out.insert(0, DEFAULT_OUTPUT_DIR)
        self.t2_ent_out.pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r3, text=t("btn_browse", self.ui_lang), font=ft, command=self._t2_sel_out).pack(side="left")

        # Row 4: Description language config
        r4 = tk.LabelFrame(top, text=t("lbl_desc_config", self.ui_lang), font=ft)
        r4.pack(fill="x", pady=(8, 4))

        self.t2_desc_combos = []
        desc_inner = tk.Frame(r4)
        desc_inner.pack(fill="x", padx=8, pady=4)

        desc_defaults = ["en", "cn", "", ""]
        for i in range(4):
            lbl_key = f"lbl_desc{i+1}"
            tk.Label(desc_inner, text=t(lbl_key, self.ui_lang), font=ft, anchor="w").pack(side="left", padx=(0, 2))
            combo = ttk.Combobox(desc_inner, font=ft, width=10, state="readonly")
            lang_vals = [LANG_NONE] + LANG_OPTIONS
            combo["values"] = [t("lang_none", self.ui_lang)] + [
                lang_display_name(c, self.ui_lang) for c in LANG_OPTIONS
            ]
            default_idx = 0
            if desc_defaults[i] in LANG_OPTIONS:
                default_idx = LANG_OPTIONS.index(desc_defaults[i]) + 1
            combo.current(default_idx)
            combo.pack(side="left", padx=(0, 12))
            self.t2_desc_combos.append(combo)

        # Button row
        r5 = tk.Frame(top)
        r5.pack(fill="x", pady=(6, 2))
        tk.Button(
            r5, text=t("btn_read", self.ui_lang), font=self.ft_bold,
            bg="#2196F3", fg="white", command=self._t2_read,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            r5, text=t("btn_save_simple", self.ui_lang), font=self.ft_bold,
            bg="#FF9800", fg="white", command=self._t2_save_simple,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            r5, text=t("btn_export", self.ui_lang), font=self.ft_bold,
            bg="#4CAF50", fg="white", command=self._t2_export,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            r5, text=t("btn_add_row", self.ui_lang), font=ft,
            command=self._t2_add_row,
        ).pack(side="left", padx=(0, 4))
        tk.Button(
            r5, text=t("btn_del_row", self.ui_lang), font=ft,
            command=self._t2_del_row,
        ).pack(side="left")

        # Status label
        self.t2_status = tk.Label(top, text="", font=ft, fg="#666", anchor="w")
        self.t2_status.pack(fill="x", pady=(2, 0))

        # Treeview preview table
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        columns = ("itemcode", "name", "price", "category")
        self.t2_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)

        col_headers = {
            "itemcode": t("col_itemcode", self.ui_lang),
            "name": t("col_name", self.ui_lang),
            "price": t("col_price", self.ui_lang),
            "category": t("col_category", self.ui_lang),
        }
        col_widths = {"itemcode": 80, "name": 300, "price": 80, "category": 150}

        for col in columns:
            self.t2_tree.heading(col, text=col_headers[col])
            self.t2_tree.column(col, width=col_widths[col], anchor="w" if col == "name" else "center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.t2_tree.yview)
        self.t2_tree.configure(yscrollcommand=vsb.set)
        self.t2_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.t2_tree.bind("<Double-1>", self._t2_edit_cell)

    def _on_provider_change(self, event=None):
        """Load saved API key when provider selection changes."""
        provider = self.t2_api_provider.get()
        saved_key = self._config.get("api_keys", {}).get(provider, "")
        self.t2_api_key.delete(0, tk.END)
        if saved_key:
            self.t2_api_key.insert(0, saved_key)

    def _save_api_key(self):
        """Save current API key for the selected provider."""
        provider = self.t2_api_provider.get()
        key = self.t2_api_key.get().strip()
        if "api_keys" not in self._config:
            self._config["api_keys"] = {}
        self._config["api_keys"][provider] = key
        self._config["api_provider"] = provider
        _save_config(self._config)
        messagebox.showinfo("OK", f"{provider} API key saved.\n{provider} API 密钥已保存。")

    def _t2_sel_file(self):
        files = askopenfilenames(filetypes=AI_FILE_TYPES)
        if files:
            self.t2_ent_file.delete(0, tk.END)
            self.t2_ent_file.insert(0, " ; ".join(files))

    def _t2_sel_out(self):
        d = askdirectory(initialdir=self.t2_ent_out.get())
        if d:
            self.t2_ent_out.delete(0, tk.END)
            self.t2_ent_out.insert(0, d)

    def _t2_get_src_lang(self) -> str:
        idx = self.t2_src_lang.current()
        return self._src_lang_codes[idx]

    def _t2_get_desc_langs(self) -> list[str]:
        """Return list of language codes for Description1-4 slots."""
        result = []
        for combo in self.t2_desc_combos:
            idx = combo.current()
            if idx <= 0:
                result.append("")
            else:
                result.append(LANG_OPTIONS[idx - 1])
        return result

    def _t2_set_status(self, msg: str):
        self.t2_status.config(text=msg)
        self.root.update_idletasks()

    def _t2_get_file_list(self) -> list[str]:
        """Parse the entry field into a list of file paths (semicolon-separated)."""
        raw = self.t2_ent_file.get().strip()
        if not raw:
            return []
        return [p.strip() for p in raw.split(";") if p.strip()]

    def _t2_is_online_mode(self) -> bool:
        return self.t2_trans_mode.current() == 1

    def _t2_read(self):
        """Read & Parse: OCR/extract text from all selected files, fill preview table."""
        file_list = self._t2_get_file_list()
        if not file_list:
            messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
            return

        use_vision = self._t2_is_online_mode()
        api_key = self.t2_api_key.get().strip()
        api_provider = self.t2_api_provider.get()

        if use_vision and not api_key:
            messagebox.showerror("Error",
                "Online mode requires an API key.\n在线模式需要 API 密钥。")
            return

        self._t2_set_status(t("msg_reading", self.ui_lang))

        def do_read():
            try:
                src_lang = self._t2_get_src_lang()
                all_items = []
                detected_lang = src_lang if src_lang != "auto" else "jp"

                for idx, filepath in enumerate(file_list):
                    fname = os.path.basename(filepath)
                    self.root.after(0, lambda f=fname, i=idx: self._t2_set_status(
                        f"({i+1}/{len(file_list)}) {f} ..."
                    ))
                    ext = os.path.splitext(filepath)[1].lower()
                    is_image_or_pdf = ext in (
                        ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif",
                        ".webp", ".pdf"
                    )

                    if use_vision and is_image_or_pdf:
                        from ai_reader import read_file_vision, detect_language
                        vision_items = read_file_vision(
                            filepath, api_key, api_provider, detected_lang
                        )
                        if vision_items:
                            offset = len(all_items)
                            for i, vi in enumerate(vision_items):
                                all_items.append({
                                    "itemcode": "%04d" % (offset + i + 1),
                                    "name": vi.get("name", ""),
                                    "price": float(vi.get("price", 0) or 0),
                                    "category": vi.get("category", "Default"),
                                })
                            continue

                    from ai_reader import read_file, detect_language
                    from menu_parser import parse_menu_text

                    lang_hint = src_lang if src_lang != "auto" else "jp"
                    raw_text = read_file(filepath, lang_hint)

                    if src_lang == "auto" and raw_text:
                        detected_lang = detect_language(raw_text)
                        print(f"[INFO] {fname}: auto-detected language = {detected_lang}")

                    self.root.after(0, lambda: self._t2_set_status(t("msg_parsing", self.ui_lang)))
                    items = parse_menu_text(raw_text, detected_lang)

                    offset = len(all_items)
                    for item in items:
                        item["itemcode"] = "%04d" % (offset + int(item["itemcode"]))
                    all_items.extend(items)

                self.root.after(0, lambda: self._t2_populate_tree(all_items, detected_lang))
                self.root.after(0, lambda: self._t2_set_status(
                    f"OK - {len(all_items)} items from {len(file_list)} file(s)"
                    + (f" [Vision API: {api_provider}]" if use_vision else " [Offline OCR]")
                ))
            except Exception as e:
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", t("msg_err", self.ui_lang).format(e)
                ))
                self.root.after(0, lambda: self._t2_set_status(""))

        threading.Thread(target=do_read, daemon=True).start()

    def _t2_populate_tree(self, items: list[dict], detected_lang: str = ""):
        """Fill the treeview with parsed items."""
        for child in self.t2_tree.get_children():
            self.t2_tree.delete(child)

        for item in items:
            self.t2_tree.insert("", "end", values=(
                item.get("itemcode", ""),
                item.get("name", ""),
                item.get("price", 0),
                item.get("category", "Default"),
            ))

        self._detected_src_lang = detected_lang

    def _t2_add_row(self):
        existing = self.t2_tree.get_children()
        next_code = "%04d" % (len(existing) + 1)
        self.t2_tree.insert("", "end", values=(next_code, "", 0, "Default"))

    def _t2_del_row(self):
        selected = self.t2_tree.selection()
        for item in selected:
            self.t2_tree.delete(item)

    def _t2_edit_cell(self, event):
        """Double-click to edit a cell in the treeview."""
        item_id = self.t2_tree.identify_row(event.y)
        column = self.t2_tree.identify_column(event.x)
        if not item_id or not column:
            return

        col_idx = int(column.replace("#", "")) - 1
        col_key = ("itemcode", "name", "price", "category")[col_idx]

        bbox = self.t2_tree.bbox(item_id, column)
        if not bbox:
            return
        x, y, w, h = bbox

        current_val = self.t2_tree.item(item_id, "values")[col_idx]

        entry = tk.Entry(self.t2_tree, font=self.ft)
        entry.insert(0, current_val)
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def save_edit(e=None):
            new_val = entry.get()
            values = list(self.t2_tree.item(item_id, "values"))
            values[col_idx] = new_val
            self.t2_tree.item(item_id, values=values)
            entry.destroy()

        def cancel_edit(e=None):
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<Escape>", cancel_edit)
        entry.bind("<FocusOut>", save_edit)

    def _t2_save_simple(self):
        """Save preview table as a simple menuCollection-format Excel for easy editing."""
        children = self.t2_tree.get_children()
        if not children:
            messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
            return

        output_dir = self.t2_ent_out.get().strip() or DEFAULT_OUTPUT_DIR
        desc_langs = self._t2_get_desc_langs()
        src_lang = getattr(self, "_detected_src_lang", "")
        if not src_lang:
            src_lang = self._t2_get_src_lang()
            if src_lang == "auto":
                src_lang = "en"

        trans_mode_idx = self.t2_trans_mode.current()
        trans_mode = "offline" if trans_mode_idx == 0 else "online"
        api_key = self.t2_api_key.get().strip()
        api_provider = self.t2_api_provider.get()

        self._t2_set_status(t("msg_translating", self.ui_lang))

        def do_save():
            try:
                from translator import Translator
                translator = Translator(mode=trans_mode, api_key=api_key,
                                        provider=api_provider)

                SIMPLE_COLUMNS = [
                    "ItemCode", "Description1", "Description2", "Description3",
                    "Description4", "Category", "MenuGroup", "TaxRate",
                    "Price", "Price1", "Price2", "Price3",
                    "SubDescription", "SubDescription1", "SubDescription2",
                    "SubDescription3", "ItemGroup", "Instruction",
                ]

                rows = []
                for child in children:
                    vals = self.t2_tree.item(child, "values")
                    name = str(vals[1])
                    price = float(vals[2]) if vals[2] else 0

                    descs = {}
                    for i, lang_code in enumerate(desc_langs):
                        key = f"Description{i+1}"
                        if not lang_code:
                            descs[key] = ""
                        elif lang_code == src_lang:
                            descs[key] = name
                        else:
                            descs[key] = translator.translate(name, src_lang, lang_code)

                    rows.append({
                        "ItemCode": vals[0],
                        "Description1": descs.get("Description1", name),
                        "Description2": descs.get("Description2", ""),
                        "Description3": descs.get("Description3", ""),
                        "Description4": descs.get("Description4", ""),
                        "Category": str(vals[3]),
                        "MenuGroup": "Default",
                        "TaxRate": 10,
                        "Price": price,
                        "Price1": 0,
                        "Price2": 0,
                        "Price3": 0,
                        "SubDescription": "",
                        "SubDescription1": "",
                        "SubDescription2": "",
                        "SubDescription3": "",
                        "ItemGroup": "OTHERS",
                        "Instruction": False,
                    })

                os.makedirs(output_dir, exist_ok=True)
                date_str = datetime.now().strftime("%Y%m%d%H%M%S")
                output_file = os.path.join(
                    output_dir, f"menuCollection-{date_str}.xlsx"
                )

                df = pd.DataFrame(rows, columns=SIMPLE_COLUMNS)
                df.to_excel(output_file, index=False, sheet_name="MenuItem")
                print(f"[INFO] Simple Excel saved to: {output_file}")

                self.root.after(0, lambda: messagebox.showinfo(
                    "Done", t("msg_done", self.ui_lang).format(output_file)
                ))
                self.root.after(0, lambda: self._t2_set_status(
                    f"OK - saved {output_file}"
                ))
            except Exception as e:
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", t("msg_err", self.ui_lang).format(e)
                ))
                self.root.after(0, lambda: self._t2_set_status(""))

        threading.Thread(target=do_save, daemon=True).start()

    def _t2_export(self):
        """Translate descriptions and export to ZiiPOS template format."""
        children = self.t2_tree.get_children()
        if not children:
            messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
            return

        if not ensure_template(DEFAULT_TEMPLATE_FILE):
            return

        output_dir = self.t2_ent_out.get().strip() or DEFAULT_OUTPUT_DIR
        desc_langs = self._t2_get_desc_langs()
        src_lang = getattr(self, "_detected_src_lang", "")
        if not src_lang:
            src_lang = self._t2_get_src_lang()
            if src_lang == "auto":
                src_lang = "en"

        trans_mode_idx = self.t2_trans_mode.current()
        trans_mode = "offline" if trans_mode_idx == 0 else "online"
        api_key = self.t2_api_key.get().strip()
        api_provider = self.t2_api_provider.get()

        self._t2_set_status(t("msg_translating", self.ui_lang))

        def do_export():
            try:
                from translator import Translator

                translator = Translator(mode=trans_mode, api_key=api_key,
                                        provider=api_provider)

                rows = []
                for child in children:
                    vals = self.t2_tree.item(child, "values")
                    rows.append({
                        "itemcode": vals[0],
                        "name": str(vals[1]),
                        "price": float(vals[2]) if vals[2] else 0,
                        "category": str(vals[3]),
                    })

                for row in rows:
                    name = row["name"]
                    row["descriptions"] = {}

                    for i, lang_code in enumerate(desc_langs):
                        if not lang_code:
                            row["descriptions"][f"Description{i+1}"] = ""
                            continue
                        if lang_code == src_lang:
                            row["descriptions"][f"Description{i+1}"] = name
                        else:
                            self.root.after(0, lambda: self._t2_set_status(
                                t("msg_translating", self.ui_lang)
                            ))
                            translated = translator.translate(name, src_lang, lang_code)
                            row["descriptions"][f"Description{i+1}"] = translated

                self.root.after(0, lambda: self._t2_set_status(t("msg_exporting", self.ui_lang)))
                output_file = self._build_and_write_template(rows, desc_langs, output_dir)

                self.root.after(0, lambda: messagebox.showinfo(
                    "Done", t("msg_done", self.ui_lang).format(output_file)
                ))
                self.root.after(0, lambda: self._t2_set_status(
                    f"OK - exported to {output_file}"
                ))
            except Exception as e:
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", t("msg_err", self.ui_lang).format(e)
                ))
                self.root.after(0, lambda: self._t2_set_status(""))

        threading.Thread(target=do_export, daemon=True).start()

    def _build_and_write_template(self, rows: list[dict], desc_langs: list[str],
                                  output_dir: str) -> str:
        """
        Build a source DataFrame from AI-parsed rows and run V1 pipeline
        to produce the full 13-sheet ZiiPOS template output.
        """
        os.makedirs(output_dir, exist_ok=True)

        source_rows = []
        for row in rows:
            src_row = {
                "ItemCode": row["itemcode"],
                "Description1": row["descriptions"].get("Description1", row["name"]),
                "Description2": row["descriptions"].get("Description2", ""),
                "Description3": row["descriptions"].get("Description3", ""),
                "Description4": row["descriptions"].get("Description4", ""),
                "Category": row["category"],
                "MenuGroup": "Default",
                "TaxRate": 10,
                "Price": row["price"],
                "Price2": 0,
                "Price3": 0,
                "Instruction": False,
            }
            source_rows.append(src_row)

        source = pd.DataFrame(source_rows)
        source = normalize_source_columns(source)

        tpl = pd.ExcelFile(DEFAULT_TEMPLATE_FILE)
        sheets = {name: pd.read_excel(tpl, name) for name in tpl.sheet_names}

        sheets["MenuGroupTable"], mg_code_map = processMenuGroup(
            source, sheets["MenuGroupTable"]
        )
        sheets["Category"] = processCategory(source, sheets["Category"], mg_code_map)
        sheets["MenuItem"] = processItem(source, sheets["MenuItem"], "00")

        date_str = datetime.now().strftime("%Y%m%d%H%M%S")
        output_file = os.path.join(output_dir, f"export_FullMenu-{date_str}.xlsx")

        writer = pd.ExcelWriter(
            output_file, engine="xlsxwriter",
            engine_kwargs={"options": {
                "strings_to_numbers": False,
                "strings_to_formulas": False,
            }}
        )
        text_fmt = writer.book.add_format({"num_format": "@"})

        overview = pd.DataFrame({"ExportTime": [datetime.now().strftime("%Y-%m-%d %H:%M")]})
        overview.to_excel(writer, sheet_name="Overview", index=False)

        for name in tpl.sheet_names:
            if name == "Overview":
                continue
            df = sheets[name]

            if name == "MenuGroupTable":
                df["Code"] = df["Code"].apply(
                    lambda v: "%02d" % int(v) if pd.notna(v) else "00"
                )
            if name == "Course":
                df["CourseCode"] = ["%02d" % (i + 1) for i in range(len(df))]
            if name == "Category":
                df["Code"] = df["Code"].apply(
                    lambda v: "%03d" % int(v) if pd.notna(v) else "001"
                )
                df["MenuGroupCode"] = df["MenuGroupCode"].apply(
                    lambda v: str(v).strip() if pd.notna(v) else "00"
                )
                df["MenuGroupList"] = df["MenuGroupList"].apply(
                    lambda v: str(v).strip() if pd.notna(v) else "00"
                )

            df.to_excel(writer, sheet_name=name, index=False)

            if name == "MenuGroupTable":
                writer.sheets[name].set_column("A:A", None, text_fmt)
            if name == "Course":
                writer.sheets[name].set_column("A:A", None, text_fmt)
            if name == "Category":
                ws = writer.sheets[name]
                ws.set_column("A:A", None, text_fmt)
                ws.set_column("B:B", None, text_fmt)

        writer.close()
        print(f"[INFO] AI Export saved to: {output_file}")
        return output_file


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
