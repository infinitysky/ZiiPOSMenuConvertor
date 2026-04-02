"""
ZiiPOS Menu Converter V4.0
- Tab 1: Excel Import (reuses V1 logic from Menu_Converter.py)
- Tab 2: PE Menu Import (parse PE POS export, extract images, convert to ZiiPOS)
"""

import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory
import sys
import os
import json
import threading
import traceback
import pandas as pd
from datetime import datetime

from i18n import t
from Menu_Converter import (
    DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_FILE,
    ensure_template, normalize_source_columns,
    processMenuGroup, processCategory, processItem,
)

CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "ZiiPOSMenuConverter",
)
CONFIG_FILE = os.path.join(CONFIG_DIR, "settings_v4.json")


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

    # ══════════════════════════════════════════════════════════
    # UI Construction
    # ══════════════════════════════════════════════════════════
    def _build_ui(self):
        root = self.root
        root.title(t("title", self.ui_lang))
        w, h = 860, 680
        sx = root.winfo_screenwidth()
        sy = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sx - w) // 2}+{(sy - h) // 2}")
        root.resizable(True, True)

        for child in root.winfo_children():
            child.destroy()

        self.ft = tkFont.Font(family="Segoe UI", size=10)
        self.ft_bold = tkFont.Font(family="Segoe UI", size=10, weight="bold")

        # ── Top bar ──
        top_frame = tk.Frame(root)
        top_frame.pack(fill="x", padx=10, pady=(8, 0))

        tk.Label(top_frame, text="UI:", font=self.ft).pack(side="left")
        for code in ("en", "cn", "jp"):
            btn = tk.Button(
                top_frame, text=code.upper(), font=self.ft, width=4,
                relief="sunken" if code == self.ui_lang else "raised",
                command=lambda c=code: self._switch_lang(c),
            )
            btn.pack(side="left", padx=2)

        tk.Label(
            top_frame, text=t("title", self.ui_lang),
            font=tkFont.Font(family="Segoe UI", size=12, weight="bold"),
        ).pack(side="right")

        # ── Notebook tabs ──
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_excel = tk.Frame(self.notebook)
        self.tab_pe = tk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text=t("tab_excel", self.ui_lang))
        self.notebook.add(self.tab_pe, text=t("tab_pe", self.ui_lang))

        self._build_tab_excel()
        self._build_tab_pe()

        self.notebook.select(self.tab_excel)

    def _switch_lang(self, lang):
        self.ui_lang = lang
        self._config["ui_lang"] = lang
        _save_config(self._config)
        self._build_ui()

    # ══════════════════════════════════════════════════════════
    # Tab 1: Excel Import (V1 logic)
    # ══════════════════════════════════════════════════════════
    def _build_tab_excel(self):
        frame = self.tab_excel
        ft = self.ft
        y = 20

        def _label(text, row):
            tk.Label(frame, text=text, font=ft, anchor="w").place(
                x=20, y=row, width=120, height=30,
            )

        def _entry(row, default=""):
            e = tk.Entry(frame, font=ft, borderwidth=1)
            if default:
                e.insert(0, default)
            e.place(x=150, y=row, width=440, height=30)
            return e

        def _btn(text, row, cmd):
            tk.Button(frame, text=text, font=ft, command=cmd).place(
                x=600, y=row, width=80, height=30,
            )

        _label(t("lbl_menu_file", self.ui_lang), y)
        self.t1_ent_menu = _entry(y)
        _btn(t("btn_select", self.ui_lang), y, self._t1_sel_menu)

        y += 42
        _label(t("lbl_output_folder", self.ui_lang), y)
        self.t1_ent_out = _entry(y, DEFAULT_OUTPUT_DIR)
        _btn(t("btn_browse", self.ui_lang), y, self._t1_sel_out)

        y += 60
        tk.Button(
            frame, text=t("btn_convert", self.ui_lang), font=ft,
            bg="#4CAF50", fg="white", width=14, command=self._t1_convert,
        ).place(x=200, y=y, width=140, height=42)

        tk.Button(
            frame, text=t("btn_close", self.ui_lang), font=ft,
            width=14, command=sys.exit,
        ).place(x=360, y=y, width=140, height=42)

    def _t1_sel_menu(self):
        f = askopenfilename(filetypes=[
            (t("file_types_excel", self.ui_lang), "*.xlsx *.xls"),
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

    # ══════════════════════════════════════════════════════════
    # Tab 2: PE Menu Import
    # ══════════════════════════════════════════════════════════
    def _build_tab_pe(self):
        frame = self.tab_pe
        ft = self.ft

        top = tk.Frame(frame)
        top.pack(fill="x", padx=10, pady=(10, 0))

        # Row 1: PE file selection
        r1 = tk.Frame(top)
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text=t("lbl_pe_file", self.ui_lang), font=ft,
                 width=14, anchor="w").pack(side="left")
        self.pe_ent_file = tk.Entry(r1, font=ft, borderwidth=1)
        self.pe_ent_file.pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r1, text=t("btn_select", self.ui_lang), font=ft,
                  command=self._pe_sel_file).pack(side="left")

        # Row 2: Output folder
        r2 = tk.Frame(top)
        r2.pack(fill="x", pady=2)
        tk.Label(r2, text=t("lbl_output_folder", self.ui_lang), font=ft,
                 width=14, anchor="w").pack(side="left")
        self.pe_ent_out = tk.Entry(r2, font=ft, borderwidth=1)
        self.pe_ent_out.insert(0, DEFAULT_OUTPUT_DIR)
        self.pe_ent_out.pack(side="left", fill="x", expand=True, padx=(4, 4))
        tk.Button(r2, text=t("btn_browse", self.ui_lang), font=ft,
                  command=self._pe_sel_out).pack(side="left")

        # Action buttons
        r3 = tk.Frame(top)
        r3.pack(fill="x", pady=(8, 2))
        tk.Button(
            r3, text={"en": "Read & Preview", "cn": "读取预览", "jp": "読取・プレビュー"}.get(self.ui_lang, "Read & Preview"),
            font=self.ft_bold, bg="#2196F3", fg="white",
            command=self._pe_read,
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            r3, text=t("btn_pe_convert", self.ui_lang),
            font=self.ft_bold, bg="#4CAF50", fg="white",
            command=self._pe_convert,
        ).pack(side="left", padx=(0, 8))

        # Status
        self.pe_status = tk.Label(top, text="", font=ft, fg="#666", anchor="w")
        self.pe_status.pack(fill="x", pady=(2, 0))

        # Preview table
        tree_frame = tk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        columns = ("itemcode", "name", "price", "category", "menugroup", "tax", "status", "image")
        self.pe_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)

        col_headers = {
            "itemcode":  t("col_itemcode", self.ui_lang),
            "name":      t("col_name", self.ui_lang),
            "price":     t("col_price", self.ui_lang),
            "category":  t("col_category", self.ui_lang),
            "menugroup": t("col_menugroup", self.ui_lang),
            "tax":       t("col_tax", self.ui_lang),
            "status":    t("col_status", self.ui_lang),
            "image":     t("col_has_image", self.ui_lang),
        }
        col_widths = {
            "itemcode": 60, "name": 220, "price": 70, "category": 140,
            "menugroup": 120, "tax": 50, "status": 50, "image": 50,
        }

        for col in columns:
            self.pe_tree.heading(col, text=col_headers.get(col, col))
            anchor = "w" if col in ("name", "category", "menugroup") else "center"
            self.pe_tree.column(col, width=col_widths.get(col, 80), anchor=anchor)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.pe_tree.yview)
        self.pe_tree.configure(yscrollcommand=vsb.set)
        self.pe_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._pe_items = []
        self._pe_image_rows = set()

    def _pe_sel_file(self):
        f = askopenfilename(filetypes=[
            (t("file_types_pe", self.ui_lang), "*.xlsx *.xls"),
        ])
        if f:
            self.pe_ent_file.delete(0, tk.END)
            self.pe_ent_file.insert(0, f)

    def _pe_sel_out(self):
        d = askdirectory(initialdir=self.pe_ent_out.get())
        if d:
            self.pe_ent_out.delete(0, tk.END)
            self.pe_ent_out.insert(0, d)

    def _pe_set_status(self, msg: str):
        self.pe_status.config(text=msg)
        self.root.update_idletasks()

    def _pe_read(self):
        """Read PE Menu file and populate preview table."""
        filepath = self.pe_ent_file.get().strip()
        if not filepath:
            messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
            return

        self._pe_set_status(t("msg_reading", self.ui_lang))

        def do_read():
            try:
                from pe_parser import read_pe_menu
                import openpyxl

                items = read_pe_menu(filepath)

                wb = openpyxl.load_workbook(filepath)
                ws = wb.active
                image_rows = set()
                for img in ws._images:
                    if hasattr(img.anchor, '_from'):
                        image_rows.add(img.anchor._from.row)
                wb.close()

                self._pe_items = items
                self._pe_image_rows = image_rows

                self.root.after(0, lambda: self._pe_populate_tree(items, image_rows))
                self.root.after(0, lambda: self._pe_set_status(
                    f"OK - {len(items)} items, {len(image_rows)} images"
                ))
            except Exception as e:
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", t("msg_err", self.ui_lang).format(e)
                ))
                self.root.after(0, lambda: self._pe_set_status(""))

        threading.Thread(target=do_read, daemon=True).start()

    def _pe_populate_tree(self, items: list[dict], image_rows: set):
        """Fill the PE preview treeview."""
        for child in self.pe_tree.get_children():
            self.pe_tree.delete(child)

        for idx, item in enumerate(items):
            row_0based = item["row_idx"] - 1
            has_img = "Y" if row_0based in image_rows else ""
            status_display = (
                t("status_show", self.ui_lang) if item["status"] == "show"
                else t("status_hide", self.ui_lang)
            )
            self.pe_tree.insert("", "end", values=(
                "%04d" % (idx + 1),
                item["name"],
                item["price"],
                item["category"],
                item["menu_group"],
                item["tax_rate"],
                status_display,
                has_img,
            ))

    def _pe_convert(self):
        """Convert PE Menu to ZiiPOS template and extract images."""
        if not self._pe_items:
            filepath = self.pe_ent_file.get().strip()
            if not filepath:
                messagebox.showerror("Error", t("msg_no_file", self.ui_lang))
                return
            self._pe_read()
            messagebox.showinfo("Info",
                "Please wait for reading to finish, then click Convert again.\n"
                "请等待读取完成后再次点击转换。")
            return

        if not ensure_template(DEFAULT_TEMPLATE_FILE):
            return

        output_dir = self.pe_ent_out.get().strip() or DEFAULT_OUTPUT_DIR
        filepath = self.pe_ent_file.get().strip()
        items = self._pe_items

        self._pe_set_status(t("msg_exporting", self.ui_lang))

        def do_convert():
            try:
                os.makedirs(output_dir, exist_ok=True)

                # Build source DataFrame for V1 pipeline
                source_rows = []
                for idx, item in enumerate(items):
                    source_rows.append({
                        "ItemCode": "%04d" % (idx + 1),
                        "Description1": item["name"],
                        "Description2": "",
                        "Description3": "",
                        "Description4": "",
                        "Category": item["category"],
                        "MenuGroup": item["menu_group"],
                        "TaxRate": item["tax_rate"],
                        "Price": item["price"],
                        "Price2": 0,
                        "Price3": 0,
                        "Instruction": False,
                    })

                source = pd.DataFrame(source_rows)
                source = normalize_source_columns(source)

                tpl = pd.ExcelFile(DEFAULT_TEMPLATE_FILE)
                sheets = {name: pd.read_excel(tpl, name) for name in tpl.sheet_names}

                sheets["MenuGroupTable"], mg_code_map = processMenuGroup(
                    source, sheets["MenuGroupTable"],
                )
                sheets["Category"] = processCategory(
                    source, sheets["Category"], mg_code_map,
                )
                sheets["MenuItem"] = processItem(source, sheets["MenuItem"], "00")

                date_str = datetime.now().strftime("%Y%m%d%H%M%S")
                output_file = os.path.join(output_dir, f"export_FullMenu-{date_str}.xlsx")

                writer = pd.ExcelWriter(
                    output_file, engine="xlsxwriter",
                    engine_kwargs={"options": {
                        "strings_to_numbers": False,
                        "strings_to_formulas": False,
                    }},
                )
                text_fmt = writer.book.add_format({"num_format": "@"})

                overview = pd.DataFrame({
                    "ExportTime": [datetime.now().strftime("%Y-%m-%d %H:%M")],
                })
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
                print(f"[INFO] PE Export saved to: {output_file}")

                # Extract images
                self.root.after(0, lambda: self._pe_set_status(
                    t("msg_saving_images", self.ui_lang)
                ))
                img_count = 0
                if filepath:
                    from pe_parser import extract_images
                    img_count = extract_images(filepath, items, output_dir)

                pics_dir = os.path.join(output_dir, "pics")
                self.root.after(0, lambda: messagebox.showinfo(
                    "Done",
                    t("msg_pe_done", self.ui_lang).format(
                        output_file, img_count, pics_dir,
                    ),
                ))
                self.root.after(0, lambda: self._pe_set_status(
                    f"OK - {output_file} | {img_count} images"
                ))

            except Exception as e:
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", t("msg_err", self.ui_lang).format(e),
                ))
                self.root.after(0, lambda: self._pe_set_status(""))

        threading.Thread(target=do_convert, daemon=True).start()
