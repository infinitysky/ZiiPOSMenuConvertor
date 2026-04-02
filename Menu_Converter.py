import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory
import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime
from tqdm import tqdm
import wget
import ssl
import traceback

ssl._create_default_https_context = ssl._create_unverified_context

DEFAULT_OUTPUT_DIR = r'C:\Ziitech\Menu'
DEFAULT_TEMPLATE_DIR = r'C:\Ziitech'
DEFAULT_TEMPLATE_FILE = os.path.join(DEFAULT_TEMPLATE_DIR, 'ZiiPOS_MenuTemplate.xlsx')
TEMPLATE_DOWNLOAD_URL = "https://download.ziicloud.com/other/ZiiPOS_MenuTemplate.xlsx"

# ──────────────────────────────────────────────────────────────
# Chinese → English column mapping (simplified source files)
# ──────────────────────────────────────────────────────────────
CN_TO_EN_COLUMNS = {
    '产品代码': 'ItemCode',
    '名称1':   'Description1',
    '名称2':   'Description2',
    '名称3':   'Description3',
    '名称4':   'Description4',
    '分类':    'Category',
    '菜单组':  'MenuGroup',
    '税率':    'TaxRate',
    '价格1':   'Price',
    '价格2':   'Price2',
    '价格3':   'Price3',
    '价格4':   'Price4',
    '子名称1': 'SubDescription',
    '子名称2': 'SubDescription1',
    '子名称3': 'SubDescription2',
    '子名称4': 'SubDescription3',
    '产品统计组': 'ItemGroup',
    '指令':    'Instruction',
}


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
def _safe(val, default=''):
    if pd.isna(val):
        return default
    return val if str(val).strip() else default


def _num(val, default=0):
    try:
        v = float(val)
        return default if pd.isna(v) else v
    except (ValueError, TypeError):
        return default


def _bool(val, default=False):
    if pd.isna(val):
        return default
    return str(val).strip().lower() in ('true', '1', 'yes')


# ──────────────────────────────────────────────────────────────
# Source normalisation
# ──────────────────────────────────────────────────────────────
def normalize_source_columns(df):
    rename_map = {c: CN_TO_EN_COLUMNS[c] for c in df.columns if c in CN_TO_EN_COLUMNS}
    if rename_map:
        df = df.rename(columns=rename_map)
        print(f"[INFO] Chinese source detected – renamed {len(rename_map)} columns")

    if 'MenuGroup' not in df.columns:
        df['MenuGroup'] = 'Default'
    else:
        df['MenuGroup'] = df['MenuGroup'].fillna('Default')
        df.loc[df['MenuGroup'].astype(str).str.strip() == '', 'MenuGroup'] = 'Default'

    if 'ItemCode' in df.columns:
        norm_ctr, instr_ctr = 1, 1
        for idx in df.index:
            v = df.at[idx, 'ItemCode']
            if pd.isna(v) or str(v).strip() == '':
                is_instr = ('Instruction' in df.columns
                            and str(df.at[idx, 'Instruction']).strip().lower() == 'true')
                if is_instr:
                    df.at[idx, 'ItemCode'] = "I%03d" % instr_ctr
                    instr_ctr += 1
                else:
                    df.at[idx, 'ItemCode'] = "%04d" % norm_ctr
                    norm_ctr += 1

    return df


# ──────────────────────────────────────────────────────────────
# Process MenuGroupTable  (rebuilt from source)
# ──────────────────────────────────────────────────────────────
def processMenuGroup(source, template):
    base = template.iloc[0].copy()
    groups = source['MenuGroup'].drop_duplicates().reset_index(drop=True)
    rows = []
    mg_code_map = {}

    for i in tqdm(range(len(groups)), desc="MenuGroup"):
        row = base.copy()
        code = "%02d" % i
        row['Code']               = code
        row['Description']        = groups[i]
        row['CultureDescription'] = groups[i]
        row['OrderIndex']         = i
        rows.append(row)
        mg_code_map[groups[i]] = code

    return pd.DataFrame(rows), mg_code_map


# ──────────────────────────────────────────────────────────────
# Process Category  (rebuilt from source)
# ──────────────────────────────────────────────────────────────
def processCategory(source, template, mg_code_map):
    base = template.iloc[0].copy()
    cats = source['Category'].drop_duplicates().reset_index(drop=True)

    cat_to_mg = {}
    for _, r in source.iterrows():
        cat = r['Category']
        if cat not in cat_to_mg and pd.notna(r.get('MenuGroup')):
            cat_to_mg[cat] = r['MenuGroup']

    rows = []
    for i in tqdm(range(len(cats)), desc="Category"):
        row = base.copy()
        cat_name = cats[i]
        mg_name = cat_to_mg.get(cat_name, 'Default')
        mg_code = mg_code_map.get(mg_name, '00')

        row['Code']              = "%03d" % (i + 1)
        row['MenuGroupCode']     = mg_code
        row['Category']          = cat_name
        row['CultureCategory']   = cat_name
        row['Enable']            = True
        row['OrderIndex']        = i
        row['CategoryGroupSort'] = f"{mg_code},{i}"
        row['MenuGroupList']     = mg_code
        rows.append(row)

    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────
# Process MenuItem  (only this sheet is rebuilt from source)
# ──────────────────────────────────────────────────────────────
def processItem(source, template, menu_group_code="00"):
    base = template.iloc[0].copy()
    rows = []
    cat_pos = {}

    src_cols = set(source.columns)

    for i in tqdm(range(len(source)), desc="MenuItem"):
        src = source.iloc[i]
        row = base.copy()

        # ── ItemCode ──
        row['ItemCode'] = _safe(src.get('ItemCode'), "%04d" % (i + 1))

        # ── Descriptions ──
        desc1 = _safe(src.get('Description1'))
        row['Description1']        = desc1
        row['Description2']        = _safe(src.get('Description2'))
        row['Description3']        = _safe(src.get('Description3'))
        row['Description4']        = _safe(src.get('Description4'))
        row['CultureDescription']  = desc1

        # ── Category + sort ──
        cat = _safe(src.get('Category'))
        row['Category'] = cat
        if cat not in cat_pos:
            cat_pos[cat] = 0
        row['MenuItemCategorySort'] = f"{cat},{cat_pos[cat]}"
        cat_pos[cat] += 1

        # ── Prices ──
        main_price = _num(src.get('Price'))
        row['Price']  = main_price
        if 'Price1' in src_cols and pd.notna(src.get('Price1')):
            row['Price1'] = _num(src.get('Price1'), main_price)
        else:
            row['Price1'] = main_price
        row['Price2'] = _num(src.get('Price2'))
        row['Price3'] = _num(src.get('Price3'))
        if 'Price4' in src_cols:
            row['HappyHourPrice4'] = _num(src.get('Price4'))

        row['OnlinePrice1'] = 0
        row['OnlinePrice2'] = 0
        row['OnlinePrice3'] = 0
        row['OnlinePrice4'] = 0

        # ── SubDescriptions ──
        row['SubDescription']  = _safe(src.get('SubDescription'))
        row['SubDescription1'] = _safe(src.get('SubDescription1'))
        row['SubDescription2'] = _safe(src.get('SubDescription2'))
        row['SubDescription3'] = _safe(src.get('SubDescription3'))

        # ── Multiple logic ──
        sub_fields  = ['SubDescription', 'SubDescription1', 'SubDescription2', 'SubDescription3']
        extra_price = ['Price1', 'Price2', 'Price3', 'Price4']
        has_sub   = any(pd.notna(src.get(c)) and str(src.get(c, '')).strip()
                        for c in sub_fields if c in src_cols)
        has_multi = any(_num(src.get(c)) > 0
                        for c in extra_price if c in src_cols)
        row['Multiple'] = has_sub and has_multi

        # ── Tax / ItemGroup / Instruction ──
        row['TaxRate']     = _num(src.get('TaxRate'), _num(base.get('TaxRate'), 10))
        row['ItemGroup']   = _safe(src.get('ItemGroup')) or _safe(base.get('ItemGroup'), 'OTHERS')
        row['Instruction'] = _bool(src.get('Instruction'))

        # ── Pass-through optional fields ──
        for col in ['Scalable', 'OpenPrice', 'OnlineStatus', 'QRCodeStatus']:
            if col in src_cols and pd.notna(src.get(col)):
                row[col] = _bool(src.get(col), row[col])
        for col in ['PrinterPort1', 'PrinterPort2', 'PrinterPort3', 'PrinterPort4',
                     'HappyHourPrice1', 'HappyHourPrice2', 'HappyHourPrice3']:
            if col in src_cols and pd.notna(src.get(col)):
                row[col] = _num(src.get(col))

        row['OrderIndex'] = i
        rows.append(row)

    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────
# Main conversion orchestrator
# ──────────────────────────────────────────────────────────────
def processMenu(source_file, template_file, output_dir,
                shop_name="", menu_group_code="00"):

    os.makedirs(output_dir, exist_ok=True)

    # ── Read source ──
    source = pd.read_excel(source_file, index_col=None)
    source = normalize_source_columns(source)
    print("[INFO] Columns after normalisation:", list(source.columns))

    # ── Read every sheet from template ──
    tpl = pd.ExcelFile(template_file)
    sheets = {name: pd.read_excel(tpl, name) for name in tpl.sheet_names}

    # ── Rebuild MenuGroupTable, Category & MenuItem ──
    sheets['MenuGroupTable'], mg_code_map = processMenuGroup(
        source, sheets['MenuGroupTable'])
    sheets['Category'] = processCategory(source, sheets['Category'], mg_code_map)
    sheets['MenuItem'] = processItem(source, sheets['MenuItem'], menu_group_code)

    # ── Build output filename ──
    date_str = datetime.now().strftime('%Y%m%d%H%M%S')
    prefix = f"-{shop_name}" if shop_name.strip() else ""
    output_file = os.path.join(output_dir, f'export_FullMenu{prefix}-{date_str}.xlsx')

    # ── Write output ──
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter',
                            engine_kwargs={'options': {
                                'strings_to_numbers': False,
                                'strings_to_formulas': False,
                            }})
    text_fmt = writer.book.add_format({'num_format': '@'})

    overview = pd.DataFrame({'ExportTime': [datetime.now().strftime('%Y-%m-%d %H:%M')]})
    overview.to_excel(writer, sheet_name='Overview', index=False)

    for name in tpl.sheet_names:
        if name == 'Overview':
            continue
        df = sheets[name]

        if name == 'MenuGroupTable':
            df['Code'] = df['Code'].apply(lambda v: "%02d" % int(v) if pd.notna(v) else "00")
        if name == 'Course':
            df['CourseCode'] = ["%02d" % (i + 1) for i in range(len(df))]
        if name == 'Category':
            df['Code'] = df['Code'].apply(lambda v: "%03d" % int(v) if pd.notna(v) else "001")
            df['MenuGroupCode'] = df['MenuGroupCode'].apply(
                lambda v: str(v).strip() if pd.notna(v) else menu_group_code)
            df['MenuGroupList'] = df['MenuGroupList'].apply(
                lambda v: str(v).strip() if pd.notna(v) else menu_group_code)

        df.to_excel(writer, sheet_name=name, index=False)

        if name == 'MenuGroupTable':
            writer.sheets[name].set_column('A:A', None, text_fmt)
        if name == 'Course':
            writer.sheets[name].set_column('A:A', None, text_fmt)
        if name == 'Category':
            ws = writer.sheets[name]
            ws.set_column('A:A', None, text_fmt)
            ws.set_column('B:B', None, text_fmt)

    writer.close()
    print(f"[INFO] Output saved to: {output_file}")
    return output_file


def ensure_template(template_file):
    """Auto-download template if missing. Returns True if ready, False otherwise."""
    if os.path.exists(template_file):
        return True

    tpl_dir = os.path.dirname(template_file)
    os.makedirs(tpl_dir, exist_ok=True)

    try:
        print(f"[INFO] Template not found, downloading from {TEMPLATE_DOWNLOAD_URL}")
        wget.download(TEMPLATE_DOWNLOAD_URL, template_file)
        print()
        if os.path.exists(template_file):
            return True
    except Exception:
        pass

    messagebox.showerror(
        "Template Missing / 模板缺失",
        f"Template file not found and download failed.\n"
        f"模板文件不存在且下载失败。\n\n"
        f"Please check your network connection and try again.\n"
        f"请检查网络连接后重试。\n\n"
        f"Or manually place the template at:\n"
        f"或手动将模板放至：\n"
        f"{template_file}"
    )
    return False


def infoProcess(source_file, template_file, output_dir,
                shop_name="", menu_group_code="00"):
    if not source_file:
        messagebox.showerror("Error", "Please select a menu file!\n请选择菜单文件！")
        return
    if not template_file:
        messagebox.showerror("Error", "Please specify a template file!\n请指定模板文件！")
        return
    if not ensure_template(template_file):
        return
    try:
        out = processMenu(source_file, template_file, output_dir,
                          shop_name, menu_group_code)
        messagebox.showinfo("Done", f"Export completed!\n导出完成！\n\n{out}")
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Error", f"Processing failed / 处理失败:\n{e}")


# ──────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        root.title("ZiiPOS Menu Converter V2.0")
        w, h = 640, 240
        sx = root.winfo_screenwidth()
        sy = root.winfo_screenheight()
        root.geometry(f'{w}x{h}+{(sx-w)//2}+{(sy-h)//2}')
        root.resizable(False, False)

        ft = tkFont.Font(family='Times', size=10)
        LBL_X, LBL_W = 20, 120
        ENT_X, ENT_W = 150, 320
        BTN_X, BTN_W = 480, 80
        ROW_H, Y_GAP = 30, 38
        y = 20

        def _row_label(text, yy):
            tk.Label(root, text=text, font=ft, fg="#333",
                     anchor="w").place(x=LBL_X, y=yy, width=LBL_W, height=ROW_H)

        def _row_entry(yy, default="", entry_w=ENT_W):
            e = tk.Entry(root, borderwidth="1px", font=ft, fg="#333")
            if default:
                e.insert(0, default)
            e.place(x=ENT_X, y=yy, width=entry_w, height=ROW_H)
            return e

        def _row_btn(text, yy, cmd):
            tk.Button(root, text=text, font=ft,
                      command=cmd).place(x=BTN_X, y=yy, width=BTN_W, height=ROW_H)

        # Row 0 – Menu File
        _row_label("Menu File", y)
        self.ent_menu = _row_entry(y)
        _row_btn("Select", y, self._sel_menu)

        # Row 1 – Output Path
        y += Y_GAP
        _row_label("Output Path", y)
        self.ent_out = _row_entry(y, DEFAULT_OUTPUT_DIR)
        _row_btn("Browse", y, self._sel_out)

        # Action buttons
        y_btn = h - 60
        tk.Button(root, text="Convert", font=ft, bg="#4CAF50", fg="white",
                  width=12, command=self._convert
                  ).place(x=w//2 - 140, y=y_btn, width=120, height=42)
        tk.Button(root, text="Close", font=ft,
                  width=12, command=sys.exit
                  ).place(x=w//2 + 20, y=y_btn, width=120, height=42)

    def _sel_menu(self):
        f = askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if f:
            self.ent_menu.delete(0, tk.END)
            self.ent_menu.insert(0, f)

    def _sel_out(self):
        d = askdirectory(initialdir=self.ent_out.get())
        if d:
            self.ent_out.delete(0, tk.END)
            self.ent_out.insert(0, d)

    def _convert(self):
        infoProcess(
            self.ent_menu.get().strip(),
            DEFAULT_TEMPLATE_FILE,
            self.ent_out.get().strip() or DEFAULT_OUTPUT_DIR,
            "",
            "00",
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
