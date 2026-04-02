"""
Flask application -- all API routes for Menu Converter V3.
"""
import os
import sys
import json
import tempfile
import traceback
import io
import shutil
from datetime import datetime

from flask import Flask, request, jsonify, render_template, send_file
from PIL import Image
from werkzeug.utils import secure_filename

LIB_DIR = os.path.join(os.path.dirname(__file__), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from i18n import STRINGS, LANG_CODES, t
from ai_reader import (
    read_file, read_file_vision, read_image_structured,
    refine_ocr_with_llm, detect_language, VISION_PROVIDERS,
)
from menu_parser import parse_menu_text, parse_menu_from_boxes
from translator import Translator, API_PROVIDERS
from Menu_Converter import (
    processMenu, ensure_template,
    DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_FILE,
)

CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                          "ZiiPOSMenuConverter")
CONFIG_FILE = os.path.join(CONFIG_DIR, "settings_v3.json")
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "mcv3_uploads")

THUMBNAIL_SIZE = (180, 180)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}
PDF_EXTS = {".pdf"}
DOC_EXTS = {".docx", ".doc"}
EXCEL_EXTS = {".xlsx", ".xls"}


def _load_config() -> dict:
    os.makedirs(CONFIG_DIR, exist_ok=True)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_config(cfg: dict):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def create_app():
    app = Flask(__name__,
                template_folder=os.path.join(os.path.dirname(__file__), "templates"),
                static_folder=os.path.join(os.path.dirname(__file__), "static"))

    if os.path.exists(UPLOAD_DIR):
        shutil.rmtree(UPLOAD_DIR, ignore_errors=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)

    app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

    @app.after_request
    def add_no_cache(resp):
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp

    # ── Page ──────────────────────────────────────────────────
    @app.route("/")
    def index():
        import time
        return render_template("index.html", cache_bust=int(time.time()))

    # ── Config ────────────────────────────────────────────────
    @app.route("/api/config/load", methods=["POST"])
    def config_load():
        cfg = _load_config()
        return jsonify({
            "ui_lang":      cfg.get("ui_lang", "en"),
            "api_provider":  cfg.get("api_provider", "OpenAI"),
            "api_keys":     cfg.get("api_keys", {}),
            "output_path":  cfg.get("output_path", DEFAULT_OUTPUT_DIR),
            "desc_langs":   cfg.get("desc_langs", ["jp", "cn", "en", ""]),
            "src_lang":     cfg.get("src_lang", "auto"),
            "trans_mode":   cfg.get("trans_mode", "offline"),
        })

    @app.route("/api/config/save", methods=["POST"])
    def config_save():
        data = request.get_json(force=True)
        cfg = _load_config()
        for k, v in data.items():
            cfg[k] = v
        _save_config(cfg)
        return jsonify({"ok": True})

    # ── i18n ──────────────────────────────────────────────────
    @app.route("/api/i18n/<lang>")
    def get_i18n(lang):
        strings = STRINGS.get(lang, STRINGS["en"])
        return jsonify(strings)

    # ── File Dialog (pywebview native) ────────────────────────
    @app.route("/api/file-dialog", methods=["POST"])
    def file_dialog():
        """
        Open native file/folder dialog via pywebview.
        Expects: {type: "open_files"|"open_dir", filetypes: [...]}
        """
        data = request.get_json(force=True)
        dialog_type = data.get("type", "open_files")

        try:
            import webview
            windows = webview.windows
            if not windows:
                return jsonify({"paths": [], "error": "No window available"})

            win = windows[0]
            if dialog_type == "open_dir":
                result = win.create_file_dialog(
                    webview.FOLDER_DIALOG,
                    allow_multiple=False,
                )
            else:
                file_types = data.get("filetypes", [
                    "Image files (*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.webp)",
                    "PDF files (*.pdf)",
                    "Word files (*.docx;*.doc)",
                    "Excel files (*.xlsx;*.xls)",
                    "All files (*.*)",
                ])
                result = win.create_file_dialog(
                    webview.OPEN_DIALOG,
                    allow_multiple=True,
                    file_types=tuple(file_types),
                )

            paths = list(result) if result else []
            return jsonify({"paths": paths})
        except Exception as e:
            traceback.print_exc()
            return jsonify({"paths": [], "error": str(e)})

    # ── Upload temp (for drag-and-drop) ─────────────────────
    @app.route("/api/upload-temp", methods=["POST"])
    def upload_temp():
        """
        Accept file uploads from drag-and-drop (where full paths aren't available).
        Saves to temp dir, returns full paths.
        """
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        saved = []
        for key in request.files:
            f = request.files[key]
            if not f.filename:
                continue
            fname = secure_filename(f.filename)
            if not fname:
                fname = f"upload_{len(saved)}"
            dest = os.path.join(UPLOAD_DIR, fname)
            counter = 1
            base, fext = os.path.splitext(dest)
            while os.path.exists(dest):
                dest = f"{base}_{counter}{fext}"
                counter += 1
            f.save(dest)
            saved.append({"path": dest, "name": os.path.basename(dest)})
        return jsonify({"files": saved})

    # ── Thumbnail ─────────────────────────────────────────────
    @app.route("/api/thumbnail")
    def thumbnail():
        """Generate thumbnail for preview. ?path=<filepath>"""
        fpath = request.args.get("path", "")
        if not fpath or not os.path.isfile(fpath):
            return "Not found", 404

        ext = os.path.splitext(fpath)[1].lower()

        if ext in IMAGE_EXTS:
            try:
                img = Image.open(fpath)
                img.thumbnail(THUMBNAIL_SIZE)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=80)
                buf.seek(0)
                return send_file(buf, mimetype="image/jpeg")
            except Exception:
                return "Cannot generate thumbnail", 500

        if ext in PDF_EXTS:
            try:
                import fitz
                doc = fitz.open(fpath)
                page = doc[0]
                pix = page.get_pixmap(dpi=72)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                img.thumbnail(THUMBNAIL_SIZE)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=80)
                buf.seek(0)
                doc.close()
                return send_file(buf, mimetype="image/jpeg")
            except Exception:
                return "Cannot generate PDF thumbnail", 500

        return "No preview", 204

    # ── Read & Parse (AI import) ──────────────────────────────
    @app.route("/api/read", methods=["POST"])
    def api_read():
        """
        Read files and parse menu items.
        {files: [path,...], mode: "online"|"offline",
         src_lang: "auto"|"jp"|..., api_key: "...", provider: "OpenAI"|...}
        """
        data = request.get_json(force=True)
        files = data.get("files", [])
        mode = data.get("mode", "offline")
        src_lang = data.get("src_lang", "auto")
        api_key = data.get("api_key", "")
        provider = data.get("provider", "OpenAI")

        if not files:
            return jsonify({"error": "No files provided"}), 400

        all_items = []
        detected_lang = "en"
        item_counter = 1

        for fpath in files:
            if not os.path.isfile(fpath):
                continue

            ext = os.path.splitext(fpath)[1].lower()
            use_vision = (
                mode == "online"
                and api_key
                and ext in (IMAGE_EXTS | PDF_EXTS)
            )

            try:
                if use_vision:
                    vision_items = read_file_vision(fpath, api_key, provider, src_lang)
                    if vision_items:
                        for vi in vision_items:
                            vi["itemcode"] = "%04d" % item_counter
                            vi.setdefault("name", "")
                            vi.setdefault("price", 0)
                            vi.setdefault("category", "Default")
                            all_items.append({
                                "itemcode": vi["itemcode"],
                                "name": vi.get("name", ""),
                                "name_alt": vi.get("name_alt", ""),
                                "price": float(vi.get("price", 0)),
                                "category": vi.get("category", "Default"),
                            })
                            item_counter += 1
                        continue

                lang_hint = src_lang if src_lang != "auto" else "jp"
                is_image = ext in IMAGE_EXTS

                if is_image:
                    structured = read_image_structured(fpath, lang_hint)

                    if structured["mode"] == "table" and structured["items"]:
                        for item in structured["items"]:
                            item["itemcode"] = "%04d" % item_counter
                            all_items.append(item)
                            item_counter += 1
                        continue

                    parsed_items = []

                    if structured["boxes"]:
                        if mode == "online" and api_key:
                            llm_items = refine_ocr_with_llm(
                                structured["boxes"], api_key, provider)
                            if llm_items:
                                for li in llm_items:
                                    parsed_items.append({
                                        "name": li.get("name", ""),
                                        "name_alt": li.get("name_alt", ""),
                                        "price": float(li.get("price", 0)),
                                        "category": li.get("category", "Default"),
                                    })

                        if not parsed_items:
                            parsed_items = parse_menu_from_boxes(
                                structured["boxes"])

                    if not parsed_items and structured["text"]:
                        det = detect_language(structured["text"])
                        if det != "en":
                            detected_lang = det
                        parsed_items = parse_menu_text(
                            structured["text"], det)

                    for item in parsed_items:
                        item["itemcode"] = "%04d" % item_counter
                        all_items.append(item)
                        item_counter += 1
                    continue

                raw_text = read_file(fpath, lang_hint)
                if raw_text:
                    det = detect_language(raw_text)
                    if det != "en":
                        detected_lang = det
                    parsed = parse_menu_text(raw_text, det)
                    for item in parsed:
                        item["itemcode"] = "%04d" % item_counter
                        all_items.append(item)
                        item_counter += 1

            except Exception as e:
                traceback.print_exc()
                return jsonify({"error": f"Error reading {os.path.basename(fpath)}: {e}"}), 500

        return jsonify({
            "items": all_items,
            "detected_lang": detected_lang,
            "count": len(all_items),
        })

    # ── Translate ─────────────────────────────────────────────
    @app.route("/api/translate", methods=["POST"])
    def api_translate():
        """
        Translate item names to fill description columns.
        {items, desc_langs: ["jp","cn","en",""], src_lang, mode, api_key, provider}
        """
        data = request.get_json(force=True)
        items = data.get("items", [])
        desc_langs = data.get("desc_langs", ["jp", "cn", "en", ""])
        src_lang = data.get("src_lang", "auto")
        mode = data.get("mode", "online")
        api_key = data.get("api_key", "")
        provider = data.get("provider", "OpenAI")

        if not items:
            return jsonify({"error": "No items to translate"}), 400

        trans = Translator(mode=mode, api_key=api_key, provider=provider)

        if src_lang == "auto":
            all_names = " ".join(it.get("name", "") for it in items)
            src_lang = detect_language(all_names)

        for item in items:
            name = item.get("name", "")
            name_alt = item.get("name_alt", "")
            descs = []
            for lang in desc_langs:
                if not lang:
                    descs.append("")
                elif lang == src_lang:
                    descs.append(name)
                elif name_alt and lang != src_lang:
                    descs.append(name_alt)
                else:
                    try:
                        descs.append(trans.translate(name, src_lang, lang))
                    except Exception:
                        descs.append(name)
            item["descriptions"] = descs

        return jsonify({"items": items, "src_lang": src_lang})

    # ── Export Simple Excel ───────────────────────────────────
    @app.route("/api/export/simple", methods=["POST"])
    def api_export_simple():
        """
        Export items to simplified menuCollection.xlsx format.
        {items, desc_langs, output_dir}
        """
        import pandas as pd

        data = request.get_json(force=True)
        items = data.get("items", [])
        desc_langs = data.get("desc_langs", ["jp", "cn", "en", ""])
        output_dir = data.get("output_dir", DEFAULT_OUTPUT_DIR)

        os.makedirs(output_dir, exist_ok=True)

        rows = []
        for item in items:
            descs = item.get("descriptions", [item.get("name", "")] * 4)
            while len(descs) < 4:
                descs.append("")
            rows.append({
                "ItemCode":       item.get("itemcode", ""),
                "Description1":   descs[0],
                "Description2":   descs[1],
                "Description3":   descs[2],
                "Description4":   descs[3],
                "Category":       item.get("category", "Default"),
                "MenuGroup":      "Default",
                "TaxRate":        10,
                "Price":          item.get("price", 0),
                "Price2":         0,
                "Price3":         0,
                "Price4":         0,
                "SubDescription":  "",
                "SubDescription1": "",
                "SubDescription2": "",
                "SubDescription3": "",
                "ItemGroup":      "OTHERS",
                "Instruction":    False,
            })

        df = pd.DataFrame(rows)
        ts = datetime.now().strftime("%Y%m%d%H%M%S")
        outfile = os.path.join(output_dir, f"menuCollection-{ts}.xlsx")
        df.to_excel(outfile, index=False)
        return jsonify({"output_file": outfile})

    # ── Export Full Menu ──────────────────────────────────────
    @app.route("/api/export/full", methods=["POST"])
    def api_export_full():
        """
        Export items to full ZiiPOS_MenuTemplate format.
        {items, desc_langs, output_dir}
        """
        import pandas as pd

        data = request.get_json(force=True)
        items = data.get("items", [])
        desc_langs = data.get("desc_langs", ["jp", "cn", "en", ""])
        output_dir = data.get("output_dir", DEFAULT_OUTPUT_DIR)

        os.makedirs(output_dir, exist_ok=True)

        rows = []
        for item in items:
            descs = item.get("descriptions", [item.get("name", "")] * 4)
            while len(descs) < 4:
                descs.append("")
            rows.append({
                "ItemCode":       item.get("itemcode", ""),
                "Description1":   descs[0],
                "Description2":   descs[1],
                "Description3":   descs[2],
                "Description4":   descs[3],
                "Category":       item.get("category", "Default"),
                "MenuGroup":      "Default",
                "TaxRate":        10,
                "Price":          item.get("price", 0),
                "Price2":         0,
                "Price3":         0,
                "Price4":         0,
                "SubDescription":  "",
                "SubDescription1": "",
                "SubDescription2": "",
                "SubDescription3": "",
                "ItemGroup":      "OTHERS",
                "Instruction":    False,
            })

        temp_source = os.path.join(output_dir, "_temp_ai_source.xlsx")
        pd.DataFrame(rows).to_excel(temp_source, index=False)

        try:
            ensure_template(DEFAULT_TEMPLATE_FILE)
            outfile = processMenu(temp_source, DEFAULT_TEMPLATE_FILE, output_dir)
            return jsonify({"output_file": outfile})
        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500
        finally:
            if os.path.exists(temp_source):
                os.remove(temp_source)

    # ── Excel Import (V1 direct) ──────────────────────────────
    @app.route("/api/excel-import", methods=["POST"])
    def api_excel_import():
        """
        Direct Excel import using V1 logic.
        {source_file, output_dir}
        """
        data = request.get_json(force=True)
        source_file = data.get("source_file", "")
        output_dir = data.get("output_dir", DEFAULT_OUTPUT_DIR)

        if not source_file or not os.path.isfile(source_file):
            return jsonify({"error": "Source file not found"}), 400

        try:
            ensure_template(DEFAULT_TEMPLATE_FILE)
            outfile = processMenu(source_file, DEFAULT_TEMPLATE_FILE, output_dir)
            return jsonify({"output_file": outfile})
        except Exception as e:
            traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    # ── Providers list ────────────────────────────────────────
    @app.route("/api/providers")
    def api_providers():
        return jsonify({
            "vision": list(VISION_PROVIDERS.keys()),
            "translation": list(API_PROVIDERS.keys()),
        })

    return app
