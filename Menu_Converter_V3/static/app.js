/* ═══════════════════════════════════════════════════════════
   ZiiPOS Menu Converter V3 -- Frontend Logic (vanilla JS)
   ═══════════════════════════════════════════════════════════ */

(function () {
  "use strict";

  /* ── State ──────────────────────────────────────────────── */
  let uiLang = "en";
  let strings = {};
  let selectedFiles = [];   // [{path, name, ext}]
  let menuItems = [];       // [{itemcode, name, price, category, name_alt, descriptions}]

  const IMAGE_EXTS = new Set([".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]);
  const PDF_EXTS = new Set([".pdf"]);

  /* ── DOM references ─────────────────────────────────────── */
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const elTitle       = $("#app-title");
  const elUiLang      = $("#ui-lang");
  const elFileGrid    = $("#file-grid");
  const elDropZone    = $("#drop-zone");
  const elFileCount   = $("#file-count");
  const elMenuTbody   = $("#menu-tbody");
  const elItemCount   = $("#item-count");
  const elStatus      = $("#status-text");
  const elProvider    = $("#ai-provider");
  const elApiKey      = $("#ai-api-key");
  const elTransMode   = $("#ai-trans-mode");
  const elSrcLang     = $("#ai-src-lang");
  const elAiOutput    = $("#ai-output-path");
  const elExcelSource = $("#excel-source");
  const elExcelOutput = $("#excel-output-path");

  /* ── Helpers ────────────────────────────────────────────── */
  function api(url, body) {
    return fetch(url, {
      method: body !== undefined ? "POST" : "GET",
      headers: { "Content-Type": "application/json" },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    }).then((r) => r.json());
  }

  function status(msg) { elStatus.textContent = msg; }

  function ext(filename) {
    const i = filename.lastIndexOf(".");
    return i >= 0 ? filename.substring(i).toLowerCase() : "";
  }

  function basename(path) {
    return path.replace(/\\/g, "/").split("/").pop();
  }

  /* ── Tabs ───────────────────────────────────────────────── */
  $$(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      $$(".tab-btn").forEach((b) => b.classList.remove("active"));
      $$(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $(`#${btn.dataset.tab}`).classList.add("active");
    });
  });

  /* ── i18n ───────────────────────────────────────────────── */
  async function loadStrings(lang) {
    strings = await api(`/api/i18n/${lang}`);
    uiLang = lang;
    applyStrings();
  }

  function applyStrings() {
    elTitle.textContent = strings.title || "ZiiPOS Menu Converter V3";
    $$("#tabBtn-ai").forEach((el) => (el.textContent = strings.tab_ai || "AI Import"));
    $$("#tabBtn-excel").forEach((el) => (el.textContent = strings.tab_excel || "Excel Import"));

    $$("[data-i18n]").forEach((el) => {
      const key = el.getAttribute("data-i18n");
      if (strings[key]) {
        if (el.tagName === "INPUT" && el.type !== "text") return;
        if (el.tagName === "OPTION") el.textContent = strings[key];
        else if (el.tagName === "INPUT") el.placeholder = strings[key];
        else el.textContent = strings[key];
      }
    });
  }

  elUiLang.addEventListener("change", async () => {
    await loadStrings(elUiLang.value);
    api("/api/config/save", { ui_lang: elUiLang.value });
  });

  /* ── Config ─────────────────────────────────────────────── */
  async function loadConfig() {
    const cfg = await api("/api/config/load", {});
    uiLang = cfg.ui_lang || "en";
    elUiLang.value = uiLang;
    elProvider.value = cfg.api_provider || "OpenAI";
    elTransMode.value = cfg.trans_mode || "offline";
    elSrcLang.value = cfg.src_lang || "auto";
    elAiOutput.value = cfg.output_path || "C:\\Ziitech\\Menu";
    elExcelOutput.value = cfg.output_path || "C:\\Ziitech\\Menu";

    const keys = cfg.api_keys || {};
    elApiKey.value = keys[elProvider.value] || "";

    const dl = cfg.desc_langs || ["jp", "cn", "en", ""];
    for (let i = 0; i < 4; i++) {
      const sel = $(`#desc-lang-${i + 1}`);
      if (sel && dl[i] !== undefined) sel.value = dl[i];
    }

    await loadStrings(uiLang);
  }

  elProvider.addEventListener("change", async () => {
    const cfg = await api("/api/config/load", {});
    const keys = cfg.api_keys || {};
    elApiKey.value = keys[elProvider.value] || "";
    api("/api/config/save", { api_provider: elProvider.value });
  });

  $("#btn-save-key").addEventListener("click", async () => {
    const cfg = await api("/api/config/load", {});
    const keys = cfg.api_keys || {};
    keys[elProvider.value] = elApiKey.value;
    await api("/api/config/save", { api_keys: keys });
    status(`API key saved for ${elProvider.value}`);
  });

  /* ── File Grid ──────────────────────────────────────────── */
  function renderFileGrid() {
    elFileGrid.innerHTML = "";

    if (selectedFiles.length === 0) {
      const dz = document.createElement("div");
      dz.className = "drop-zone";
      dz.id = "drop-zone";
      dz.innerHTML = `<div class="drop-icon">&#128206;</div><p data-i18n="msg_drop_hint">${strings.msg_drop_hint || "Drag & drop files here or click Add Files"}</p>`;
      elFileGrid.appendChild(dz);
    } else {
      selectedFiles.forEach((f, idx) => {
        const card = document.createElement("div");
        card.className = "file-card";

        const thumbEl = document.createElement("div");
        thumbEl.className = "thumb";

        const fileExt = ext(f.name);
        if (IMAGE_EXTS.has(fileExt) || PDF_EXTS.has(fileExt)) {
          const img = document.createElement("img");
          img.src = `/api/thumbnail?path=${encodeURIComponent(f.path)}`;
          img.alt = f.name;
          img.loading = "lazy";
          const fallback = PDF_EXTS.has(fileExt) ? "📄" : "🖼";
          img.onerror = () => { thumbEl.textContent = fallback; img.remove(); };
          thumbEl.appendChild(img);
        } else if (fileExt === ".docx" || fileExt === ".doc") {
          thumbEl.textContent = "📝";
        } else if (fileExt === ".xlsx" || fileExt === ".xls") {
          thumbEl.textContent = "📊";
        } else {
          thumbEl.textContent = "📁";
        }

        const fname = document.createElement("div");
        fname.className = "fname";
        fname.textContent = f.name;
        fname.title = f.path;

        const removeBtn = document.createElement("button");
        removeBtn.className = "remove-btn";
        removeBtn.textContent = "×";
        removeBtn.addEventListener("click", () => {
          selectedFiles.splice(idx, 1);
          renderFileGrid();
        });

        card.appendChild(thumbEl);
        card.appendChild(fname);
        card.appendChild(removeBtn);
        elFileGrid.appendChild(card);
      });
    }

    elFileCount.textContent = `${selectedFiles.length} files`;
  }

  async function handleDroppedFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    status("Uploading dropped files...");

    const formData = new FormData();
    for (let i = 0; i < fileList.length; i++) {
      formData.append(`file_${i}`, fileList[i]);
    }

    try {
      const resp = await fetch("/api/upload-temp", { method: "POST", body: formData });
      const data = await resp.json();
      if (data.files && data.files.length) {
        for (const f of data.files) {
          if (!selectedFiles.find((sf) => sf.path === f.path)) {
            selectedFiles.push({ path: f.path, name: f.name, ext: ext(f.name) });
          }
        }
        renderFileGrid();
        status(`Added ${data.files.length} file(s)`);
      }
    } catch (e) {
      status(`Upload error: ${e.message}`);
    }
  }

  // Drag-and-drop on the entire file grid area
  elFileGrid.addEventListener("dragover", (e) => {
    e.preventDefault();
    elFileGrid.classList.add("drag-over");
    const dz = $("#drop-zone");
    if (dz) dz.classList.add("drag-over");
  });
  elFileGrid.addEventListener("dragleave", (e) => {
    if (elFileGrid.contains(e.relatedTarget)) return;
    elFileGrid.classList.remove("drag-over");
    const dz = $("#drop-zone");
    if (dz) dz.classList.remove("drag-over");
  });
  elFileGrid.addEventListener("drop", (e) => {
    e.preventDefault();
    elFileGrid.classList.remove("drag-over");
    const dz = $("#drop-zone");
    if (dz) dz.classList.remove("drag-over");
    if (e.dataTransfer.files && e.dataTransfer.files.length) {
      handleDroppedFiles(e.dataTransfer.files);
    }
  });

  $("#btn-add-files").addEventListener("click", async () => {
    const result = await api("/api/file-dialog", {
      type: "open_files",
      filetypes: [
        "All supported (*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.webp;*.pdf;*.docx;*.doc;*.xlsx;*.xls)",
        "Images (*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.webp)",
        "PDF (*.pdf)",
        "Word (*.docx;*.doc)",
        "Excel (*.xlsx;*.xls)",
      ],
    });
    if (result.paths && result.paths.length) {
      for (const p of result.paths) {
        if (!selectedFiles.find((f) => f.path === p)) {
          selectedFiles.push({ path: p, name: basename(p), ext: ext(p) });
        }
      }
      renderFileGrid();
    }
  });

  $("#btn-clear-files").addEventListener("click", () => {
    selectedFiles = [];
    renderFileGrid();
  });

  /* ── Editable Table ─────────────────────────────────────── */
  function renderTable() {
    elMenuTbody.innerHTML = "";
    menuItems.forEach((item, idx) => {
      const tr = document.createElement("tr");
      tr.dataset.idx = idx;

      const tdCheck = document.createElement("td");
      tdCheck.className = "col-check";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.addEventListener("change", () => tr.classList.toggle("selected", cb.checked));
      tdCheck.appendChild(cb);

      const tdCode = document.createElement("td");
      tdCode.contentEditable = "true";
      tdCode.textContent = item.itemcode;
      tdCode.addEventListener("blur", () => { menuItems[idx].itemcode = tdCode.textContent.trim(); });

      const tdName = document.createElement("td");
      tdName.contentEditable = "true";
      tdName.className = "col-name";
      tdName.textContent = item.name;
      tdName.addEventListener("blur", () => { menuItems[idx].name = tdName.textContent.trim(); });

      const tdPrice = document.createElement("td");
      tdPrice.contentEditable = "true";
      tdPrice.textContent = item.price;
      tdPrice.addEventListener("blur", () => {
        const v = parseFloat(tdPrice.textContent);
        menuItems[idx].price = isNaN(v) ? 0 : v;
        tdPrice.textContent = menuItems[idx].price;
      });

      const tdCat = document.createElement("td");
      tdCat.contentEditable = "true";
      tdCat.textContent = item.category;
      tdCat.addEventListener("blur", () => { menuItems[idx].category = tdCat.textContent.trim(); });

      tr.appendChild(tdCheck);
      tr.appendChild(tdCode);
      tr.appendChild(tdName);
      tr.appendChild(tdPrice);
      tr.appendChild(tdCat);
      elMenuTbody.appendChild(tr);
    });

    elItemCount.textContent = `${menuItems.length} items`;
  }

  // Tab navigation between editable cells
  document.getElementById("menu-table").addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    const td = e.target.closest("td[contenteditable]");
    if (!td) return;

    e.preventDefault();
    const tr = td.parentElement;
    const cells = Array.from(tr.querySelectorAll("td[contenteditable]"));
    const ci = cells.indexOf(td);

    if (!e.shiftKey) {
      if (ci < cells.length - 1) {
        cells[ci + 1].focus();
      } else {
        const nextRow = tr.nextElementSibling;
        if (nextRow) {
          const nextCells = nextRow.querySelectorAll("td[contenteditable]");
          if (nextCells.length) nextCells[0].focus();
        }
      }
    } else {
      if (ci > 0) {
        cells[ci - 1].focus();
      } else {
        const prevRow = tr.previousElementSibling;
        if (prevRow) {
          const prevCells = prevRow.querySelectorAll("td[contenteditable]");
          if (prevCells.length) prevCells[prevCells.length - 1].focus();
        }
      }
    }
  });

  $("#btn-add-row").addEventListener("click", () => {
    const code = String(menuItems.length + 1).padStart(4, "0");
    menuItems.push({ itemcode: code, name: "", price: 0, category: "Default" });
    renderTable();
    const lastRow = elMenuTbody.lastElementChild;
    if (lastRow) {
      const nameTd = lastRow.querySelectorAll("td[contenteditable]")[1];
      if (nameTd) nameTd.focus();
    }
  });

  $("#btn-del-row").addEventListener("click", () => {
    const checkboxes = elMenuTbody.querySelectorAll("input[type=checkbox]:checked");
    if (checkboxes.length === 0) return;
    const indices = new Set();
    checkboxes.forEach((cb) => {
      const tr = cb.closest("tr");
      if (tr) indices.add(parseInt(tr.dataset.idx));
    });
    menuItems = menuItems.filter((_, i) => !indices.has(i));
    menuItems.forEach((item, i) => { item.itemcode = String(i + 1).padStart(4, "0"); });
    renderTable();
  });

  $("#check-all").addEventListener("change", (e) => {
    const checked = e.target.checked;
    elMenuTbody.querySelectorAll("input[type=checkbox]").forEach((cb) => {
      cb.checked = checked;
      cb.closest("tr").classList.toggle("selected", checked);
    });
  });

  /* ── Read & Parse ───────────────────────────────────────── */
  $("#btn-read").addEventListener("click", async () => {
    if (selectedFiles.length === 0) {
      status(strings.msg_no_file || "Please select a file!");
      return;
    }

    status(strings.msg_reading || "Reading file...");
    setBusy(true);

    try {
      const result = await api("/api/read", {
        files: selectedFiles.map((f) => f.path),
        mode: elTransMode.value,
        src_lang: elSrcLang.value,
        api_key: elApiKey.value,
        provider: elProvider.value,
      });

      if (result.error) {
        status(`Error: ${result.error}`);
        return;
      }

      menuItems = result.items || [];
      renderTable();
      status(`Parsed ${menuItems.length} items (detected: ${result.detected_lang})`);
    } catch (e) {
      status(`Error: ${e.message}`);
    } finally {
      setBusy(false);
    }
  });

  /* ── Translate & Export ─────────────────────────────────── */
  $("#btn-translate-export").addEventListener("click", async () => {
    if (menuItems.length === 0) {
      status("No items to export");
      return;
    }

    const descLangs = getDescLangs();
    status(strings.msg_translating || "Translating...");
    setBusy(true);

    try {
      const transResult = await api("/api/translate", {
        items: menuItems,
        desc_langs: descLangs,
        src_lang: elSrcLang.value,
        mode: elTransMode.value,
        api_key: elApiKey.value,
        provider: elProvider.value,
      });

      if (transResult.error) {
        status(`Error: ${transResult.error}`);
        return;
      }

      status(strings.msg_exporting || "Exporting...");
      const exportResult = await api("/api/export/full", {
        items: transResult.items,
        desc_langs: descLangs,
        output_dir: elAiOutput.value,
      });

      if (exportResult.error) {
        status(`Error: ${exportResult.error}`);
      } else {
        status(`Export completed: ${exportResult.output_file}`);
      }
    } catch (e) {
      status(`Error: ${e.message}`);
    } finally {
      setBusy(false);
    }
  });

  /* ── Save Simple Excel ──────────────────────────────────── */
  $("#btn-save-simple").addEventListener("click", async () => {
    if (menuItems.length === 0) {
      status("No items to save");
      return;
    }

    const descLangs = getDescLangs();
    status(strings.msg_translating || "Translating...");
    setBusy(true);

    try {
      const transResult = await api("/api/translate", {
        items: menuItems,
        desc_langs: descLangs,
        src_lang: elSrcLang.value,
        mode: elTransMode.value,
        api_key: elApiKey.value,
        provider: elProvider.value,
      });

      if (transResult.error) {
        status(`Error: ${transResult.error}`);
        return;
      }

      status(strings.msg_exporting || "Exporting...");
      const result = await api("/api/export/simple", {
        items: transResult.items,
        desc_langs: descLangs,
        output_dir: elAiOutput.value,
      });

      if (result.error) {
        status(`Error: ${result.error}`);
      } else {
        status(`Saved: ${result.output_file}`);
      }
    } catch (e) {
      status(`Error: ${e.message}`);
    } finally {
      setBusy(false);
    }
  });

  /* ── Browse buttons ─────────────────────────────────────── */
  $("#btn-ai-browse").addEventListener("click", async () => {
    const result = await api("/api/file-dialog", { type: "open_dir" });
    if (result.paths && result.paths.length) {
      elAiOutput.value = result.paths[0];
      api("/api/config/save", { output_path: result.paths[0] });
    }
  });

  $("#btn-excel-browse").addEventListener("click", async () => {
    const result = await api("/api/file-dialog", { type: "open_dir" });
    if (result.paths && result.paths.length) {
      elExcelOutput.value = result.paths[0];
      api("/api/config/save", { output_path: result.paths[0] });
    }
  });

  $("#btn-excel-select").addEventListener("click", async () => {
    const result = await api("/api/file-dialog", {
      type: "open_files",
      filetypes: ["Excel files (*.xlsx;*.xls)"],
    });
    if (result.paths && result.paths.length) {
      elExcelSource.value = result.paths[0];
    }
  });

  /* ── Excel Convert ──────────────────────────────────────── */
  $("#btn-excel-convert").addEventListener("click", async () => {
    const src = elExcelSource.value.trim();
    if (!src) {
      status(strings.msg_no_file || "Please select a file!");
      return;
    }

    status("Converting...");
    setBusy(true);

    try {
      const result = await api("/api/excel-import", {
        source_file: src,
        output_dir: elExcelOutput.value.trim(),
      });

      if (result.error) {
        status(`Error: ${result.error}`);
      } else {
        status(`Export completed: ${result.output_file}`);
      }
    } catch (e) {
      status(`Error: ${e.message}`);
    } finally {
      setBusy(false);
    }
  });

  /* ── Desc lang config save ──────────────────────────────── */
  for (let i = 1; i <= 4; i++) {
    $(`#desc-lang-${i}`).addEventListener("change", () => {
      api("/api/config/save", { desc_langs: getDescLangs() });
    });
  }

  /* ── Utilities ──────────────────────────────────────────── */
  function getDescLangs() {
    return [1, 2, 3, 4].map((i) => $(`#desc-lang-${i}`).value);
  }

  function setBusy(busy) {
    $$(".btn").forEach((b) => (b.disabled = busy));
  }

  /* ── Init ───────────────────────────────────────────────── */
  loadConfig().then(() => {
    renderFileGrid();
    renderTable();
    status("Ready");
  });
})();
