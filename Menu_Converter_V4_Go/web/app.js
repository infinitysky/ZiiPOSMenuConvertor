let strings = {};
let uiLang = "en";

async function api(path, body) {
  const opts = { method: body ? "POST" : "GET", headers: {} };
  if (body) {
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(path, opts);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || res.statusText);
  return data;
}

function showToast(msg, isError = false) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.style.background = isError ? "#c62828" : "#323232";
  el.classList.remove("hidden");
  setTimeout(() => el.classList.add("hidden"), 5000);
}

function applyStrings() {
  document.getElementById("title").textContent = strings.title || "";
  document.getElementById("tab-excel").textContent = strings.tab_excel || "";
  document.getElementById("tab-pe").textContent = strings.tab_pe || "";
  document.getElementById("lbl-menu-file").textContent = strings.lbl_menu_file || "";
  document.getElementById("lbl-pe-file").textContent = strings.lbl_pe_file || "";
  document.getElementById("lbl-output-excel").textContent = strings.lbl_output_folder || "";
  document.getElementById("lbl-output-pe").textContent = strings.lbl_output_folder || "";
  document.getElementById("btn-excel-file").textContent = strings.btn_select || "";
  document.getElementById("btn-pe-file").textContent = strings.btn_select || "";
  document.getElementById("btn-excel-out").textContent = strings.btn_browse || "";
  document.getElementById("btn-pe-out").textContent = strings.btn_browse || "";
  document.getElementById("btn-excel-convert").textContent = strings.btn_convert || "";
  document.getElementById("btn-pe-read").textContent = strings.btn_pe_read || "";
  document.getElementById("btn-pe-convert").textContent = strings.btn_pe_convert || "";
  document.querySelectorAll("#pe-table th[data-col]").forEach((th) => {
    const key = th.dataset.col;
    th.textContent = strings[key] || th.textContent;
  });
}

async function loadLang(lang) {
  uiLang = lang;
  strings = await api(`/api/i18n/${lang}`);
  applyStrings();
  document.querySelectorAll(".lang-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.lang === lang);
  });
  await api("/api/config", { ui_lang: lang });
}

async function pickFile(inputId) {
  const data = await api("/api/dialog/file", {});
  if (data.path) document.getElementById(inputId).value = data.path;
}

async function pickDir(inputId) {
  const initial = document.getElementById(inputId).value;
  const data = await api("/api/dialog/dir", { initial });
  if (data.path) document.getElementById(inputId).value = data.path;
}

function renderPETable(items) {
  const tbody = document.querySelector("#pe-table tbody");
  tbody.innerHTML = "";
  items.forEach((item, idx) => {
    const tr = document.createElement("tr");
    const status =
      item.status === "show" ? strings.status_show || "Show" : strings.status_hide || "Hide";
    const itemCode = (item.pe_id && String(item.pe_id).trim()) || String(idx + 1).padStart(4, "0");
    tr.innerHTML = `
      <td>${escapeHtml(itemCode)}</td>
      <td>${escapeHtml(item.name)}</td>
      <td>${item.price}</td>
      <td>${escapeHtml(item.category)}</td>
      <td>${escapeHtml(item.menu_group)}</td>
      <td>${item.tax_rate}</td>
      <td>${status}</td>
      <td>${item.has_image ? "Y" : ""}</td>`;
    tbody.appendChild(tr);
  });
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(`panel-${tab.dataset.tab}`).classList.add("active");
  });
});

document.querySelectorAll(".lang-btn").forEach((btn) => {
  btn.addEventListener("click", () => loadLang(btn.dataset.lang));
});

document.getElementById("btn-excel-file").addEventListener("click", () => pickFile("excel-source"));
document.getElementById("btn-pe-file").addEventListener("click", () => pickFile("pe-source"));
document.getElementById("btn-excel-out").addEventListener("click", () => pickDir("excel-output"));
document.getElementById("btn-pe-out").addEventListener("click", () => pickDir("pe-output"));

document.getElementById("btn-excel-convert").addEventListener("click", async () => {
  try {
    const data = await api("/api/excel/convert", {
      source_file: document.getElementById("excel-source").value,
      output_dir: document.getElementById("excel-output").value,
    });
    showToast(`${data.message}\n${data.output_file}`);
  } catch (e) {
    showToast(e.message, true);
  }
});

document.getElementById("btn-pe-read").addEventListener("click", async () => {
  const status = document.getElementById("pe-status");
  status.textContent = strings.msg_reading || "...";
  try {
    const data = await api("/api/pe/read", {
      source_file: document.getElementById("pe-source").value,
    });
    renderPETable(data.items || []);
    status.textContent = `OK - ${data.item_count} items, ${data.image_count} images`;
  } catch (e) {
    status.textContent = "";
    showToast(e.message, true);
  }
});

document.getElementById("btn-pe-convert").addEventListener("click", async () => {
  const status = document.getElementById("pe-status");
  status.textContent = strings.msg_exporting || "...";
  try {
    const data = await api("/api/pe/convert", {
      source_file: document.getElementById("pe-source").value,
      output_dir: document.getElementById("pe-output").value,
    });
    status.textContent = `OK - ${data.output_file} | ${data.image_count} images`;
    showToast(`${data.message}\n${data.output_file}\nImages: ${data.image_count} -> ${data.pics_dir}`);
  } catch (e) {
    status.textContent = "";
    showToast(e.message, true);
  }
});

(async function init() {
  const cfg = await api("/api/config");
  document.getElementById("excel-output").value = cfg.output_path || "C:\\Ziitech\\Menu";
  document.getElementById("pe-output").value = cfg.output_path || "C:\\Ziitech\\Menu";
  await loadLang(cfg.ui_lang || "en");
})();
