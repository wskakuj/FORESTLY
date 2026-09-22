/* Forestly — logika frontendu (most pywebview <-> Python) */
"use strict";

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

let SCHEMA = null;
let VALUES = {};
let RUNNING = false;
let activeTab = null;
let setValuesTimer = null;

/* Pokazywanie kontrolki warunkowej: id kontrolki -> id checkboxa */
const DEPENDS = {
  all_template: "all_str_tyt",
  all_skroty: "all_custom_skroty",
  xl_global_size: "xl_global_font",
  tpl_MIETEK_village: "tpl_MIETEK_single",
  tpl_MIETEK_area_v: "tpl_MIETEK_area",
  tpl_TAKSATOR_village: "tpl_TAKSATOR_single",
  tpl_TAKSATOR_area_v: "tpl_TAKSATOR_area",
};
/* Ikony sekcji w menu bocznym (SVG inline) */
const SECTION_ICONS = {
  "MIETEK": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>',
  "TAKSATOR": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8M8 17h8M8 9h2"/></svg>',
  "ROZLICZANIE": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><path d="M8 6h8M8 10h3M13 10h3M8 14h3M13 14h3M8 18h3"/></svg>',
  "KONWERTER PDF": '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>',
};

/* Ikony okna układu PDF */
const TRASH_SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>';
const RESTORE_SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>';

/* Grupy checkboxów z auto-powrotem do "Wszystkie", gdy nic nie zaznaczone */
const AUTO_RETURN_ALL = { word_filters: true, wydruki_filters: false };

/* ------------------------------------------------------------------ pywebview */
function api() { return window.pywebview.api; }
function ready() {
  return new Promise(res => {
    if (window.pywebview && window.pywebview.api) return res();
    window.addEventListener("pywebviewready", () => res());
  });
}

/* -------------------------------------------------------------------- render */
function renderAll(cfg) {
  SCHEMA = cfg.schema;
  VALUES = cfg.values || {};
  $("#app-version").textContent = "wersja " + (SCHEMA.version || "");
  document.title = SCHEMA.app_name;
  const nameEl = $("#app-name");
  if (nameEl && SCHEMA.app_name) nameEl.textContent = SCHEMA.app_name;

  const nav = $("#sidebar");
  nav.innerHTML = "";
  const content = $("#content-scroll");
  content.innerHTML = "";

  /* zwijane grupy sekcji */
  const collapsedSections = loadCollapsedSections();
  const groups = {};
  for (const tab of SCHEMA.tabs) {
    const [section, name] = tab.key.split("|");
    if (!groups[section]) {
      const g = document.createElement("div");
      g.className = "nav-group" + (collapsedSections.indexOf(section) >= 0 ? " collapsed" : "");
      g.dataset.section = section;
      const count = SCHEMA.tabs.filter(t => t.key.split("|")[0] === section).length;
      const header = document.createElement("div");
      header.className = "nav-section";
      const col = secColor(section);
      header.innerHTML = '<span class="ns-ic" style="color:' + col.fg + ";background:" + col.bg + '">' +
        (SECTION_ICONS[section] || "") + "</span>" +
        "<span>" + escapeHtml(section) + "</span>" +
        '<span class="nav-count">' + count + "</span>" +
        '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';
      header.onclick = () => toggleSection(g, section);
      g.appendChild(header);
      nav.appendChild(g);
      groups[section] = g;
    }
    const item = document.createElement("div");
    item.className = "nav-item";
    item.dataset.key = tab.key;
    item.innerHTML = `<span class="dot"></span><span>${escapeHtml(name)}</span>`;
    item.title = tab.tooltip || "";
    item.onclick = () => showTab(tab.key);
    groups[section].appendChild(item);

    const view = document.createElement("div");
    view.className = "tab-view";
    view.dataset.key = tab.key;
    view.innerHTML = `<h1>${escapeHtml(name)}</h1>` +
      (tab.tooltip ? `<div class="tab-desc">${escapeHtml(tab.tooltip)}</div>` : "");
    renderControls(view, tab);
    content.appendChild(view);
  }
  buildStartView(nav, content);
  showTab(START_KEY);
  applyDeps();
}

/* ------------------------------------------- nawigacja: Start, grupy, szukajka */
const START_KEY = "__start__";
const HOME_SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>';
const SECTION_DESC = {
  "MIETEK": "Wydruki TXT z DBF, konwersje Word / PDF, scalanie i wykazy",
  "TAKSATOR": "Szablony STR_TYT, układanie Exceli i PDF dla wsi",
  "ROZLICZANIE": "Powierzchnie, mietki, halizny i Excel z MDB",
  "KONWERTER PDF": "Dowolne dokumenty i obrazy na PDF",
};

const SECTION_COLORS = {
  "MIETEK": { "fg": "#2dd4a7", "bg": "rgba(45, 212, 167, 0.14)" },
  "TAKSATOR": { "fg": "#38bdf8", "bg": "rgba(56, 163, 248, 0.14)" },
  "ROZLICZANIE": { "fg": "#f5b942", "bg": "rgba(245, 185, 66, 0.14)" },
  "KONWERTER PDF": { "fg": "#a78bfa", "bg": "rgba(167, 139, 250, 0.14)" },
};
function secColor(sec) { return SECTION_COLORS[sec] || { "fg": "var(--accent)", "bg": "rgba(45, 212, 167, 0.12)" }; }
function secKey(sec) { return "__sec__" + sec; }

function loadCollapsedSections() {
  try { return JSON.parse(localStorage.getItem("klp-nav-collapsed") || "[]"); }
  catch (e) { return []; }
}
function saveCollapsedSections(list) {
  try { localStorage.setItem("klp-nav-collapsed", JSON.stringify(list)); } catch (e) {}
}
function toggleSection(g, section) {
  const closed = g.classList.toggle("collapsed");
  const list = loadCollapsedSections().filter(s => s !== section);
  if (closed) list.push(section);
  saveCollapsedSections(list);
}
function buildStartView(nav, content) {
  const item = document.createElement("div");
  item.className = "nav-item start";
  item.dataset.key = START_KEY;
  item.innerHTML = '<span class="ns-ic ns-start">' + HOME_SVG + "</span><span>Start</span>";
  item.title = "Ekran startowy";
  item.onclick = () => showTab(START_KEY);
  nav.insertBefore(item, nav.firstChild);

  const firstOf = {};
  for (const t of SCHEMA.tabs) {
    const sec = t.key.split("|")[0];
    if (!firstOf[sec]) firstOf[sec] = t.key;
  }
  const autoTab = (SCHEMA.tabs.find(t => t.key.indexOf("Pełny Automat") >= 0) || {}).key || "";

  const cards = Object.keys(firstOf).map(sec => {
    const col = secColor(sec);
    return '<div class="start-card" data-target="' + escapeHtml(secKey(sec)) + '">' +
      '<div class="sc-ic" style="color:' + col.fg + ";background:" + col.bg + '">' + (SECTION_ICONS[sec] || "") + "</div>" +
      '<div class="sc-body"><div class="sc-title">' + escapeHtml(sec) + "</div>" +
      '<div class="sc-sub">' + escapeHtml(SECTION_DESC[sec] || "") + "</div>" +
      '<div class="sc-go">Zobacz ' + firstOf[sec].split("|")[1].split("(")[0].trim() + " →</div></div>" +
    "</div>";
  }).join("");

  const view = document.createElement("div");
  view.className = "tab-view hidden";
  view.dataset.key = START_KEY;
  view.innerHTML =
    "<h1>Start</h1>" +
    '<div class="tab-desc">Wybierz, co chcesz dziś zrobić — najczęściej używane funkcje są na górze.</div>' +
    '<div class="start-hero" data-target="' + escapeHtml(autoTab) + '">' +
      '<div class="sh-ic">' + (SECTION_ICONS["MIETEK"] || "") + "</div>" +
      '<div class="sh-txt"><div class="sh-title">Pełny Automat (1-Click)</div>' +
      '<div class="sh-sub">Halizny → TXT → Word → PDF → scalanie. Cały proces jednym kliknięciem.</div></div>' +
      '<div class="sh-go">Otwórz →</div>' +
    "</div>" +
    '<div class="start-grid">' + cards + "</div>";
  $$(".start-hero,.start-card", view).forEach(c => {
    if (c.dataset.target) c.onclick = () => showTab(c.dataset.target);
  });
  content.insertBefore(view, content.firstChild);
  buildSectionViews(content);
}

function buildSectionViews(content) {
  const bySec = {};
  for (const t of SCHEMA.tabs) {
    const sec = t.key.split("|")[0];
    (bySec[sec] = bySec[sec] || []).push(t);
  }
  for (const sec of Object.keys(bySec)) {
    const col = secColor(sec);
    const back = document.createElement("div");
    back.className = "sec-back btn ghost small";
    back.textContent = "← Start";
    back.onclick = () => showTab(START_KEY);

    const tabsHtml = bySec[sec].map(t => {
      const tip = (t.tooltip || "").replace(/\s*\n\s*/g, " ").trim();
      return '<div class="sec-card" data-target="' + escapeHtml(t.key) + '">' +
        '<div class="sc-title">' + escapeHtml(t.key.split("|")[1]) + "</div>" +
        (tip ? '<div class="sc-sub">' + escapeHtml(tip) + "</div>" : "") +
        '<div class="sc-go">Otwórz →</div></div>';
    }).join("");

    const view = document.createElement("div");
    view.className = "tab-view hidden";
    view.dataset.key = secKey(sec);
    view.innerHTML =
      '<div class="sec-head"><span class="sec-ic" style="color:' + col.fg + ";background:" + col.bg + '">' +
        (SECTION_ICONS[sec] || "") + "</span>" +
        "<h1>" + escapeHtml(sec) + "</h1></div>" +
      '<div class="tab-desc">' + escapeHtml(SECTION_DESC[sec] || "") + "</div>" +
      '<div class="sec-grid">' + tabsHtml + "</div>";
    view.insertBefore(back, view.firstChild);
    $$(".sec-card", view).forEach(c => {
      if (c.dataset.target) c.onclick = () => showTab(c.dataset.target);
    });
    content.appendChild(view);
  }
}

function showTab(key) {
  activeTab = key;
  $$(".tab-view").forEach(v => v.classList.toggle("hidden", v.dataset.key !== key));
  $$(".nav-item").forEach(i => i.classList.toggle("active", i.dataset.key === key));
  const active = $$(".nav-item").find(i => i.dataset.key === key);
  if (active) {
    const g = active.closest(".nav-group");
    if (g) g.classList.remove("collapsed");
    active.scrollIntoView({ block: "nearest" });
  }
}

function renderControls(view, tab) {
  for (const c of tab.controls) {
    switch (c.kind) {
      case "path": view.appendChild(renderPath(c)); break;
      case "text": view.appendChild(renderText(c)); break;
      case "check": view.appendChild(renderCheck(c)); break;
      case "checks": view.appendChild(renderChecks(c)); break;
      case "select": view.appendChild(renderSelect(c)); break;
      case "margins": view.appendChild(renderMargins(c)); break;
      case "fonts": view.appendChild(renderFonts(c)); break;
      case "dashboard": view.appendChild(renderDashboard(c)); break;
      case "info": view.appendChild(renderInfo(c)); break;
    }
  }
  const actions = document.createElement("div");
  actions.className = "actions";
  for (const b of tab.buttons) {
    const btn = document.createElement("button");
    btn.className = "btn " + (b.style === "secondary" ? "secondary" : "primary run-btn");
    btn.textContent = b.label;
    btn.title = b.tooltip || "";
    if (b.style !== "secondary") btn.classList.add("run-big");
    btn.onclick = () => runTask(b.task);
    actions.appendChild(btn);
  }
  view.appendChild(actions);
}

function el(tag, cls, html) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html !== undefined) e.innerHTML = html;
  return e;
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}
function cssEscape(s) { return s.replace(/"/g, '\\"'); }

/* kontrolka: ścieżka */
function renderPath(c) {
  const row = el("div", "field");
  row.innerHTML = `<label>${escapeHtml(c.label)}</label>`;
  const group = el("div", "input-group");
  const input = el("input");
  input.type = "text"; input.placeholder = c.ph || "";
  input.dataset.cid = c.id; input.dataset.kind = "path";
  input.value = VALUES[c.id] || "";
  input.oninput = scheduleSetValues;
  group.appendChild(input);

  const browse = el("button", "btn secondary", "Przeglądaj");
  browse.onclick = async () => {
    const r = await api().browse(c.id, c.browse || "folder");
    if (r.path) { input.value = r.path; scheduleSetValues(); }
  };
  group.appendChild(browse);

  const hist = el("button", "btn secondary", "🕒");
  hist.title = "Pokaż historię ostatnio używanych folderów";
  hist.onclick = async () => {
    const r = await api().get_history();
    showHistoryMenu(hist, r.history || [], input);
  };
  group.appendChild(hist);
  row.appendChild(group);
  return row;
}

function showHistoryMenu(anchor, history, input) {
  $$(".hist-menu").forEach(m => m.remove());
  if (!history.length) { toast("Historia folderów jest pusta", "warn"); return; }
  const menu = el("div", "modal-backdrop");
  const modal = el("div", "modal");
  modal.style.width = "560px";
  modal.appendChild(el("h3", null, "Ostatnio używane foldery"));
  const list = el("div", "order-list");
  history.forEach(p => {
    const it = el("div", "order-item");
    it.innerHTML = `<span class="grip">📂</span><span class="oi-name" title="${escapeHtml(p)}">${escapeHtml(p)}</span>`;
    it.style.cursor = "pointer";
    it.onclick = () => { input.value = p; scheduleSetValues(); menu.remove(); };
    list.appendChild(it);
  });
  modal.appendChild(list);
  const row = el("div", "modal-row");
  const close = el("button", "btn secondary", "Zamknij");
  close.onclick = () => menu.remove();
  row.appendChild(close); modal.appendChild(row);
  menu.appendChild(modal); menu.onclick = e => { if (e.target === menu) menu.remove(); };
  $("#modal-root").appendChild(menu);
}

/* kontrolka: tekst */
function renderText(c) {
  const row = el("div", "field");
  row.innerHTML = `<label>${escapeHtml(c.label)}</label>`;
  const input = el("input");
  input.type = "text"; input.placeholder = c.ph || "";
  input.dataset.cid = c.id; input.dataset.kind = "text";
  input.value = VALUES[c.id] !== undefined ? VALUES[c.id] : (c.default || "");
  input.oninput = scheduleSetValues;
  row.appendChild(input);
  return row;
}

/* kontrolka: checkbox */
function renderCheck(c) {
  const row = el("label", "check-row");
  const input = el("input");
  input.type = "checkbox"; input.dataset.cid = c.id; input.dataset.kind = "check";
  input.checked = !!VALUES[c.id];
  input.onchange = () => { applyDeps(); scheduleSetValues(); };
  row.appendChild(input);
  const span = el("span", null, escapeHtml(c.label) +
    (c.tooltip ? ` <span class="hint" title="${escapeHtml(c.tooltip)}">ⓘ</span>` : ""));
  row.appendChild(span);
  if (c.tooltip) row.title = c.tooltip;
  return row;
}

/* kontrolka: grupa checkboxów z "Wszystkie" */
function renderChecks(c) {
  const wrap = el("div", null);
  wrap.innerHTML = `<div class="subtitle">${escapeHtml(c.label)}</div>`;
  const grid = el("div", "checks-grid");
  for (const choice of c.choices) {
    const row = el("label", "check-row");
    row.dataset.choice = choice;
    const input = el("input");
    input.type = "checkbox";
    input.dataset.cid = c.id; input.dataset.kind = "checks-item"; input.dataset.choice = choice;
    const st = (VALUES[c.id] || {});
    input.checked = !!st[choice];
    input.onchange = () => onGroupChange(c, grid);
    row.appendChild(input);
    row.appendChild(el("span", null, escapeHtml(choice)));
    grid.appendChild(row);
  }
  applyGroupStates(c, grid);
  if (c.tooltip) grid.title = c.tooltip;
  wrap.appendChild(grid);
  return wrap;
}

function onGroupChange(c, grid) {
  const all = grid.querySelector('[data-choice="Wszystkie"] input');
  const changed = document.activeElement && document.activeElement.dataset.choice
    ? document.activeElement.dataset.choice : null;
  if (changed === "Wszystkie" && all.checked) {
    grid.querySelectorAll("input").forEach(i => { if (i.dataset.choice !== "Wszystkie") i.checked = false; });
  } else if (changed !== "Wszystkie") {
    if (all.checked) all.checked = false;
    if (AUTO_RETURN_ALL[c.id]) {
      const any = grid.querySelectorAll("input").some ? Array.from(grid.querySelectorAll("input"))
        .some(i => i.checked && i.dataset.choice !== "Wszystkie") : false;
      if (!any) all.checked = true;
    }
  }
  applyGroupStates(c, grid);
  scheduleSetValues();
}

function applyGroupStates(c, grid) {
  const all = grid.querySelector('[data-choice="Wszystkie"] input');
  if (!all) return;
  grid.querySelectorAll(".check-row").forEach(r => {
    if (r.dataset.choice !== "Wszystkie") r.classList.toggle("disabled", all.checked);
  });
}

/* kontrolka: select */
function renderSelect(c) {
  const row = el("div", "field");
  row.innerHTML = `<label>${escapeHtml(c.label)}</label>`;
  const sel = el("select");
  sel.dataset.cid = c.id; sel.dataset.kind = "select";
  const current = VALUES[c.id] !== undefined ? VALUES[c.id] : c.default;
  for (const v of (c.values || [])) {
    const o = el("option", null, escapeHtml(v));
    o.value = v;
    if (v === current) o.selected = true;
    sel.appendChild(o);
  }
  if (c.free && current && !(c.values || []).includes(current)) {
    const o = el("option", null, escapeHtml(current));
    o.value = current; o.selected = true;
    sel.appendChild(o);
  }
  sel.onchange = scheduleSetValues;
  row.appendChild(sel);
  return row;
}

/* kontrolka: marginesy */
const MARGIN_TYPES = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "HALIZNY", "WYK_NEG", "OPIS", "ZEST1", "WK_ZM1"];
const MARGIN_SIDES = ["T", "B", "L", "R"];

function renderMargins(c) {
  const wrap = el("div", null);
  wrap.innerHTML = `<div class="subtitle">${escapeHtml(c.label)}</div>`;
  const table = el("table", "margins-table");
  table.dataset.cid = c.id; table.dataset.kind = "margins";
  const head = el("tr", null, "<th>Typ pliku</th><th>Góra</th><th>Dół</th><th>Lewo</th><th>Prawo</th>");
  table.appendChild(head);
  const data = VALUES[c.id] || {};
  for (const ftype of MARGIN_TYPES) {
    const tr = el("tr");
    tr.dataset.ftype = ftype;
    tr.appendChild(el("td", null, ftype));
    for (const side of MARGIN_SIDES) {
      const td = el("td");
      const inp = el("input");
      inp.type = "text";
      inp.dataset.side = side;
      inp.value = (data[ftype] && data[ftype][side] !== undefined) ? data[ftype][side] : "1.5";
      inp.oninput = scheduleSetValues;
      td.appendChild(inp);
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  wrap.appendChild(table);
  return wrap;
}

/* kontrolka: czcionki Excel */
function renderFonts(c) {
  const wrap = el("div", null);
  wrap.innerHTML = `<div class="subtitle">${escapeHtml(c.label)}</div>`;
  const grid = el("div", "fonts-grid");
  grid.dataset.cid = c.id; grid.dataset.kind = "fonts";
  for (const f of c.fonts) {
    const row = el("div", "font-row");
    row.dataset.sheet = f.sheet;
    row.innerHTML = `<label>${escapeHtml(f.sheet)} (od w. ${f.start_row}):</label>`;
    const inp = el("input");
    inp.type = "text";
    inp.value = (VALUES[c.id] && VALUES[c.id][f.sheet] !== undefined) ? VALUES[c.id][f.sheet] : f.default;
    inp.oninput = scheduleSetValues;
    row.appendChild(inp);
    grid.appendChild(row);
  }
  wrap.appendChild(grid);
  return wrap;
}

/* dashboard */
function renderDashboard(c) {
  const wrap = el("div", "card");
  wrap.dataset.cid = c.id;
  wrap.innerHTML = `<div class="subtitle">Postęp procesu</div>`;
  const dash = el("div", "dashboard");
  dash.id = "dashboard";
  c.steps.forEach((name, i) => {
    const st = el("div", "dstep");
    st.dataset.step = i;
    st.innerHTML = `<div class="num">KROK ${i + 1}</div><div class="name">${escapeHtml(name)}</div><div class="state"></div>`;
    dash.appendChild(st);
  });
  wrap.appendChild(dash);
  return wrap;
}

function renderInfo(c) {
  return el("div", "info", escapeHtml(c.text));
}

/* ------------------------------------------------------- zależności widoczności */
function isCheckedEl(el) {
  if (!el) return false;
  if (el.tagName === "INPUT") return !!el.checked;
  const inp = el.querySelector("input");
  return inp ? !!inp.checked : false;
}

function applyDeps() {
  for (const [cid, dep] of Object.entries(DEPENDS)) {
    const row = $(`[data-cid="${cssEscape(cid)}"]`);
    const depEl = $(`[data-cid="${cssEscape(dep)}"]`);
    if (row && depEl) {
      const checked = isCheckedEl(depEl);
      const field = row.closest(".field") || row;
      field.classList.toggle("hidden", !checked);
    }
  }
}

/* --------------------------------------------------------- zbieranie wartości */
function collectValues() {
  const out = {};
  $$("[data-cid]").forEach(node => {
    const cid = node.dataset.cid;
    const kind = node.dataset.kind;
    if (kind === "path" || kind === "text") {
      out[cid] = node.value;
    } else if (kind === "check") {
      out[cid] = node.checked;
    } else if (kind === "select") {
      out[cid] = node.value;
    } else if (kind === "checks-item") {
      out[cid] = out[cid] || {};
      out[cid][node.dataset.choice] = node.checked;
    } else if (kind === "margins") {
      const data = {};
      $$("tr[data-ftype]", node).forEach(tr => {
        data[tr.dataset.ftype] = {};
        $$("input", tr).forEach(inp => { data[tr.dataset.ftype][inp.dataset.side] = inp.value; });
      });
      out[cid] = data;
    } else if (kind === "fonts") {
      const data = {};
      $$(".font-row", node).forEach(r => {
        data[r.dataset.sheet] = r.querySelector("input").value;
      });
      out[cid] = data;
    }
  });
  return out;
}

function scheduleSetValues() {
  clearTimeout(setValuesTimer);
  setValuesTimer = setTimeout(async () => {
    try { await api().set_values(collectValues()); } catch (e) { /* noop */ }
  }, 400);
}

/* ------------------------------------------------------------------ uruchomienie */
async function runTask(task) {
  if (RUNNING) { toast("Zadanie już trwa — najpierw zatrzymaj bieżące.", "warn"); return; }
  try { await api().set_values(collectValues()); } catch (e) { /* noop */ }
  /* zadania czysto dialogowe — nie idą do backendu jako procesy */
  if (task === "web_manual_merge") { openManualMerge(); return; }
  if (task.startsWith("open_order:")) { openOrderDialog(task.split(":")[1]); return; }

  const r = await api().run(task);
  if (!r.ok) {
    toast(r.error || "Nie udało się uruchomić zadania.", "error");
  }
}

/* --------------------------------------------------------------------- polling */
async function pollLoop() {
  try {
    const events = await api().poll();
    for (const ev of events) handleEvent(ev);
  } catch (e) { /* pywebview jeszcze niegotowy */ }
  setTimeout(pollLoop, 220);
}

function handleEvent(ev) {
  switch (ev.type) {
    case "log": appendLog(ev.text); break;
    case "clear_log": $("#log").innerHTML = ""; break;
    case "status":
      $("#status-text").textContent = ev.text;
      $("#status-dot").className = ev.color === "#D83B01" ? "err" : "";
      break;
    case "progress": {
      const fill = $("#progress-fill");
      const val = typeof ev.value === "number" ? ev.value : parseFloat(ev.value) || 0;
      fill.style.width = Math.max(0, Math.min(100, val * 100)) + "%";
      const fileEl = $("#progress-file");
      let t = "";
      if (ev.desc) t = ev.desc;
      if (ev.file) t += (t ? " — " : "") + ev.file;
      if (ev.current != null && ev.total) t += ` (${ev.current}/${ev.total})`;
      fileEl.textContent = t;
      break;
    }
    case "dashboard": {
      const st = $(`#dashboard .dstep[data-step="${ev.step}"]`);
      if (st) {
        st.classList.remove("running", "done", "error");
        if (ev.status) st.classList.add(ev.status);
        st.querySelector(".state").textContent = ev.text || "";
      }
      break;
    }
    case "dashboard_reset":
      $$("#dashboard .dstep").forEach(st => {
        st.classList.remove("running", "done", "error");
        st.querySelector(".state").textContent = "";
      });
      break;
    case "dialog": showDialog(ev); break;
    case "state": setRunning(!!ev.running); break;
  }
}

function appendLog(text) {
  const log = $("#log");
  const div = el("div", null, escapeHtml(text));
  if (/BŁĄD|Error|Traceback/i.test(text)) div.className = "err";
  else if (/\[UWAGA\]|warn/i.test(text)) div.className = "warn";
  else if (/ZAKOŃCZONO|pomyślnie/i.test(text)) div.className = "ok";
  log.appendChild(div);
  while (log.children.length > 3000) log.removeChild(log.firstChild);
  log.scrollTop = log.scrollHeight;
}

function setRunning(running) {
  RUNNING = running;
  $$(".run-btn").forEach(b => b.disabled = running);
  $("#btn-stop").classList.toggle("hidden", !running);
  const dot = $("#status-dot");
  if (running) { dot.className = "busy"; }
  else { dot.className = ""; }
}

/* -------------------------------------------------------------------- dialogi */
function showDialog(ev) {
  if (ev.kind === "confirm") {
    if (window.confirm(ev.title + "\n\n" + ev.message)) {
      api().dialog_reply(ev.id, true);
    } else {
      api().dialog_reply(ev.id, false);
    }
  } else {
    const t = toast(ev.message, ev.kind === "error" ? "error" : ev.kind === "warn" ? "warn" : "ok");
    const root = $("#toast-root");
    root.appendChild(t);
  }
}

function toast(msg, kind) {
  const t = el("div", "toast " + (kind || ""), escapeHtml(msg));
  $("#toast-root").appendChild(t);
  setTimeout(() => t.remove(), 6000);
  return t;
}

/* ------------------------------------------------- dialog układu PDF (kolejność) */
async function openOrderDialog(mode) {
  const r = await api().get_pdf_order(mode);
  if (!r.ok) { toast(r.error || "Błąd", "error"); return; }
  const templates = SCHEMA.pdf_order_templates;
  let order = r.order.slice();
  let excluded = (r.excluded || []).slice();
  // klucze wykluczone, których nie ma w kolejności — dokładamy na spód listy
  for (const k of excluded) if (!order.includes(k)) order.push(k);

  const backdrop = el("div", "modal-backdrop");
  const modal = el("div", "modal");
  modal.appendChild(el("h3", null, "Skonfiguruj układ PDF"));
  modal.appendChild(el("div", "modal-msg",
    "Ustal kolejność elementów w scalanym dokumencie PDF (tryb " + escapeHtml(mode) + ").\n" +
    "Przełącznik po prawej stronie wączy/wyłącza element: WYŁĄCZONY element " +
    "nie trafi do finalnego scalonego PDF."));
  const list = el("div", "order-list");

  function isIncluded(key) { return !excluded.includes(key); }

  function renderList() {
    list.innerHTML = "";
    order.forEach((key, i) => {
      const tpl = templates.find(t => t.key === key);
      const on = isIncluded(key);
      const it = el("div", "order-item" + (on ? "" : " excluded"));
      it.innerHTML = `<span class="grip">☰</span>` +
        `<span class="oi-name">${i + 1}. ${escapeHtml(tpl ? tpl.label : key)}` +
        `${on ? "" : ' <span class="excl-badge">WYKLUCZONY</span>'}</span>`;
      const up = el("button", null, "▲");
      up.title = "W górę";
      up.onclick = () => { if (i > 0) { [order[i - 1], order[i]] = [order[i], order[i - 1]]; renderList(); } };
      const dn = el("button", null, "▼");
      dn.title = "W dół";
      dn.onclick = () => { if (i < order.length - 1) { [order[i + 1], order[i]] = [order[i], order[i + 1]]; renderList(); } };
      const tg = el("button", "order-toggle" + (on ? "" : " off"),
        on ? TRASH_SVG : RESTORE_SVG);
      tg.title = on ? "Wyklucz z finalnego PDF" : "Przywróć do scalania";
      tg.onclick = () => {
        if (on) excluded.push(key);
        else excluded = excluded.filter(k => k !== key);
        renderList();
      };
      it.appendChild(up); it.appendChild(dn); it.appendChild(tg);
      list.appendChild(it);
    });
  }
  renderList();
  modal.appendChild(list);
  const row = el("div", "modal-row");
  const cancel = el("button", "btn secondary", "Anuluj");
  cancel.onclick = () => backdrop.remove();
  const save = el("button", "btn primary", "Zapisz układ");
  save.onclick = async () => {
    const active = order.filter(k => !excluded.includes(k));
    const rr = await api().save_pdf_order(mode, active, excluded);
    if (rr.ok) { toast("Zapisano układ PDF.", "ok"); backdrop.remove(); }
    else toast(rr.error || "Błąd", "error");
  };
  row.appendChild(cancel); row.appendChild(save);
  modal.appendChild(row);
  backdrop.appendChild(modal);
  $("#modal-root").appendChild(backdrop);
}

/* --------------------------------------------------- ręczne scalanie PDF (web) */
async function openManualMerge() {
  const r = await api().manual_merge_list();
  if (!r.ok) { toast(r.error || "Błąd", "error"); return; }
  let files = r.files.slice();

  const backdrop = el("div", "modal-backdrop");
  const modal = el("div", "modal");
  modal.style.width = "600px";
  modal.appendChild(el("h3", null, "Ręczne scalanie PDF"));
  modal.appendChild(el("div", "modal-msg",
    "Ustaw kolejność plików z folderu:\n" + escapeHtml(r.src)));
  const list = el("div", "order-list");
  function renderList() {
    list.innerHTML = "";
    if (!files.length) { list.appendChild(el("div", "order-item", "Brak plików PDF w folderze źródłowym.")); return; }
    files.forEach((name, i) => {
      const it = el("div", "order-item");
      it.innerHTML = `<span class="grip">☰</span><span class="oi-name">${i + 1}. ${escapeHtml(name)}</span>`;
      const up = el("button", null, "▲");
      up.onclick = () => { if (i > 0) { [files[i - 1], files[i]] = [files[i], files[i - 1]]; renderList(); } };
      const dn = el("button", null, "▼");
      dn.onclick = () => { if (i < files.length - 1) { [files[i + 1], files[i]] = [files[i], files[i + 1]]; renderList(); } };
      const rm = el("button", null, "✕");
      rm.title = "Pomiń plik";
      rm.onclick = () => { files.splice(i, 1); renderList(); };
      it.appendChild(up); it.appendChild(dn); it.appendChild(rm);
      list.appendChild(it);
    });
  }
  renderList();
  modal.appendChild(list);
  const nameRow = el("div", "field");
  nameRow.style.marginTop = "12px";
  nameRow.innerHTML = `<label>Nazwa pliku wynikowego:</label>`;
  const nameInp = el("input");
  nameInp.type = "text"; nameInp.value = "PDF_polaczony.pdf";
  nameRow.appendChild(nameInp);
  modal.appendChild(nameRow);
  const row = el("div", "modal-row");
  const cancel = el("button", "btn secondary", "Anuluj");
  cancel.onclick = () => backdrop.remove();
  const merge = el("button", "btn primary", "Scal pliki");
  merge.onclick = async () => {
    merge.disabled = true;
    const rr = await api().manual_merge(files, nameInp.value.trim() || "PDF_polaczony.pdf");
    merge.disabled = false;
    if (rr.ok) { toast("Scalono: " + rr.file, "ok"); backdrop.remove(); }
    else toast(rr.error || "Błąd scalania", "error");
  };
  row.appendChild(cancel); row.appendChild(merge);
  modal.appendChild(row);
  backdrop.appendChild(modal);
  $("#modal-root").appendChild(backdrop);
}

/* ------------------------------------------------------------------- start */
window.addEventListener("DOMContentLoaded", async () => {
  $("#btn-clear-log").onclick = async () => { await api().clear_log(); $("#log").innerHTML = ""; };
  $("#btn-log-toggle").onclick = () => {
    const p = $("#log-panel");
    p.classList.toggle("collapsed");
    $("#btn-log-toggle").textContent = p.classList.contains("collapsed") ? "▼" : "▲";
  };
  $("#btn-stop").onclick = async () => {
    await api().stop();
    toast("Zatrzymywanie — program zakończy po bieżącym kroku…", "warn");
  };
  $("#btn-check-update").onclick = async () => {
    toast("Sprawdzam aktualizacje…");
    await api().check_update();
  };
  const themeBtn = $("#btn-theme");
  const applyTheme = t => {
    document.documentElement.dataset.theme = t;
    themeBtn.textContent = t === "light" ? "🌙 Ciemny" : "☀ Jasny";
  };
  applyTheme(localStorage.getItem("klp-theme") || "dark");
  themeBtn.onclick = () => {
    const t = document.documentElement.dataset.theme === "light" ? "dark" : "light";
    localStorage.setItem("klp-theme", t);
    applyTheme(t);
  };
  await ready();
  const cfg = await api().get_config();
  renderAll(cfg);
  pollLoop();
});
