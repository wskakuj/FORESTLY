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
      g.className = "nav-group" +
        ((collapsedSections === null || collapsedSections.indexOf(section) >= 0)
          ? " collapsed" : "");
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
  wireTerritory();
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
  /* brak zapisu w localStorage (pierwsze uruchomienie) => wszystkie grupy
     zwinięte domyślnie; null = "nie było wyboru użytkownika" */
  try {
    const raw = localStorage.getItem("klp-nav-collapsed");
    if (raw === null) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
}
function saveCollapsedSections(list) {
  try { localStorage.setItem("klp-nav-collapsed", JSON.stringify(list)); } catch (e) {}
}
function toggleSection(g, section) {
  const closed = g.classList.toggle("collapsed");
  const list = (loadCollapsedSections() || []).filter(s => s !== section);
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
  /* baza GDOŚ ładuje się dopiero przy pierwszym otwarciu zakładki */
  const g = document.querySelector('.tab-view:not(.hidden) .gdos-wrap[data-lazy="1"]');
  if (g) gdosLoad();
  const active = $$(".nav-item").find(i => i.dataset.key === key);
  if (active) {
    const g = active.closest(".nav-group");
    if (g) g.classList.remove("collapsed");
    active.scrollIntoView({ block: "nearest" });
  }
}

function renderOneControl(c) {
  switch (c.kind) {
    case "path": return renderPath(c);
    case "text": return renderText(c);
    case "check": return renderCheck(c);
    case "checks": return renderChecks(c);
    case "select": return renderSelect(c);
    case "margins": return renderMargins(c);
    case "fonts": return renderFonts(c);
    case "dashboard": return renderDashboard(c);
    case "info": return renderInfo(c);
    case "gdos_table": return renderGdos(c);
    case "group": return renderGroup(c);
  }
  return null;
}

function renderGroup(c) {
  /* zwijana grupa kontrolek (np. kreator strony tytułowej w 1-Click) */
  const wrap = el("div", null);
  const det = el("details", "group-details");
  if (c.collapsed === false) det.open = true;
  const sum = el("summary", null, escapeHtml(c.label));
  if (c.tooltip) sum.title = c.tooltip;
  det.appendChild(sum);
  const inner = el("div", "group-body");
  for (const sub of (c.controls || [])) {
    const node = renderOneControl(sub);
    if (node) inner.appendChild(node);
  }
  det.appendChild(inner);
  wrap.appendChild(det);
  return wrap;
}

function renderControls(view, tab) {
  for (const c of tab.controls) {
    const node = renderOneControl(c);
    if (node) view.appendChild(node);
  }
  const actions = document.createElement("div");
  actions.className = "actions";
  for (const b of tab.buttons) {
    const btn = document.createElement("button");
    const isRun = b.style !== "secondary";
    btn.className = "btn " + (isRun ? "primary run-btn" : "secondary");
    btn.innerHTML = (isRun ? ICON("play") : "") + "<span>" + escapeHtml(b.label) + "</span>";
    btn.title = b.tooltip || "";
    if (b.style !== "secondary") btn.classList.add("run-big");
    btn.onclick = () => { LAST_TASK_LABEL = b.label; runTask(b.task); };
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

/* ------------------------------------------------------------- ikony SVG */
const ICONS = {
  play:    '<polygon points="7 4 21 12 7 20"/>',
  folder:  '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
  clock:   '<circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 15.5 14"/>',
  stop:    '<rect x="5" y="5" width="14" height="14" rx="2"/>',
  sun:     '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  refresh: '<polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>',
  trash:   '<polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>',
  plus:    '<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>',
  save:    '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>',
  chevron: '<polyline points="6 9 12 15 18 9"/>'
};
function ICON(name) {
  const path = ICONS[name];
  if (!path) return "";
  return '<svg class="bi" viewBox="0 0 24 24" fill="none" stroke="currentColor" ' +
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + path + "</svg>";
}

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

  const browse = el("button", "btn secondary");
  browse.innerHTML = ICON("folder") + "<span>Przeglądaj</span>";
  browse.onclick = async () => {
    const r = await api().browse(c.id, c.browse || "folder");
    if (r.path) { input.value = r.path; scheduleSetValues(); }
  };
  group.appendChild(browse);

  const hist = el("button", "btn secondary");
  hist.innerHTML = ICON("clock");
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
/* rozwijana lista z możliwością wpisania własnej wartości (woj./powiat/gmina).
   dlId — id ukrytego <datalist> jako źródła opcji (aktualizowanego przez wireTerritory) */
function buildCombo(dlId, current) {
  const box = el("div", "combo");
  const inp = el("input");
  inp.type = "text";
  inp.value = current || "";
  inp.placeholder = "wybierz z listy lub wpisz własne";
  const btn = el("div", "combo-btn",
    '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" ' +
    'stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">' +
    '<polyline points="6 9 12 15 18 9"/></svg>');
  const list = el("div", "combo-list");
  box.appendChild(inp);
  box.appendChild(btn);
  box.appendChild(list);

  const options = () => {
    const dl = document.getElementById(dlId);
    return dl ? [...dl.options].map(o => o.value) : [];
  };
  const isOpen = () => box.classList.contains("open");
  const close = () => box.classList.remove("open");

  const buildList = (filter) => {
    list.innerHTML = "";
    const f = (filter || "").trim().toUpperCase();
    const opts = options().filter(v => !f || v.toUpperCase().includes(f));
    if (!opts.length) {
      list.appendChild(el("div", "combo-empty",
        "brak pozycji — wpisz własną wartość"));
      return;
    }
    opts.forEach(v => {
      const it = el("div", "combo-opt" + (v === inp.value ? " sel" : ""),
                    escapeHtml(v));
      it.onmousedown = e => e.preventDefault();
      it.onclick = () => {
        inp.value = v;
        close();
        inp.dispatchEvent(new Event("change", { bubbles: true }));
      };
      list.appendChild(it);
    });
  };
  const open = () => { buildList(""); box.classList.add("open"); };

  btn.onclick = () => { isOpen() ? close() : (open(), inp.focus()); };
  inp.onclick = () => { if (!isOpen()) open(); };
  inp.onfocus = () => { if (!isOpen()) open(); };
  inp.onblur = () => setTimeout(() => {
    if (!box.contains(document.activeElement)) close();
  }, 140);
  inp.onkeydown = e => {
    if (e.key === "Escape") { close(); return; }
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (!isOpen()) { open(); return; }
      const items = [...list.querySelectorAll(".combo-opt")];
      if (!items.length) return;
      let i = items.findIndex(x => x.classList.contains("hl"));
      if (i >= 0) items[i].classList.remove("hl");
      i = e.key === "ArrowDown"
        ? (i + 1) % items.length
        : (i - 1 + items.length) % items.length;
      items[i].classList.add("hl");
      items[i].scrollIntoView({ block: "nearest" });
    } else if (e.key === "Enter" && isOpen()) {
      e.preventDefault();
      const hl = list.querySelector(".combo-opt.hl")
              || list.querySelector(".combo-opt.sel");
      if (hl) hl.click(); else close();
    }
  };
  inp.oninput = () => { if (isOpen()) buildList(inp.value); scheduleSetValues(); };
  inp.onchange = () => scheduleSetValues();
  return { box, inp };
}

function renderSelect(c) {
  const row = el("div", "field");
  row.innerHTML = `<label>${escapeHtml(c.label)}</label>`;
  const current = VALUES[c.id] !== undefined ? VALUES[c.id] : c.default;
  let sel;
  if (c.free) {
    /* edytowalna lista rozwijana: wygląd i zachowanie jak select
       (klik rozwija pełną listę), ale można też wpisać własną wartość */
    const dlId = "dl-" + c.id;
    const dl = el("datalist");
    dl.id = dlId;
    for (const v of (c.values || [])) {
      const o = el("option"); o.value = v; dl.appendChild(o);
    }
    row.appendChild(dl);
    const combo = buildCombo(dlId, current);
    sel = combo.inp;
    row.appendChild(combo.box);
  } else {
    sel = el("select");
    for (const v of (c.values || [])) {
      const o = el("option", null, escapeHtml(v));
      o.value = v;
      if (v === current) o.selected = true;
      sel.appendChild(o);
    }
    sel.onchange = scheduleSetValues;
  }
  sel.dataset.cid = c.id; sel.dataset.kind = "select";
  /* dla comboboxa input jest już wewnątrz .combo — nie przenosimy go */
  if (!row.contains(sel)) row.appendChild(sel);
  return row;
}

/* województwo -> powiat -> gmina: odświeża podpowiedzi po zmianie,
   żeby listy nie były zamarznięte na domyślnym województwie */
function wireTerritory() {
  const T = SCHEMA.tpl_territory || {};
  const prefixes = new Set();
  $$("[data-cid]").forEach(n => {
    const m = String(n.dataset.cid || "").match(/^(\w+?)_(woj|powiat|gmina)$/);
    if (m) prefixes.add(m[1]);
  });
  for (const pref of prefixes) {
    const woj = document.querySelector(`[data-cid="${pref}_woj"]`);
    const pow = document.querySelector(`[data-cid="${pref}_powiat"]`);
    const gm  = document.querySelector(`[data-cid="${pref}_gmina"]`);
    if (!woj || !pow || !gm) continue;
    const setDl = (cid, lista) => {
      const dl = document.getElementById("dl-" + cid);
      if (!dl) return;
      dl.innerHTML = "";
      for (const v of (lista || [])) {
        const o = el("option"); o.value = v; dl.appendChild(o);
      }
    };
    const fillPowiat = () => {
      const w = String(woj.value || "").trim().toUpperCase();
      setDl(`${pref}_powiat`, Object.keys(T[w] || {}));
    };
    const fillGmina = () => {
      const w = String(woj.value || "").trim().toUpperCase();
      const p = String(pow.value || "").trim().toUpperCase();
      setDl(`${pref}_gmina`, (T[w] || {})[p] || []);
    };
    woj.addEventListener("change", () => { fillPowiat(); fillGmina(); });
    pow.addEventListener("change", fillGmina);
    fillPowiat();
    fillGmina();
  }
}

/* kontrolka: marginesy */
const MARGIN_TYPES = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "HALIZNY", "WYK_NEG", "OPIS", "ZEST1", "WK_ZM1"];
const MARGIN_SIDES = ["T", "B", "L", "R"];

function renderMargins(c) {
  const wrap = el("div", null);
  const det = el("details", "margins-details");
  const sum = el("summary", null,
    escapeHtml((c.label || "Ustawienia marginesów") + " (zwiń / rozwiń)"));
  det.appendChild(sum);
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
  det.appendChild(table);
  wrap.appendChild(det);
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
/* ------------------------------------- kontrolka: edytor bazy obszarów GDOŚ */
/* kontrolka: baza obszarów GDOŚ — ładuje dane dopiero przy pierwszym
   otwarciu zakładki (przy 40 tys.+ obszarach renderowanie na starcie
   zamrażało cały interfejs); lista stronicowana + z wyszukiwarką */
const GDOS = { wrap: null, rows: [], q: "", page: 0, per: 50, loaded: false };

function renderGdos(c) {
  const box = el("div", "field gdos-field");
  box.innerHTML = `<label>${escapeHtml(c.label || "Obszary ochrony przyrody")}</label>` +
    `<div class="gdos-hint">Nazwa musi być taka sama jak w pliku wynikowym GDOŚ (np. „Ostoja Międzychodzko-Sierakowska”). Publikacja PZO, powiązanie i opis trafiają do opisu ogólnego. W pustych polach pokazuję przykłady.</div>`;
  const wrap = el("div", "gdos-wrap");
  wrap.dataset.lazy = "1";
  wrap.innerHTML = '<div class="gdos-hint">Baza wczyta się po otwarciu tej zakładki.</div>';
  box.appendChild(wrap);
  GDOS.wrap = wrap;
  return box;
}

function gdosLoad() {
  if (!GDOS.wrap) return;
  GDOS.wrap.dataset.lazy = "0";
  GDOS.loaded = false;
  GDOS.wrap.innerHTML = '<div class="gdos-hint">Wczytywanie bazy obszarów… (przy dużej bazie może chwilę potrwać)</div>';
  api().gdos_list().then(r => {
    GDOS.rows = (r && r.ok && Array.isArray(r.rows)) ? r.rows : [];
    GDOS.loaded = true;
    GDOS.q = "";
    GDOS.page = 0;
    gdosRender();
    if (r && !r.ok) toast(r.error || "Nie udało się wczytać bazy obszarów.", "error");
  }).catch(() => {
    GDOS.wrap.innerHTML = '<div class="gdos-hint">Nie udało się wczytać bazy (program jeszcze się uruchamia?).</div>';
  });
}

function gdosFiltered() {
  const q = GDOS.q.trim().toLowerCase();
  if (!q) return GDOS.rows;
  return GDOS.rows.filter(r =>
    (r.nazwa || "").toLowerCase().includes(q) ||
    (r.kod || "").toLowerCase().includes(q) ||
    (r.typ || "").toLowerCase().includes(q));
}

const GDOS_PH = {
  nazwa: "np. Ostoja Międzychodzko-Sierakowska",
  typ: "OSO lub SOO — puste dla parku krajobrazowego",
  kod: "np. PLH300036",
  pzo: "np. Dz. Urz. Woj. Wielkopolskiego z 2014 r. poz. 1793",
  powiazanie: "np. PZO wskazuje zagrożenia związane m.in. z cięciami starodrzewów, pracami w okresach wrażliwych oraz usuwaniem drzew dziuplastych i martwego drewna...",
  opis: "np. Sierakowski Park Krajobrazowy: według stanu na 21 września 2026 r. nie zgłoszono sprzeciwu do zadań gospodarki leśnej na gruntach prywatnych."
};

function gdosCard(row) {
  const nowy = !String(row.nazwa || "").trim();
  const item = el("div", "gdos-item" + (nowy ? "" : " collapsed"));
  const bind = (w, f) => { w.oninput = () => { row[f] = w.value; }; };

  const head = el("div", "gdos-item-head");
  const tog = el("button", "gdos-tog");
  tog.type = "button";
  tog.title = "Rozwiń / zwiń szczegóły obszaru";
  tog.innerHTML = ICON("chevron");

  const nazwa = el("input", "gdos-nazwa");
  nazwa.type = "text";
  nazwa.placeholder = GDOS_PH.nazwa;
  nazwa.value = row.nazwa != null ? String(row.nazwa) : "";
  nazwa.title = "Nazwa obszaru — dokładnie taka jak w pliku GDOŚ";
  bind(nazwa, "nazwa");

  const typ = el("input", "gdos-typ");
  typ.type = "text";
  typ.placeholder = GDOS_PH.typ;
  typ.value = row.typ != null ? String(row.typ) : "";
  typ.title = "Typ obszaru";
  bind(typ, "typ");

  const kod = el("input", "gdos-kod");
  kod.type = "text";
  kod.placeholder = GDOS_PH.kod;
  kod.value = row.kod != null ? String(row.kod) : "";
  kod.title = "Kod obszaru Natura 2000";
  bind(kod, "kod");

  const del = el("button", "btn secondary gdos-del");
  del.innerHTML = ICON("trash");
  del.title = "Usuń ten obszar";
  del.onclick = () => {
    const i = GDOS.rows.indexOf(row);
    if (i >= 0) GDOS.rows.splice(i, 1);
    gdosRender();
  };

  head.appendChild(tog);
  head.appendChild(nazwa);
  head.appendChild(typ);
  head.appendChild(kod);
  head.appendChild(del);
  item.appendChild(head);

  const body = el("div", "gdos-item-body");
  [["pzo", "Publikacja PZO"], ["powiazanie", "Powiązanie z gospodarką leśną"],
   ["opis", "Opis (pozostałe formy)"]].forEach(f => {
    const fld = el("div", "gdos-fld");
    fld.appendChild(el("label", null, escapeHtml(f[1])));
    const ta = el("textarea");
    ta.rows = 3;
    ta.placeholder = GDOS_PH[f[0]];
    ta.value = row[f[0]] != null ? String(row[f[0]]) : "";
    bind(ta, f[0]);
    fld.appendChild(ta);
    body.appendChild(fld);
  });
  item.appendChild(body);
  tog.onclick = () => item.classList.toggle("collapsed");
  return item;
}

function gdosRender() {
  const wrap = GDOS.wrap;
  if (!wrap || !GDOS.loaded) return;
  const fil = gdosFiltered();
  const pages = Math.max(1, Math.ceil(fil.length / GDOS.per));
  if (GDOS.page >= pages) GDOS.page = pages - 1;
  const from = GDOS.page * GDOS.per;
  const slice = fil.slice(from, from + GDOS.per);

  wrap.innerHTML = "";

  /* pasek narzędzi: szukajka + licznik + odśwież */
  const top = el("div", "gdos-toolbar");
  const srch = el("input", "gdos-search");
  srch.type = "text";
  srch.placeholder = "Szukaj po nazwie, kodzie lub typie…";
  srch.value = GDOS.q;
  srch.oninput = () => { GDOS.q = srch.value; GDOS.page = 0; gdosRender(); };
  const info = el("div", "gdos-info");
  info.textContent = GDOS.q.trim()
    ? `Znaleziono ${fil.length} z ${GDOS.rows.length} obszarów`
    : `${GDOS.rows.length} obszarów w bazie`;
  const rl = el("button", "btn ghost small", "Odśwież");
  rl.title = "Pobierz bazę ponownie z dysku (np. po imporcie z Excela)";
  rl.onclick = () => gdosLoad();
  top.appendChild(srch);
  top.appendChild(info);
  top.appendChild(rl);
  wrap.appendChild(top);

  const list = el("div", "gdos-list");
  slice.forEach(row => list.appendChild(gdosCard(row)));
  if (!slice.length) {
    list.appendChild(el("div", "gdos-hint",
      "Brak obszarów pasujących do wyszukiwania."));
  }
  wrap.appendChild(list);

  /* stronicowanie */
  const pager = el("div", "gdos-pager");
  const prev = el("button", "btn secondary small", "‹ Poprzednia");
  prev.disabled = GDOS.page <= 0;
  prev.onclick = () => { GDOS.page--; gdosRender(); };
  const lab = el("span", "gdos-page-info",
    `Strona ${GDOS.page + 1} z ${pages} — pozycje ` +
    `${fil.length ? from + 1 : 0}–${from + slice.length}`);
  const next = el("button", "btn secondary small", "Następna ›");
  next.disabled = GDOS.page >= pages - 1;
  next.onclick = () => { GDOS.page++; gdosRender(); };
  pager.appendChild(prev);
  pager.appendChild(lab);
  pager.appendChild(next);
  wrap.appendChild(pager);

  /* akcje */
  const acts = el("div", "actions");
  const add = el("button", "btn secondary");
  add.innerHTML = ICON("plus") + "<span>Dodaj obszar</span>";
  add.onclick = () => {
    GDOS.rows.push({ nazwa: "", typ: "", kod: "", pzo: "", powiazanie: "", opis: "" });
    GDOS.q = "";
    GDOS.page = Math.floor((GDOS.rows.length - 1) / GDOS.per);
    gdosRender();
    const karty = wrap.querySelectorAll(".gdos-item");
    if (karty.length) karty[karty.length - 1].querySelector(".gdos-nazwa").focus();
  };
  const save = el("button", "btn primary");
  save.innerHTML = ICON("save") + "<span>Zapisz bazę</span>";
  save.onclick = async () => {
    const czyste = GDOS.rows.filter(o => String(o.nazwa || "").trim());
    if (czyste.length !== GDOS.rows.length) toast("Pominięto obszary bez nazwy.", "warn");
    save.disabled = true;
    save.innerHTML = ICON("save") + "<span>Zapisuję…</span>";
    try {
      const r = await api().gdos_save(JSON.stringify(czyste));
      if (r && r.ok) toast("✓ Zapisano bazę obszarów (" + r.count + ")", "done");
      else toast((r && r.error) || "Nie udało się zapisać bazy.", "error");
    } catch (e) {
      toast("Nie udało się zapisać bazy.", "error");
    }
    save.disabled = false;
    save.innerHTML = ICON("save") + "<span>Zapisz bazę</span>";
  };
  acts.appendChild(add);
  acts.appendChild(save);
  wrap.appendChild(acts);
}

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
  LAST_STATUS_ERR = false;
  STOP_REQUESTED = false;
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
    case "gdos_changed":
      /* baza została nadpisana (import z Excela) — odśwież edytor,
         żeby nie nadpisał nowej bazy starymi danymi z pamięci */
      if (GDOS.wrap && GDOS.wrap.dataset.lazy === "0") gdosLoad();
      break;
    case "clear_log": $("#log").innerHTML = ""; break;
    case "status":
      $("#status-text").textContent = ev.text;
      LAST_STATUS_ERR = (ev.color === "#D83B01");
      $("#status-dot").className = LAST_STATUS_ERR ? "err" : "";
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
    case "changelog": showChangelog(ev); break;
    case "state": setRunning(!!ev.running); break;
    case "toast": toast(ev.text, ev.kind); break;
  }
}

function appendLog(text) {
  const log = $("#log");
  const div = el("div", null, escapeHtml(text));
  if (/BŁĄD|Error|Traceback/i.test(text)) div.className = "err";
  else if (/\[UWAGA\]|UWAGA:|ostrzeżenie|warn/i.test(text)) div.className = "warn";
  else if (/ZAKOŃCZONO|pomyślnie|\[OK\]/i.test(text)) div.className = "ok";
  log.appendChild(div);
  while (log.children.length > 3000) log.removeChild(log.firstChild);
  log.scrollTop = log.scrollHeight;
}

let LAST_TASK_LABEL = null;    /* nazwa ostatnio uruchomionego zadania */
let LAST_STATUS_ERR = false;   /* czy ostatni status był błędem (czerwony) */
let STOP_REQUESTED = false;   /* czy użytkownik kliknął „Zatrzymaj" */

/* ------------------------------------------------- licznik czasu zadania */
let TIMER_T0 = null, TIMER_IV = null;
function fmtDur(ms) {
  const total = Math.floor(ms / 1000);
  const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), ss = total % 60;
  const p = n => String(n).padStart(2, "0");
  return (h ? h + ":" + p(m) : p(m)) + ":" + p(ss);
}
function startTimer() {
  if (TIMER_IV) { clearInterval(TIMER_IV); TIMER_IV = null; }
  TIMER_T0 = Date.now();
  const t = $("#task-timer");
  t.classList.remove("hidden");
  t.textContent = "⏱ 0:00";
  TIMER_IV = setInterval(() => {
    t.textContent = "⏱ " + fmtDur(Date.now() - TIMER_T0);
  }, 1000);
}
function stopTimerKeep() {
  if (TIMER_IV) { clearInterval(TIMER_IV); TIMER_IV = null; }
  if (TIMER_T0 != null) $("#task-timer").textContent = "⏱ " + fmtDur(Date.now() - TIMER_T0);
  TIMER_T0 = null;
}

function setRunning(running) {
  const was = RUNNING;
  RUNNING = running;
  $$(".run-btn").forEach(b => b.disabled = running);
  $("#btn-stop").classList.toggle("hidden", !running);
  const dot = $("#status-dot");
  if (running) { dot.className = "busy"; startTimer(); }
  else { dot.className = ""; if (was) stopTimerKeep(); }
  /* koniec zadania → powiadomienie + dźwięk */
  if (was && !running) notifyTaskEnd();
}

/* ------------------------------------------------ powiadomienie o zakończeniu */
function notifyTaskEnd() {
  const label = LAST_TASK_LABEL;
  LAST_TASK_LABEL = null;
  if (!label) return;
  if (STOP_REQUESTED) { STOP_REQUESTED = false; return; }
  if (LAST_STATUS_ERR) {
    toast("✗ Zadanie zakończone błędem: " + label + " — szczegóły w logu", "error");
    return;
  }
  const dur = ($("#task-timer") || {}).textContent || "";
  toast("✓ Zakończono: " + label + (dur ? " — czas: " + dur.replace("⏱ ", "") : ""), "done", {
    label: "Otwórz folder wyników",
    ms: 12000,
    run: async () => {
      try {
        const r = await api().open_last_output();
        if (!r.ok) toast(r.error || "Nie udało się otworzyć folderu.", "warn");
      } catch (e) { toast("Nie udało się otworzyć folderu.", "warn"); }
    }
  });
  playChime();
  /* gdy okno jest schowane — spróbuj systemowego powiadomienia (jeśli zgoda) */
  if (document.hidden && typeof Notification !== "undefined" &&
      Notification.permission === "granted") {
    try { new Notification("Forestly — zakończono", { body: label }); } catch (e) {}
  }
}

let _chimeCtx = null;
function playChime() {
  try {
    if (!_chimeCtx) _chimeCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (_chimeCtx.state === "suspended") _chimeCtx.resume();
    const notes = [[880, 0], [1174.66, 0.18]];   /* A5 → D6 */
    for (const [freq, delay] of notes) {
      const o = _chimeCtx.createOscillator();
      const g = _chimeCtx.createGain();
      o.type = "sine";
      o.frequency.value = freq;
      o.connect(g); g.connect(_chimeCtx.destination);
      const t0 = _chimeCtx.currentTime + delay;
      g.gain.setValueAtTime(0.0001, t0);
      g.gain.exponentialRampToValueAtTime(0.22, t0 + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t0 + 0.55);
      o.start(t0); o.stop(t0 + 0.6);
    }
  } catch (e) { /* dźwięk niedostępny — powiadomienie i tak się pokaże */ }
}

/* ------------------------------------------- „Co nowego": czyszczenie treści */
function cleanChangelogText(s) {
  s = String(s || "");
  /* encje HTML wkradające się do changelogu (m.in. przy kopiowaniu z podglądu) */
  const A = "&";
  const pairs = [
    [A + "amp;#x20;", " "],
    [A + "amp;nbsp;", " "],
    [A + "#x20;", " "],
    [A + "nbsp;", " "],
    [A + "#160;", " "],
    [A + "#xa0;", " "],
    [A + "quot;", '"'],
    [A + "#39;", "'"],
    [A + "lt;", "<"],
    [A + "gt;", ">"],
    [A + "amp;", A]
  ];
  for (const [from, to] of pairs) s = s.split(from).join(to);
  /* drugi przebieg — na wypadek podwójnie zaszyfrowanych encji */
  for (const [from, to] of pairs) s = s.split(from).join(to);
  /* znaczniki markdown (pogrubienie/kursywa) — tekst zostaje, gwiazdki znikają */
  s = s.replace(/\*\*([^*]+)\*\*/g, "$1").replace(/\*([^*]+)\*/g, "$1");
  /* ukośniki ucieczki markdown (\-, \[OK\], \> itp.) — zostaje sam znak */
  s = s.replace(/\\([-*_\\[\]()#<>~|`])/g, "$1");
  return s;
}

function renderChangelogBody(box, text) {
  const lines = cleanChangelogText(text).split(/\r?\n/);
  for (const raw of lines) {
    const line = raw.replace(/\s+$/, "");
    if (!line.trim()) { box.appendChild(el("div", "cl-gap", "")); continue; }
    let m;
    if ((m = line.match(/^#{1,6}\s+(.*)$/))) {
      box.appendChild(el("div", "cl-head", escapeHtml(m[1])));
    } else if ((m = line.match(/^[-*•·]\s*(.*)$/))) {
      const row = el("div", "cl-item");
      row.appendChild(el("span", "cl-bullet", "•"));
      row.appendChild(el("span", null, escapeHtml(m[1])));
      box.appendChild(row);
    } else {
      box.appendChild(el("div", "cl-line", escapeHtml(line)));
    }
  }
}

/* -------------------------------------------------------------------- dialogi */
function showChangelog(ev) {
  const wrap = el("div", "modal-backdrop");
  const modal = el("div", "modal");
  modal.style.width = "620px";
  modal.appendChild(el("h3", null, "Co nowego w " + (ev.version || "")));
  const body = el("div", "changelog-body");
  renderChangelogBody(body, ev.body || "");
  modal.appendChild(body);
  const row = el("div", "modal-row");
  const close = el("button", "btn primary");
  close.textContent = "Zamknij";
  close.onclick = () => wrap.remove();
  row.appendChild(close);
  modal.appendChild(row);
  wrap.appendChild(modal);
  wrap.onclick = e => { if (e.target === wrap) wrap.remove(); };
  $("#modal-root").appendChild(wrap);
}

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

function toast(msg, kind, action) {
  const t = el("div", "toast " + (kind || ""), escapeHtml(msg));
  if (action && action.label) {
    const b = el("button", "toast-btn");
    b.type = "button";
    b.textContent = action.label;
    b.onclick = async () => {
      const r = action.run ? await action.run() : null;
      t.remove();
      return r;
    };
    t.appendChild(b);
  }
  $("#toast-root").appendChild(t);
  setTimeout(() => t.remove(), (action && action.ms) || 6000);
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
      const aliasTxt = (tpl && tpl.aliases && tpl.aliases.length)
        ? ` <span class="oi-aliases">(${escapeHtml(tpl.aliases.join(", "))})</span>`
        : "";
      it.innerHTML = `<span class="grip">☰</span>` +
        `<span class="oi-name">${i + 1}. ${escapeHtml(tpl ? tpl.label : key)}` +
        aliasTxt +
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
    STOP_REQUESTED = true;
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
