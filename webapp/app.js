/* Forestly — logika frontendu (most pywebview <-> Python) */
"use strict";

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

let SCHEMA = null;
let VALUES = {};
let RUNNING = false;
let LAST_STATUS_TEXT = "";
let activeTab = null;
let setValuesTimer = null;

/* Pokazywanie kontrolki warunkowej: id kontrolki -> id checkboxa */
/* drzewko do ekranu postępu kreatora — animowany GIF
   (gołe drzewko -> zielone liście -> złota jesień -> opadają, w pętli) */
const BRANCH_HTML =
  '<img src="drzewko.gif" alt="" draggable="false">';

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
  wireOpisOgExclusive();
}

/* „Pełne opisy ogólne" <-> „Skrócone opisy ogólne" — wzajemne wykluczanie
   (oba odznaczone = opisy ogólne pomijane); folder GDOŚ potrzebny w obu
   wariantach, więc pokazujemy go przy każdym z nich */
function wireOpisOgExclusive() {
  /* kreator 1-Click + zakładka Opisy ogólne — ta sama para przełączników */
  const pary = [["all_pelny_opis_og", "all_krotki_opis_og", "all_gdos"],
                ["opis_og_pelny", "opis_og_krotki", "opis_og_gdos"]];
  for (const [pid, kid, gid] of pary) {
  const pelny = document.querySelector(`[data-cid="${pid}"]`);
  const krotki = document.querySelector(`[data-cid="${kid}"]`);
  const gdos = document.querySelector(`[data-cid="${gid}"]`);
  if (!pelny || !krotki) continue;
  const odswiez = () => {
    if (gdos) {
      const row = gdos.closest(".field") || gdos;
      row.classList.toggle("hidden", !pelny.checked && !krotki.checked);
    }
    scheduleSetValues();
  };
  pelny.addEventListener("change", () => {
    if (pelny.checked) krotki.checked = false;
    odswiez();
  });
  krotki.addEventListener("change", () => {
    if (krotki.checked) pelny.checked = false;
    odswiez();
  });
  odswiez();
  }
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
  /* wejście w Pełny Automat (także z ekranu Start) otwiera kreatora */
  if (key === allTabKey()) openWizard();
  const active = $$(".nav-item").find(i => i.dataset.key === key);
  if (active) {
    const g = active.closest(".nav-group");
    if (g) g.classList.remove("collapsed");
    active.scrollIntoView({ block: "nearest" });
  }  mpMaybeClose();
}

function renderOneControl(c) {
  switch (c.kind) {
    case "path": return renderPath(c);
    case "text": return renderText(c);
    case "check": return renderCheck(c);
    case "checks": return renderChecks(c);
    case "select": return renderSelect(c);
    case "margins": return renderMargins(c);
    case "czcionki": return renderCzcionki(c);
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
  /* Pełny Automat (1-Click): klasyczny formularz chowam — zakładka
     otwiera kreatora krok po kroku (wizard), który korzysta z tych pól */
  const isAll = tab.key.indexOf("Pełny Automat") >= 0;
  let target = view;
  if (isAll) {
    const hide = document.createElement("div");
    hide.style.display = "none";
    view.appendChild(hide);
    target = hide;
  }
  for (const c of tab.controls) {
    const node = renderOneControl(c);
    if (node) target.appendChild(node);
  }
  if (isAll) {
    WIZ.home = target;
    WIZ.children = [...target.childNodes];
    return;  /* przyciski akcji ma kreator */
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

/* ================== Kreator Pełnego Automatu (1-Click) ==================
   Zakładka 1-Click otwiera okno kreatora: krótki opis → lokalizacje →
   strona tytułowa i daty → marginesy → nazwiska → uruchomienie → postęp.
   Kreator korzysta z PRAWDZIWYCH pól formularza (przenosi je do okna),
   więc zapisywanie i odczyt ustawień działa bez zmian. */
const WIZ = { open: false, step: 0, home: null, children: [] };
const WIZ_NAV = ["Lokalizacje", "Strona tytułowa i daty", "Marginesy",
                 "Nazwiska", "Uruchomienie"];

function allTabKey() {
  return (SCHEMA.tabs.find(t => t.key.indexOf("Pełny Automat") >= 0) || {}).key || "";
}

function wizField(cid) {
  const n = WIZ.home && WIZ.home.querySelector('[data-cid="' + cssEscape(cid) + '"]');
  /* cała kontrolka — dla checków label.check-row (gdyby przenieść sam input,
     straciłby etykietę i ginął przy kolejnym renderWizStep) */
  return n ? (n.closest(".field, .check-row") || n) : null;
}

function wizVal(cid) {
  const n = WIZ.home && WIZ.home.querySelector('[data-cid="' + cssEscape(cid) + '"]');
  if (!n) return "";
  if (n.type === "checkbox") return n.checked;
  return String(n.value || "").trim();
}

function openWizard() {
  if (WIZ.open || !allTabKey() || !WIZ.home) return;
  WIZ.open = true;
  WIZ.step = 0;
  const wz = document.createElement("div");
  wz.id = "wizard";
  wz.innerHTML =
    '<div class="wiz-box">' +
      '<div class="wiz-head">' +
        '<div class="wiz-title">Pełny Automat (1-Click)</div>' +
        '<div class="wiz-dots" id="wiz-dots"></div>' +
        '<button class="wiz-close" id="wiz-close" title="Zamknij kreatora">×</button>' +
      '</div>' +
      '<div class="wiz-body" id="wiz-body"></div>' +
    '</div>';
  document.body.appendChild(wz);
  wz.querySelector("#wiz-close").onclick = () => closeWizard();
  renderWizStep();
}

function closeWizard() {
  const wz = document.getElementById("wizard");
  if (wz) wz.remove();
  /* okno podglądu (systemowe) żyje własnym życiem — zamykamy je razem
     z kreatorem, żeby nie zostało na ekranie samo */
  try { api().close_preview_window(); } catch (e) { /* okno już zamknięte */ }
  /* kontrolki wracają na swoje miejsce (w oryginalnej kolejności) */
  if (WIZ.home) WIZ.children.forEach(n => WIZ.home.appendChild(n));
  WIZ.open = false;
  showTab(START_KEY);
}

function renderWizStep() {
  if (!WIZ.open) return;
  /* opuszczamy krok marginesów/czcionek (3) — okno podglądu ma się zamknąć,
     żeby nie wisiało nad kolejnymi krokami kreatora */
  if (WIZ.lastStep != null && WIZ.lastStep === 3 && WIZ.step !== 3) {
    try { api().close_preview_window(); } catch (e) { /* już zamknięte */ }
  }
  WIZ.lastStep = WIZ.step;
  const body = document.getElementById("wiz-body");
  const dots = document.getElementById("wiz-dots");
  if (!body) return;
  /* wszystkie pola wracają najpierw do ukrytego formularza — dzięki temu
     przechodzenie Wstecz/Dalej nie gubi kontrolek, a odczyty wartości,
     zapis ustawień i collectValues działają zawsze */
  if (WIZ.home) WIZ.children.forEach(n => WIZ.home.appendChild(n));
  dots.innerHTML = "";
  WIZ_NAV.forEach((n, i) => {
    const d = document.createElement("div");
    d.className = "wiz-dot" + (i === WIZ.step ? " on" : (i < WIZ.step ? " done" : ""));
    d.title = n;
    dots.appendChild(d);
  });
  /* kroki (kółka) chowamy na czas trwania całego procesu — nie mają
     sensu, gdy trwa generowanie (ekran postępu) */
  dots.classList.toggle("hidden", WIZ.step === 6);
  body.innerHTML = "";
  const st = document.createElement("div");
  st.className = "wiz-step";
  body.appendChild(st);

  const back = el("button", "btn secondary", "‹ Wstecz");
  back.onclick = () => { if (WIZ.step > 0 && WIZ.step < 6) { WIZ.step--; renderWizStep(); mpMaybeClose(); } };
  const next = el("button", "btn primary wiz-big", "Dalej ›");
  const nav = document.createElement("div");
  nav.className = "wiz-nav";

  const moveTo = (f) => { if (f) st.appendChild(f); };

  if (WIZ.step === 0) {
    st.innerHTML =
      '<div class="wiz-hero">' +
        '<div class="wh-ic">' + (SECTION_ICONS["MIETEK"] || "🚜") + '</div>' +
        '<h2>Cały proces jednym kliknięciem</h2>' +
        '<div class="wiz-sub">Halizny → TXT z DBF → pliki Word (ze stronami tytułowymi' +
        ' i opisami ogólnymi) → PDF → scalenie w jeden dokument.<br>' +
        'Przeprowadzę Cię przez cztery krótkie kroki — potem zrobię wszystko sama.</div>' +
      '</div>';
    /* przełącznik nowych szablonów (HTML → PDF bez Worda) */
    const nsw = el("label", "check-row wiz-newtpl");
    nsw.innerHTML = '<input type="checkbox" id="wiz-nowe-szablony"><span>' +
      '<b>Nowe szablony wydruków</b><br>' +
      '<span class="opis">Włączony: szybsze generowanie PDF bez Worda i nowy,' +
      ' czytelniejszy wygląd. Wyłączony: wszystko działa jak dotychczas.</span></span>';
    const nsInp = nsw.querySelector("input");
    nsInp.checked = (localStorage.getItem("forestly_nowe_szablony") ?? "1") !== "0";
    nsInp.addEventListener("change", () => {
      localStorage.setItem("forestly_nowe_szablony", nsInp.checked ? "1" : "0");
      scheduleSetValues();
    });
    const heroBox = st.querySelector(".wiz-hero");
    if (heroBox) heroBox.appendChild(nsw);
    /* synchronizacja stanu z backendem od razu (żeby default też doszedł) */
    scheduleSetValues();
    back.classList.add("hidden");
    next.innerHTML = "<span>Zaczynamy ›</span>";
    next.onclick = () => { WIZ.step = 1; renderWizStep(); };
  } else if (WIZ.step === 1) {
    st.innerHTML = '<h2>Gdzie są mietki i gdzie zapisać wyniki?</h2>' +
      '<div class="wiz-sub">Wskaż folder z danymi źródłowymi (obręby z plikami DBF)' +
      ' oraz folder docelowy, w którym powstanie cała dokumentacja.</div>';
    moveTo(wizField("all_src"));
    moveTo(wizField("all_dst"));
    next.onclick = () => { WIZ.step = 2; renderWizStep(); };
  } else if (WIZ.step === 2) {
    st.innerHTML = '<h2>Strona tytułowa i daty</h2>' +
      '<div class="wiz-sub">Z tych danych powstanie strona tytułowa każdej wsi.' +
      ' „Stan na” zastępuje daty we wszystkich dokumentach Word,' +
      ' a pola 10-lecia — okres w WSK_ZB.</div>';
    const grp = WIZ.home.querySelector("details.group-details");
    if (grp) {
      grp.open = true;
      moveTo(grp.parentElement);
    }
    moveTo(wizField("all_pelny_opis_og"));
    moveTo(wizField("all_krotki_opis_og"));
    moveTo(wizField("all_gdos"));
    moveTo(wizField("all_custom_skroty"));
    moveTo(wizField("all_skroty"));
    next.onclick = () => { WIZ.step = 3; renderWizStep(); };
  } else if (WIZ.step === 3) {
    st.innerHTML = '<h2>Marginesy wydruków (cm)</h2>' +
      '<div class="wiz-sub">Odstępy od krawędzi strony dla poszczególnych typów' +
      ' wydruków. Rozwiń listę, jeśli chcesz coś zmienić.</div>';
    const md = WIZ.home.querySelector("details.margins-details");
    if (md) {
      md.open = true;
      moveTo(md.parentElement);
    }
    const cd = WIZ.home.querySelector("details.czcionki-details");
    if (cd) moveTo(cd.parentElement);
    next.onclick = () => { WIZ.step = 4; renderWizStep(); };
  } else if (WIZ.step === 4) {
    /* trzy warianty: z nazwiskami / bez nazwisk / obie wersje (dwa foldery) */
    const big = document.createElement("div");
    big.className = "wiz-check-big";
    big.innerHTML = '<h2>Nazwiska w REJESTRZE</h2>' +
      '<div class="wiz-sub">Wybierz, która wersja wydruków ma powstać. ' +
      '"Obie wersje" uruchamia proces dwukrotnie i tworzy dwa foldery wynikowe.</div>';
    const f = wizField("remove_names");
    const fo = wizField("all_obie_wersje");
    const inpR = f ? f.querySelector('input[type="checkbox"]') : null;
    const inpO = fo ? fo.querySelector('input[type="checkbox"]') : null;
    const kafle = el("div", "wiz-tiles");
    const opcje = [
      { txt: "Z nazwiskami",
        sub: "REJESTR z pełnymi nazwiskami właścicieli",
        set: () => { if (inpO) inpO.checked = false;
                     if (inpR) inpR.checked = false; } },
      { txt: "Bez nazwisk",
        sub: "nazwiska usunięte z REJESTRU (oraz z 1. strony)",
        set: () => { if (inpO) inpO.checked = false;
                     if (inpR) inpR.checked = true; } },
      { txt: "Obie wersje",
        sub: "dwa foldery: 'Z nazwiskami' i 'Bez nazwisk'",
        set: () => { if (inpO) inpO.checked = true;
                     if (inpR) inpR.checked = true; } },
    ];
    const odswiezKafle = () => {
      const obie = !!(inpO && inpO.checked);
      const bez = !!(inpR && inpR.checked);
      const wybrana = obie ? 2 : (bez ? 1 : 0);
      Array.from(kafle.children).forEach((k, i) =>
        k.classList.toggle("sel", i === wybrana));
    };
    opcje.forEach(o => {
      const k = el("div", "wiz-tile");
      k.innerHTML = '<div class="wiz-tile-t">' + escapeHtml(o.txt) + '</div>' +
                    '<div class="wiz-tile-s">' + escapeHtml(o.sub) + '</div>';
      k.onclick = () => {
        o.set();
        if (inpR) inpR.dispatchEvent(new Event("change", { bubbles: true }));
        if (inpO) inpO.dispatchEvent(new Event("change", { bubbles: true }));
        odswiezKafle();
      };
      kafle.appendChild(k);
    });
    if (inpR) inpR.addEventListener("change", odswiezKafle);
    if (inpO) inpO.addEventListener("change", odswiezKafle);
    big.appendChild(kafle);
    odswiezKafle();
    st.appendChild(big);

    /* mapa (opcjonalnie): wskaż plik albo przeciągnij i upuść */
    const karta = el("div", "wiz-map-card");
    karta.innerHTML = '<h2>Mapy (opcjonalnie)</h2>' +
      '<div class="wiz-sub">Wskaż folder z mapami (jpg, png, tiff) albo przeciągnij je ' +
      'poniżej — można wiele naraz. Każdą mapę dopasuję do wsi PO NAZWIE ' +
      '(nazwa pliku ma zawierać nazwę wsi, np. "CHORZEWO mapa.jpg") i dołączę ' +
      'jako PDF na końcu pakietu.</div>';
    const fm = wizField("all_mapa");
    if (fm) {
      fm.classList.add("wiz-map-field");
      karta.appendChild(fm);
    }
    const dz = el("div", "wiz-drop");
    dz.innerHTML = '<b>Przeciągnij i upuść mapy tutaj</b>' +
                   '<span>jpg · png · tiff · można wiele naraz · dopasuję po nazwie wsi</span>';
    const czysc = el("button", "btn ghost small wiz-map-clear",
                     "Wyczyść przeciągnięte mapy");
    czysc.type = "button";
    czysc.onclick = async () => {
      try {
        await api().mapa_drop_clear();
        toast("Zapomniano przeciągnięte mapy.", "ok");
      } catch (e) { toast("Nie udało się wyczyścić.", "warn"); }
    };
    karta.appendChild(czysc);
    dz.addEventListener("dragover", e => {
      e.preventDefault();
      dz.classList.add("over");
    });
    dz.addEventListener("dragleave", () => dz.classList.remove("over"));
    let lastDropFolder = "";
    dz.addEventListener("drop", async e => {
      e.preventDefault();
      dz.classList.remove("over");
      const pliki = e.dataTransfer.files
        ? Array.from(e.dataTransfer.files)
        : [];
      if (!pliki.length) return;
      const grafiki = pliki.filter(f => /\.(jpe?g|png|tiff?)$/i.test(f.name));
      const odrzucone = pliki.length - grafiki.length;
      if (odrzucone)
        toast("Pominięto " + odrzucone + " plik(ów) — mapy to jpg / png / tiff.", "warn");
      if (!grafiki.length) return;
      dz.classList.add("busy");
      // najpierw czytamy WSZYSTKIE pliki (odczyty ruszają natychmiast,
      // równolegle — nie wolno ich odkładać na po await-ach, bo po
      // zakończeniu zdarzenia drop niektóre silniki blokują dostęp),
      // dopiero potem wysyłamy do programu
      let odczyty = [];
      try {
        odczyty = await Promise.all(grafiki.map(async plik => {
          const buf = new Uint8Array(await plik.arrayBuffer());
          let b64 = "";
          const K = 32768;
          for (let i = 0; i < buf.length; i += K)
            b64 += String.fromCharCode.apply(null, buf.subarray(i, i + K));
          return [plik.name, btoa(b64)];
        }));
      } catch (err) {
        toast("Nie udało się wczytać plików: " + err, "warn");
      }
      let ok = 0, folder = "";
      for (const [nazwa, b64] of odczyty) {
        try {
          const r = await api().save_mapa_drop(nazwa, b64);
          if (r && r.ok) {
            ok++;
            folder = r.folder || "";
          } else {
            toast(nazwa + ": " + ((r && r.error) || "nie udało się zapisać"), "warn");
          }
        } catch (err) {
          toast(nazwa + ": nie udało się zapisać (" + err + ")", "warn");
        }
      }
      dz.classList.remove("busy");
      if (ok) {
        const inp = fm ? fm.querySelector("input") : null;
        if (inp && folder && (!inp.value || inp.value === lastDropFolder)) {
          inp.value = folder;      // pokaż folder z przeciągniętymi mapami
          inp.dispatchEvent(new Event("input", { bubbles: true }));
        }
        lastDropFolder = folder || lastDropFolder;
        toast("Zapisano " + ok + " map(y) — dopasuję po nazwie wsi.", "ok");
      }
    });
    karta.appendChild(dz);
    st.appendChild(karta);
    next.onclick = () => { WIZ.step = 5; renderWizStep(); };
  } else if (WIZ.step === 5) {
    st.innerHTML = '<h2>Wszystko gotowe!</h2>' +
      '<div class="wiz-sub">Tak uruchomię proces — jeszcze możesz coś zmienić,' +
      ' wracając do poprzednich kroków.</div>';
    const sum = document.createElement("div");
    sum.className = "wiz-summary";
    const rows = [
      ["Mietki (źródło)", wizVal("all_src") || "— nie wskazano —"],
      ["Folder wyników", wizVal("all_dst") || "— nie wskazano —"],
      ["Obszar", [wizVal("all_tpl_gmina"), wizVal("all_tpl_powiat"),
                  wizVal("all_tpl_woj")].filter(Boolean).join(" / ") || "—"],
      ["Stan na", wizVal("all_tpl_stan") || "—"],
      ["Okres 10-lecia WSK_ZB", (String(wizVal("all_wsk_od")) + " – " +
                  wizVal("all_wsk_do")).replace(/^ – $|^ – | – $/g, "").trim() || "—"],
      ["Nazwiska w REJESTRZE", wizVal("all_obie_wersje")
        ? "OBIE WERSJE — dwa foldery ('Z nazwiskami' i 'Bez nazwisk')"
        : (wizVal("remove_names")
           ? "usuwane z REJESTRU" : "REJESTR z pełnymi nazwiskami")],
      ["Mapy", wizVal("all_mapa") || "— bez map —"],
      ["Własne skróty i symbole", wizVal("all_custom_skroty")
        ? (wizVal("all_skroty") || "(nie wskazano pliku)") : "domyślne z programu"],
      ["Opisy ogólne", wizVal("all_pelny_opis_og")
        ? (wizVal("all_gdos") ? "pełne, z formami ochrony z GDOŚ"
                              : "pełne, bez folderu GDOŚ")
        : wizVal("all_krotki_opis_og")
          ? (wizVal("all_gdos") ? "skrócone, lista form z GDOŚ"
                                : "skrócone, bez folderu GDOŚ")
          : "pomijane"],
    ];
    rows.forEach(r => {
      sum.appendChild(el("div", "wiz-row",
        '<div class="wiz-row-k">' + escapeHtml(r[0]) + '</div>' +
        '<div class="wiz-row-v">' + escapeHtml(r[1]) + '</div>'));
    });
    st.appendChild(sum);
    const extra = el("button", "btn ghost", "Skonfiguruj układ PDF…");
    extra.onclick = () => openOrderDialog("ALL");
    nav.appendChild(extra);
    next.innerHTML = ICON("play") + "<span>Generuj dokumenty</span>";
    next.classList.add("wiz-run");
    next.onclick = async () => {
      WIZ.step = 6;
      renderWizStep();
      await runTask("start_pipeline:ALL");
      /* jeśli backend nie wystartował (np. brak ścieżek) — nie czekaj w nieskończoność */
      setTimeout(() => { if (!RUNNING && WIZ.open && WIZ.step === 6
                           && !document.getElementById("wiz-done")) wizDone(false); }, 900);
    };
  } else if (WIZ.step === 6) {
    st.innerHTML =
      '<div class="wiz-prog">' +
        '<div class="wiz-branch" id="wiz-spin"></div>' +
        '<div class="wiz-op" id="wiz-op">Rozpoczynam…</div>' +
        '<div class="wiz-bar"><div class="wiz-fill" id="wiz-fill"></div></div>' +
        '<div class="wiz-file" id="wiz-file"></div>' +
      '</div>';
    const br = document.getElementById("wiz-spin");
    if (br) br.innerHTML = BRANCH_HTML;
    const stopBtn = el("button", "btn secondary wiz-stop", "Przerwij zadanie");
    stopBtn.id = "wiz-stop";
    stopBtn.onclick = async () => {
      try { await api().stop(); } catch (e) {}
      STOP_REQUESTED = true;
      stopBtn.disabled = true;
      stopBtn.textContent = "Przerywanie…";
      toast("Zatrzymywanie — program zakończy po bieżącym kroku…", "warn");
    };
    /* dolny prawy róg okienka kreatora — wstawiamy do pasku nawigacji
       (lokalna zmienna nav: w momencie renderu nie jest jeszcze w DOM) */
    nav.appendChild(stopBtn);
    const dash = WIZ.home.querySelector("#dashboard");
    if (dash) st.appendChild(dash.closest(".card"));
    back.classList.add("hidden");
    next.classList.add("hidden");
  }

  nav.appendChild(back);
  nav.appendChild(next);
  st.appendChild(nav);
  /* podgląd marginesów żyje tylko przy widocznej tabeli marginesów —
     zmiana kroku kreatora go zamyka */
  mpMaybeClose();
}

/* koniec zadania w kroku postępu — ekran "Ukończono" */
function wizDone(ok) {
  if (!WIZ.open || WIZ.step !== 6) return;
  const stopBtn = document.getElementById("wiz-stop");
  if (stopBtn) stopBtn.remove();
  const przerwano = (LAST_STATUS_TEXT || "").indexOf("Przerwano") === 0;
  const spin = document.getElementById("wiz-spin");
  if (spin) {
    spin.id = "wiz-done";
    spin.className = "wiz-done" + (ok && !przerwano ? "" : " err");
    spin.textContent = przerwano ? "⏹" : (ok ? "✓" : "✗");
  }
  const op = document.getElementById("wiz-op");
  if (op) op.textContent = przerwano
    ? "Przerwano przez użytkownika"
    : (ok ? "Ukończono!"
         : "Zadanie zakończone z błędem — szczegóły w dzienniku zdarzeń");
  const fill = document.getElementById("wiz-fill");
  if (fill && ok) fill.style.width = "100%";
  const nav = document.querySelector("#wiz-body .wiz-nav");
  if (nav) {
    nav.querySelectorAll("button").forEach(b => b.remove());
    const openBtn = el("button", "btn secondary", "Otwórz folder wyników");
    openBtn.onclick = async () => {
      try {
        const r = await api().open_last_output();
        if (r && r.ok === false) toast(r.error || "Nie udało się otworzyć folderu.", "warn");
      } catch (e) { toast("Nie udało się otworzyć folderu.", "warn"); }
    };
    const closeBtn = el("button", "btn primary wiz-big", "Zamknij");
    closeBtn.onclick = () => closeWizard();
    nav.appendChild(openBtn);
    nav.appendChild(closeBtn);
  }
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
  input.dataset.attr = c.attr || c.id;
  input.checked = !!VALUES[c.id];
  /* jeśli to już DRUGA kontrolka na tej samej zmiennej (np. word_remove_names
     na remove_names_var) — przejmij stan pierwszej, żeby na starcie i przy
     każdej akcji obie trzymały tę samą wartość */
  {
    const inne = $$('input[type="checkbox"][data-attr="' + input.dataset.attr + '"]')
      .filter(o => o !== input);
    if (inne.length) input.checked = inne[0].checked;
  }
  input.onchange = () => {
    /* ta sama wartość może być współdzielona przez kilka kontrolek
       (np. remove_names_var w 1-Click oraz w zakładce Word) — bez tej
       synchronizacji przestarzała kopia z drugiej zakładki nadpisywała
       wybór i nazwiska zostawały w REJESTRZE */
    if (input.dataset.attr) {
      $$('input[type="checkbox"][data-attr="' + input.dataset.attr + '"]')
        .forEach(o => { if (o !== input) o.checked = input.checked; });
    }
    applyDeps();
    scheduleSetValues();
  };
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
    input.onchange = () => onGroupChange(c, grid, input);
    row.appendChild(input);
    row.appendChild(el("span", null, escapeHtml(choice)));
    grid.appendChild(row);
  }
  applyGroupStates(c, grid);
  if (c.tooltip) grid.title = c.tooltip;
  wrap.appendChild(grid);
  return wrap;
}

function onGroupChange(c, grid, changedInput) {
  /* changedInput — input, na którym odpalił onchange (a nie document.activeElement,
     bo w PyWebView focus nie zawsze trafia w checkbox); NodeList nie ma .some,
     więc stary zapis 'list.some ? ... : false' zawsze zwracał false i po zaznaczeniu
     np. HALIZNY opcja 'Wszystkie' wracała sama z powrotem */
  const all = grid.querySelector('[data-choice="Wszystkie"] input');
  const changed = changedInput && changedInput.dataset.choice
    ? changedInput.dataset.choice : null;
  if (changed === "Wszystkie" && all.checked) {
    grid.querySelectorAll("input").forEach(i => { if (i.dataset.choice !== "Wszystkie") i.checked = false; });
  } else if (changed !== "Wszystkie") {
    if (all.checked) all.checked = false;
    if (AUTO_RETURN_ALL[c.id]) {
      const any = Array.from(grid.querySelectorAll("input"))
        .some(i => i.checked && i.dataset.choice !== "Wszystkie");
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

/* ---------------------------------------------- podgląd marginesów (na żywo) */
const MP_TYPY = ["OPTAX", "REJESTR1", "TAB_KLW3", "WSKAZ1", "WSK_ZB",
                 "ZEST1", "HALIZNY", "WYK_NEG"];
/* typ raportu -> wiersz tabeli marginesów (WSK_ZB drukuje się na marginesach Opisu) */
const MP_WIERSZ = { WSK_ZB: "OPIS" };
let MP_CLEANUP = null;
let MP_CID = null;          /* cid tabeli marginesów, którą podgląd słucha */

function closeMarginsPreview() {
  const d = document.getElementById("mp-dock");
  if (d) d.remove();
  if (MP_CLEANUP) { MP_CLEANUP(); MP_CLEANUP = null; }
  MP_CID = null;
}

/* zamyka podgląd, gdy tabela marginesów, której słucha, zniknęła z widoku
   (zmiana kroku kreatora albo przejście do innej zakładki) */
function mpMaybeClose() {
  if (!document.getElementById("mp-dock")) return;
  const t = MP_CID && document.querySelector(`[data-cid="${MP_CID}"]`);
  if (!t || !t.offsetParent) closeMarginsPreview();
}

/* Podgląd to DOK po prawej stronie okna — aplikacja (wraz z pełną tabelą
   marginesów w kreatorze) przesuwa się w lewo, nic nie jest zasłonięte
   ani ucięte. Edytuje się prawdziwą tabelę, podgląd odświeża się sam. */
async function openMarginsPreview(cid, mode) {
  /* drugie kliknięcie tego samego przycisku = zamknięcie podglądu */
  if (MP_CID === cid) { closeMarginsPreview(); return; }
  closeMarginsPreview();
  MP_CID = cid;
  let typ = "OPTAX";
  let timer = null;
  let gen = 0;              /* numer żądania — ignorujemy odpowiedzi nieaktualne */

  const dock = el("div", "mp-dock");
  dock.id = "mp-dock";

  /* wolne, pływające okno: pozycję i rozmiar można zmieniać ręcznie
     (przeciąganie za pasek tytułowy + uchwyt w prawym dolnym rogu);
     ustawienia zapamiętujemy między otwarciami */
  let pos = null;
  try { pos = JSON.parse(localStorage.getItem("mpWindow") || "null"); } catch (e) {}
  if (!pos || !pos.w || !Number.isFinite(pos.x) || !Number.isFinite(pos.y)) {
    const w = Math.min(640, Math.round(window.innerWidth * 0.42));
    const h = Math.min(window.innerHeight - 48, 980);
    pos = { x: window.innerWidth - w - 16, y: 24, w, h };
  }
  pos.w = Math.max(340, Math.min(pos.w, window.innerWidth - 20));
  pos.h = Math.max(320, Math.min(pos.h, window.innerHeight - 20));
  pos.x = Math.max(0, Math.min(pos.x, window.innerWidth - 120));
  pos.y = Math.max(0, Math.min(pos.y, window.innerHeight - 60));
  dock.style.left = pos.x + "px";
  dock.style.top = pos.y + "px";
  dock.style.width = pos.w + "px";
  dock.style.height = pos.h + "px";
  const zapiszOkno = () => {
    try {
      localStorage.setItem("mpWindow", JSON.stringify({
        x: dock.offsetLeft, y: dock.offsetTop,
        w: dock.offsetWidth, h: dock.offsetHeight }));
    } catch (e) { /* prywatny tryb przeglądarki — trudno */ }
  };

  const head = el("div", "mp-head");
  head.appendChild(el("h3", null, "Podgląd marginesów"));
  const sel = el("select", "mp-typ");
  for (const t of MP_TYPY) {
    const o = el("option", null, t);
    o.value = t;
    sel.appendChild(o);
  }
  sel.onchange = () => { typ = sel.value; schedule(); };
  head.appendChild(sel);
  const pelny = el("button", "btn secondary small", "⛶");
  pelny.title = "Podgląd na cały ekran (ESC lub ⛶ wraca)";
  pelny.onclick = () => ustawPelnyEkran(!dock.classList.contains("full"));
  head.appendChild(pelny);
  const zamknij = el("button", "btn secondary small", "✕");
  zamknij.title = "Zamknij podgląd";
  zamknij.onclick = closeMarginsPreview;
  head.appendChild(zamknij);
  dock.appendChild(head);
  /* przeciąganie za pasek tytułowy (przyciski i lista nie ruszają okna) */
  let drag = null;
  head.addEventListener("mousedown", e => {
    if (dock.classList.contains("full")) return;
    if (e.target.closest("button, select, input")) return;
    drag = { dx: e.clientX - dock.offsetLeft, dy: e.clientY - dock.offsetTop };
    document.body.classList.add("mp-dragging");
    e.preventDefault();
  });
  const onDragMove = e => {
    if (!drag) return;
    dock.style.left = Math.max(-40, e.clientX - drag.dx) + "px";
    dock.style.top = Math.max(0, e.clientY - drag.dy) + "px";
  };
  const onDragUp = () => {
    if (!drag) return;
    drag = null;
    document.body.classList.remove("mp-dragging");
    zapiszOkno();
  };
  document.addEventListener("mousemove", onDragMove);
  document.addEventListener("mouseup", onDragUp);


  const info = el("div", "mp-note",
    "Podgląd 1. strony A4 (tak, jak wyjdzie z druku). Zmieniaj marginesy " +
    "w tabeli po lewej — podgląd odświeża się na żywo. " +
    "Drugi raz kliknięty przycisk podglądu zamyka panel.");
  dock.appendChild(info);

  const wrap = el("div", "mp-sheet-wrap");
  const sheet = el("div", "mp-sheet");
  const frame = el("iframe", "mp-frame");
  frame.setAttribute("title", "Podgląd dokumentu");
  sheet.appendChild(frame);
  wrap.appendChild(sheet);
  dock.appendChild(wrap);

  const errMsg = el("div", "mp-error hidden");
  dock.appendChild(errMsg);

  /* skalowanie arkusza A4 do szerokości panelu (także po pełnym ekranie) */
  let mpSzer = 794, mpWys = 1123;
  function ustawSkale() {
    const dostepne = Math.max(200, wrap.clientWidth - 12);
    const skala = Math.min(1, dostepne / mpSzer);
    sheet.style.transform = "scale(" + skala + ")";
    sheet.style.transformOrigin = "top left";
    /* po przeskalowaniu arkusz nie rezerwuje pełnej wysokości w doku */
    sheet.style.marginBottom = (mpWys * (skala - 1)) + "px";
  }
  const onResize = () => { if (document.getElementById("mp-dock")) ustawSkale(); };
  window.addEventListener("resize", onResize);
  /* ręczna zmiana rozmiaru okna (uchwyt w rogu) też przelicza skalę.
     MutationObserver na atrybucie style łapie zmiany szerokości/wysokości
     (uchwyt "resize" ustawia style inline), ResizeObserver — pozostałe */
  const poZmianieRozmiaru = () => { ustawSkale(); zapiszOkno(); };
  const ro = new ResizeObserver(poZmianieRozmiaru);
  ro.observe(dock);
  const mo = new MutationObserver(poZmianieRozmiaru);
  mo.observe(dock, { attributes: true, attributeFilter: ["style"] });

  function ustawPelnyEkran(on) {
    dock.classList.toggle("full", on);
    ustawSkale();
  }
  const onKey = e => {
    if (e.key === "Escape" && dock.classList.contains("full")) ustawPelnyEkran(false);
  };
  document.addEventListener("keydown", onKey);

  document.body.appendChild(dock);

  /* każde wpisanie w prawdziwej tabeli marginesów odświeża podgląd */
  const czcCid = "czcionki_" + mode;
  const onInput = e => {
    if (e.target.closest && e.target.closest(
        `[data-cid="${cid}"], [data-cid="${czcCid}"]`)) schedule();
  };
  document.addEventListener("input", onInput);
  MP_CLEANUP = () => {
    document.removeEventListener("input", onInput);
    window.removeEventListener("resize", onResize);
    window.removeEventListener("keydown", onKey);
    document.removeEventListener("mousemove", onDragMove);
    document.removeEventListener("mouseup", onDragUp);
    document.body.classList.remove("mp-dragging");
    ro.disconnect();
    mo.disconnect();
  };

  function schedule() {
    if (!document.getElementById("mp-dock")) { closeMarginsPreview(); return; }
    clearTimeout(timer);
    timer = setTimeout(refresh, 400);
  }

  async function refresh() {
    if (document.getElementById("mp-dock") !== dock) return;  /* zamknięto */
    const allVals = collectValues();
    const data = allVals[cid] || {};
    const czc = allVals[czcCid] || {};
    const wiersz = MP_WIERSZ[typ] || typ;
    const T = (data[wiersz] || {}).T || "1.5", B = (data[wiersz] || {}).B || "1.5",
          L = (data[wiersz] || {}).L || "2.5", R = (data[wiersz] || {}).R || "1.5";
    const moje = ++gen;
    errMsg.classList.add("hidden");
    sheet.classList.add("loading");
    let r = null;
    try { r = await api().get_margins_preview(typ, data, czc); }
    catch (e) { r = { ok: false, error: String(e) }; }
    if (moje !== gen) return;                 /* przyszła nieaktualna odpowiedź */
    if (!r.ok) {
      sheet.classList.add("hidden");
      errMsg.textContent = r.error || "Nie udało się przygotować podglądu.";
      errMsg.classList.remove("hidden");
      return;
    }
    sheet.classList.remove("hidden");
    info.textContent = (r.zrodlo || "dokument przykładowy") +
                       "  (" + typ + ", " + (r.poziom ? "poziomo" : "pionowo") + ")";
    /* podgląd JEDNEJ kartki A4 — dokładnie ta wielkość, co w druku;
       treść poza pierwszą stroną jest przycięta, a marginesy (padding)
       widać jako białe pole wokół treści */
    const szer = r.poziom ? 1123 : 794;       /* A4 w px przy 96 dpi */
    const wys = r.poziom ? 794 : 1123;
    mpSzer = szer; mpWys = wys;
    sheet.style.width = szer + "px";
    sheet.style.height = wys + "px";
    frame.style.width = szer + "px";
    frame.style.height = wys + "px";
    /* Cała treść trafia do "okna" dokładnie wielkości POLA DRUKU
       (strona minus marginesy) i jest przycinana po jego krawędziach —
       dzięki temu KAŻDY margines (także dół i prawo) widać na podglądzie,
       a zbyt szeroka tabela jest przycinana jak przy druku. */
    const css = `<style>
      html, body { background: #fff !important; max-width: none !important;
                  margin: 0 !important; padding: 0 !important;
                  overflow: hidden !important; }
      #mp-page { position: relative; width: ${szer}px; height: ${wys}px; }
      #mp-win { position: absolute; left: ${L}cm; top: ${T}cm;
                right: ${R}cm; bottom: ${B}cm; overflow: hidden; }
    </style>
    <script>
    (function () {
      function mpWrap() {
        if (document.getElementById("mp-win")) return;
        var p = document.createElement("div"); p.id = "mp-page";
        var w = document.createElement("div"); w.id = "mp-win";
        p.appendChild(w);
        while (document.body.firstChild) w.appendChild(document.body.firstChild);
        document.body.appendChild(p);
      }
      if (document.readyState === "loading")
        document.addEventListener("DOMContentLoaded", mpWrap);
      else mpWrap();
    })();
    <\/script>`;
    frame.onload = () => sheet.classList.remove("loading");
    frame.srcdoc = (r.html || "") + css;
    ustawSkale();
  }

  refresh();
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
const MARGIN_TYPES = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "HALIZNY", "WYK_NEG",
                     "OPIS", "ZEST1", "WSK_ZB", "WK_ZM1"];
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
  /* podgląd na żywo — tylko tam, gdzie marginesy sterują nowymi szablonami */
  if (c.mode === "ALL" || c.mode === "NS") {
    const pv = el("button", "btn ghost small mp-open-btn");
    pv.type = "button";
    pv.textContent = "👁  Podgląd dokumentu z marginesami";
    pv.onclick = async () => {
      /* osobne okno SYSTEMOWE — można je przesunąć po całym ekranie;
         gdyby nie wyszło (starsza wersja pywebview), stary podgląd w aplikacji */
      try {
        const r = await api().open_preview_window(c.mode || "ALL");
        if (r && r.ok) return;
      } catch (e) { /* spadamy do podglądu w obrębie programu */ }
      openMarginsPreview(c.id, c.mode);
    };
    wrap.appendChild(pv);
  }
  return wrap;
}

/* kontrolka: czcionki raportów nowego wyglądu (tytuł / tabela) */
function renderCzcionki(c) {
  const fonts = c.fonts || [""];
  const defPt = { tytul: "12", tabela: "8.6" };
  const wrap = el("div", null);
  const det = el("details", "czcionki-details");
  const sum = el("summary", null, escapeHtml(c.label || "Ustawienia czcionek:"));
  if (c.tooltip) sum.title = c.tooltip;
  det.appendChild(sum);
  const table = el("table", "czcionki-table");
  table.dataset.cid = c.id; table.dataset.kind = "czcionki";
  const head = el("tr", null, "<th>Raport</th>" +
                  "<th>Tytuł — rozmiar [pt]</th><th>Tytuł — czcionka</th>" +
                  "<th>Tabela — rozmiar [pt]</th><th>Tabela — czcionka</th>");
  table.appendChild(head);
  const data = VALUES[c.id] || {};
  for (const typ of (c.types || [])) {
    const tr = el("tr");
    tr.dataset.ftype = typ;
    tr.appendChild(el("td", null, typ));
    for (const gdzie of ["tytul", "tabela"]) {
      const vals = (data[typ] || {})[gdzie] || {};
      const inp = el("input", "czc-pt");
      inp.type = "text";
      inp.dataset.gdzie = gdzie; inp.dataset.co = "pt";
      inp.value = vals.pt !== undefined && vals.pt !== "" ? vals.pt : defPt[gdzie];
      inp.oninput = scheduleSetValues;
      const td1 = el("td", null, "");
      td1.appendChild(inp);
      tr.appendChild(td1);
      const sel = el("select", "czc-font");
      sel.dataset.gdzie = gdzie; sel.dataset.co = "font";
      for (const f of fonts) {
        const o = el("option", null, f === "" ? "(domyślna)" : f);
        o.value = f;
        if (f === (vals.font || "")) o.selected = true;
        sel.appendChild(o);
      }
      sel.onchange = scheduleSetValues;
      const td2 = el("td", null, "");
      td2.appendChild(sel);
      tr.appendChild(td2);
    }
    table.appendChild(tr);
  }
  det.appendChild(table);
  const hint = el("div", "hint-row",
    "Tytuł i tekst tabeli osobno dla każdego raportu. Obiekt, „Stan na” " +
    "i AGENCJA zostają z oryginalną czcionką.");
  det.appendChild(hint);
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
    } else if (kind === "czcionki") {
      const data = {};
      $$("tr[data-ftype]", node).forEach(tr => {
        const t = {};
        $$("[data-gdzie]", tr).forEach(el2 => {
          t[el2.dataset.gdzie] = t[el2.dataset.gdzie] || {};
          t[el2.dataset.gdzie][el2.dataset.co] = el2.value;
        });
        data[tr.dataset.ftype] = t;
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
  /* przełącznik "nowe szablony" (ekran powitalny kreatora 1-Click) */
  const _ns = document.getElementById("wiz-nowe-szablony");
  if (_ns) out.nowe_szablony = _ns.checked;
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
      LAST_STATUS_TEXT = ev.text || "";
      LAST_STATUS_ERR = (ev.color === "#D83B01");
      $("#status-dot").className = LAST_STATUS_ERR ? "err" : "";
      { const wo = document.getElementById("wiz-op");
        if (wo) wo.textContent = ev.text; }
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
      const wf = document.getElementById("wiz-fill");
      if (wf) wf.style.width = Math.max(0, Math.min(100, val * 100)) + "%";
      const wfl = document.getElementById("wiz-file");
      if (wfl) wfl.textContent = t;
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
  /* koniec zadania → ekran „Ukończono" w kreatorze + powiadomienie */
  if (was && !running) { wizDone(!LAST_STATUS_ERR); notifyTaskEnd(); }
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
    const backdrop = el("div", "modal-backdrop");
    const modal = el("div", "modal");
    if (ev.title) modal.appendChild(el("h3", null, escapeHtml(ev.title)));
    const msg = el("div", "modal-msg");
    msg.style.whiteSpace = "pre-line";
    msg.textContent = ev.message || "";
    modal.appendChild(msg);
    const row = el("div", "modal-row");
    const nie = el("button", "btn secondary", "Anuluj");
    const tak = el("button", "btn primary", "OK");
    nie.onclick = () => { backdrop.remove(); api().dialog_reply(ev.id, false); };
    tak.onclick = () => { backdrop.remove(); api().dialog_reply(ev.id, true); };
    row.appendChild(nie); row.appendChild(tak);
    modal.appendChild(row);
    backdrop.appendChild(modal);
    backdrop.onclick = e => {
      if (e.target === backdrop) { backdrop.remove(); api().dialog_reply(ev.id, false); }
    };
    document.body.appendChild(backdrop);
    setTimeout(() => tak.focus(), 50);
  } else {
    const t = toast(ev.title ? (ev.title + ": " + ev.message) : ev.message,
                    ev.kind === "error" ? "error" : ev.kind === "warn" ? "warn" : "ok");
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
  // zawsze pokazuj wszystkie szablony — zapis ze starszej wersji programu
  // mógł nie zawierać nowszych pozycji (HALIZNY, WYK_NEG, WK_ZM1...)
  for (const t of templates) if (!order.includes(t.key)) order.push(t.key);
  let dragKey = null;

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
      /* przeciąganie po liście (drag & drop) — alternatywa dla strzałek */
      it.draggable = true;
      it.addEventListener("dragstart", (e) => {
        dragKey = key;
        it.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
        try { e.dataTransfer.setData("text/plain", key); } catch (err) {}
      });
      it.addEventListener("dragend", () => {
        dragKey = null;
        it.classList.remove("dragging");
        list.querySelectorAll(".drag-over").forEach(x => x.classList.remove("drag-over"));
      });
      it.addEventListener("dragover", (e) => {
        e.preventDefault();
        if (dragKey !== null && dragKey !== key) it.classList.add("drag-over");
      });
      it.addEventListener("dragleave", () => it.classList.remove("drag-over"));
      it.addEventListener("drop", (e) => {
        e.preventDefault();
        it.classList.remove("drag-over");
        if (dragKey === null || dragKey === key) return;
        const from = order.indexOf(dragKey);
        const to = order.indexOf(key);
        if (from < 0 || to < 0) return;
        order.splice(from, 1);
        order.splice(to, 0, dragKey);
        renderList();
      });
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
