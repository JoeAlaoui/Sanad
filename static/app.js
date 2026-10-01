// ============================================================
// Helpdesk Dashboard - logique front-end (multi-volets)
// ============================================================
const JOURS = ["Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"];
const MOIS = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];

// Palette PCHC — reflète exactement COLOR_HEX défini dans pchc_import.py (source de vérité
// partagée avec le paramétrage statut→couleur, le PDF et l'export Excel/PPTX).
const PCHC_PALETTE = {green:"#25935F", blue:"#1794CF", yellow:"#F2C94C", orange:"#F2994A", red:"#DC2828", grey:"#94A3AD"};

function esc(s){
  if(s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}

if(window.Chart && window.ChartDataLabels){
  Chart.register(ChartDataLabels);
  Chart.defaults.set('plugins.datalabels', {display:false});
}

// ---------- Filtrage générique des tableaux de détail (Lot 3) ----------
// Un input.table-filter-input avec data-target="idDuTableau" filtre les lignes <tbody>
// du tableau ciblé en direct, sur l'ensemble du texte de la ligne (insensible à la casse).
// Délégation d'événement sur le document : fonctionne aussi pour les tableaux re-rendus
// dynamiquement après le premier chargement (pas besoin de ré-attacher l'écouteur).
document.addEventListener("input", (e)=>{
  if(!e.target.classList || !e.target.classList.contains("table-filter-input")) return;
  const table = document.getElementById(e.target.dataset.target);
  if(!table) return;
  const q = e.target.value.trim().toLowerCase();
  table.querySelectorAll("tbody tr").forEach(tr=>{
    tr.style.display = (!q || tr.textContent.toLowerCase().includes(q)) ? "" : "none";
  });
});

// ---------- Résilience réseau (Lot 1 — fiabilisation) ----------
// Intercepte les échecs fetch au niveau réseau (serveur injoignable, coupure Wi-Fi/VPN) pour
// informer l'utilisateur au lieu d'un échec silencieux. Les erreurs HTTP (4xx/5xx) restent
// gérées au cas par cas par chaque appelant qui vérifie res.ok ; celles qui ne le font pas
// (et finissent en promesse rejetée, ex. réponse HTML inattendue passée à .json()) sont
// rattrapées par le handler global ci-dessous.
const _nativeFetch = window.fetch.bind(window);
let _lastNetworkErrorToastAt = 0;
function notifyNetworkError(msg){
  const now = Date.now();
  if(now - _lastNetworkErrorToastAt > 4000){
    _lastNetworkErrorToastAt = now;
    toast(msg);
  }
}
window.fetch = function(...args){
  return _nativeFetch(...args).catch(err=>{
    notifyNetworkError("⚠️ Connexion au serveur interrompue — vérifiez votre réseau ou relancez l'application.");
    throw err;
  });
};
window.addEventListener("unhandledrejection", (event)=>{
  console.error("Erreur non gérée :", event.reason);
  notifyNetworkError("⚠️ Une erreur est survenue. Réessayez ou rechargez la page.");
});

let currentModule = "tarkhiss";
let currentUser = null;
let modulesInfo = {};
let state = { ym: null, calls: {}, emails: {}, analysis: {} };
let prevState = { calls: {}, emails: {}, analysis: {} };
let yoyState = { calls: {}, emails: {}, analysis: {} };
let globalSettings = { agency_name: "", logo_filename: null };
let moduleSettings = { contacts: [], greeting: "Bonjour,", signature_name: "", signature_function: "", signature_phone: "" };
let charts = {};
let annualChart = null, dailyLoadChart = null, weeklyEmailsChart = null, channelChart = null;

// Moussanada
let mState = { ym: null, month: null, timeseries: {}, notes: {} };
let mCharts = {};
let mAggMode = "racine"; // "racine" | "leaf"
let mAnnualChart = null, mBacklogChart = null, mCategoryTrendChart = null;

// ---------- API helpers ----------
function apiM(path){ return `/api/${currentModule}${path}`; }

// ---------- Helpers génériques ----------
function pad(n){ return n.toString().padStart(2,"0"); }

function businessDaysOfMonth(ym){
  const [y,m] = ym.split("-").map(Number);
  const days = [];
  const last = new Date(y, m, 0).getDate();
  for(let d=1; d<=last; d++){
    const date = new Date(y, m-1, d);
    const dow = date.getDay();
    days.push({ date: `${y}-${pad(m)}-${pad(d)}`, dow, label: JOURS[dow], weekend: (dow===0||dow===6) });
  }
  return days;
}

function prevYm(ym){
  const [y,m] = ym.split("-").map(Number);
  const d = new Date(y, m-2, 1);
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}`;
}

// Même mois, année précédente — pour le comparatif N vs N-1 (admin + superviseur)
function prevYearYm(ym){
  const [y,m] = ym.split("-").map(Number);
  return `${y-1}-${pad(m)}`;
}

function monthLabel(ym){
  const [y,m] = ym.split("-").map(Number);
  return `${MOIS[m-1]} ${y}`;
}

function toast(msg){
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"), 2600);
}

// ---------- Overlay de chargement (Lot 5) — imports/générations volumineuses ----------
function showBusy(message){
  const overlay = document.getElementById("busyOverlay");
  const text = document.getElementById("busyText");
  if(!overlay) return;
  if(text) text.textContent = message || "Traitement en cours...";
  overlay.style.display = "flex";
}
function hideBusy(){
  const overlay = document.getElementById("busyOverlay");
  if(overlay) overlay.style.display = "none";
}

function saveIndicatorPulse(){
  const el = document.getElementById("saveIndicator");
  el.textContent = "✓ Enregistré";
  el.classList.add("show");
  setTimeout(()=>el.classList.remove("show"), 1800);
  // Une écriture vient d'avoir lieu : rafraîchir l'indicateur "Dernière synchro" du volet actif
  refreshLastSync();
}

function relativeTimeFr(iso){
  const d = new Date(iso);
  const diffMin = Math.round((Date.now() - d.getTime()) / 60000);
  if(diffMin < 1) return "à l'instant";
  if(diffMin < 60) return `il y a ${diffMin} min`;
  const diffH = Math.round(diffMin / 60);
  if(diffH < 24) return `il y a ${diffH} h`;
  const diffJ = Math.round(diffH / 24);
  if(diffJ === 1) return "hier";
  if(diffJ < 7) return `il y a ${diffJ} j`;
  return d.toLocaleDateString("fr-FR", {day:"2-digit", month:"short"});
}

async function refreshLastSync(){
  const el = document.getElementById("lastSyncIndicator");
  if(!el || currentModule === undefined) return;
  try{
    const res = await fetch(`/api/${currentModule}/last-sync`);
    const data = await res.json();
    if(!data.last_sync){
      el.classList.add("empty");
      return;
    }
    el.classList.remove("empty");
    const diffH = (Date.now() - new Date(data.last_sync).getTime()) / 3600000;
    el.classList.toggle("stale", diffH > 72);
    el.textContent = `Synchro ${relativeTimeFr(data.last_sync)}`;
  } catch(e){
    el.classList.add("empty");
  }
}

function debounce(fn, delay){
  let t;
  return (...args)=>{ clearTimeout(t); t = setTimeout(()=>fn(...args), delay); };
}

function pulseAutosave(elId){
  const el = document.getElementById(elId);
  if(!el) return;
  el.textContent = "💾 Sauvegarde auto...";
  el.classList.add("show");
  setTimeout(()=>{ el.textContent = "✓ Sauvegardé automatiquement"; }, 400);
  setTimeout(()=>el.classList.remove("show"), 2200);
}

// ---------- hh:mm:ss ----------
function parseHMS(str){
  if(!str) return 0;
  str = str.trim();
  const parts = str.split(":").map(p=>Number(p));
  if(parts.some(isNaN)) return 0;
  if(parts.length === 3) return parts[0]*3600 + parts[1]*60 + parts[2];
  if(parts.length === 2) return parts[0]*60 + parts[1];
  if(parts.length === 1) return parts[0];
  return 0;
}

function formatHMS(totalSeconds){
  totalSeconds = Math.round(totalSeconds || 0);
  const h = Math.floor(totalSeconds/3600);
  const m = Math.floor((totalSeconds%3600)/60);
  const s = totalSeconds%60;
  return `${pad(h)}:${pad(m)}:${pad(s)}`;
}

// ============================================================
// MODULES (Tarkhiss / Moussanada)
// ============================================================
// pchc est un sous-volet de tarkhiss (Reporting Métier, nichée sous Tarkhiss dans la sidebar).
// parentOf() ramène tout module au "groupe" affiché dans le sélecteur de premier niveau.
function parentOf(mod){ return mod === "pchc" ? "tarkhiss" : mod; }

async function loadModules(){
  const res = await fetch("/api/modules");
  modulesInfo = await res.json();
  const allowedModules = (currentUser?.modules) || Object.keys(modulesInfo);
  // Sélecteur de premier niveau : Tarkhiss (englobe pchc) / Moussanada — pchc n'a pas de bouton propre
  const parentKeys = Object.keys(modulesInfo).filter(k=>k !== "pchc" && (allowedModules.includes(k) || (k === "tarkhiss" && allowedModules.includes("pchc"))));
  const sw = document.getElementById("moduleSwitch");
  sw.innerHTML = parentKeys.map(key=>`
    <button data-parent="${key}" class="${key===parentOf(currentModule)?'active':''}">${modulesInfo[key].label}</button>
  `).join("");
  sw.querySelectorAll("button").forEach(b=>{
    b.addEventListener("click", ()=>{
      if(b.dataset.parent === parentOf(currentModule)) return;
      switchModule(b.dataset.parent);
    });
  });
  if(!allowedModules.includes(currentModule)){
    currentModule = allowedModules.includes("tarkhiss") ? "tarkhiss" : (allowedModules[0] || "tarkhiss");
  }
  applyModuleBrand();
}

function defaultTabForRole(){
  if(!currentUser) return "dashboard";
  if(currentUser.role === "hotliner") return "knowledge";
  if(currentModule === "moussanada") return "m-dashboard";
  if(currentModule === "pchc") return "p-dashboard";
  return "dashboard";
}

function applyModuleBrand(){
  const info = modulesInfo[currentModule] || {label: currentModule, subtitle:"", ready:true};
  document.getElementById("brandMark").textContent = info.label[0];
  document.getElementById("brandTitle").textContent = info.label;
  document.getElementById("brandSub").textContent = info.subtitle;
  refreshLastSync();
  document.querySelectorAll(".module-switch button").forEach(b=>{
    b.classList.toggle("active", b.dataset.parent === parentOf(currentModule));
  });
  ["adminModuleLabel","adminModuleLabel2","adminModuleLabel3","adminModuleLabel4","adminModuleLabel5","adminModuleLabel6","adminModuleLabel7","adminModuleLabel8"].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.textContent = info.label;
  });
  document.getElementById("moduleNotReadyBanner").style.display = info.ready ? "none" : "block";
  // Sélecteur de mois : masqué pour le hotliner sauf sur Tarkhiss (seul volet où il a un accès
  // lecture seule — Dashboard et Vue annuelle). Actions serveur (export/email) restreintes de
  // même : le hotliner ne fait que consulter, jamais générer/envoyer.
  const hotlinerReadOnly = currentUser && currentUser.role === "hotliner";
  document.getElementById("monthPickerWrap").style.display = (hotlinerReadOnly && currentModule !== "tarkhiss") ? "none" : "";
  ["xlsxBtn", "pdfBtn", "mailToggleBtn"].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.style.display = hotlinerReadOnly ? "none" : "";
  });

  // Filtre les items de nav (et titres de sous-groupe) selon le GROUPE de volet actif ET le rôle.
  // Quand le groupe actif est "tarkhiss", les deux sous-volets (Hotline&Emails + Métier) restent
  // visibles simultanément dans la sidebar (accordéon ouvert), sans écran intermédiaire.
  const activeParent = parentOf(currentModule);
  document.querySelectorAll(".nav-item, .nav-subgroup-title, .nav-subgroup").forEach(btn=>{
    const m = btn.dataset.module;
    const modOk = (m === "shared" || parentOf(m) === activeParent);
    const rolesAttr = btn.dataset.roles;
    const roleOk = !rolesAttr || (currentUser && rolesAttr.split(",").includes(currentUser.role));
    btn.style.display = (modOk && roleOk) ? "" : "none";
  });

  document.getElementById("tarkhissCsvCard").style.display = currentModule === "tarkhiss" ? "" : "none";
  document.getElementById("moussanadaAdminNote").style.display = currentModule === "moussanada" ? "block" : "none";

  // Si l'onglet actif n'est plus valide (volet ou rôle), basculer sur l'onglet par défaut
  const activeBtn = document.querySelector(".nav-item.active");
  const activeVisible = activeBtn && activeBtn.style.display !== "none";
  if(!activeVisible){
    activateTab(defaultTabForRole());
  }
}

function activateTab(tab){
  const prevTab = document.querySelector(".nav-item.active")?.dataset.tab;
  if(prevTab === "admin" && tab !== "admin"){
    revertThemePreviewIfUnsaved();
  }
  document.querySelectorAll(".nav-item").forEach(b=>b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
  document.getElementById("tab-"+tab).classList.add("active");
  updateBreadcrumb(tab);
  renderTab(tab);
}

// ---------- Fil d'ariane (Lot G) — dérivé de la structure de nav existante, pas de mapping
// statique à maintenir séparément : on lit le data-module et le sous-groupe le plus proche.
function updateBreadcrumb(tab){
  const bc = document.getElementById("breadcrumb");
  if(!bc) return;
  const navItem = document.querySelector(`.nav-item[data-tab="${tab}"]:not([data-real-module])`);
  if(!navItem){ bc.innerHTML = ""; return; }
  const itemLabel = navItem.textContent.trim().replace(/\s+/g, " ");
  const module = navItem.dataset.module;
  const parts = [];
  if(module === "tarkhiss" || module === "pchc"){
    parts.push("Tarkhiss");
    let sib = navItem.previousElementSibling;
    while(sib && !sib.classList.contains("nav-subgroup-title")) sib = sib.previousElementSibling;
    if(sib) parts.push(sib.textContent.trim());
  } else if(module === "moussanada"){
    parts.push("Moussanada");
  }
  parts.push(itemLabel);
  bc.innerHTML = parts.map((p,i)=> i === parts.length-1
    ? `<span class="bc-current">${esc(p)}</span>`
    : `<span>${esc(p)}</span><span class="bc-sep">›</span>`
  ).join("");

  const favBtn = document.getElementById("favoriteToggleBtn");
  if(favBtn) updateFavoriteBtn(tab);
}

// ---------- Favoris (Lot G) — épingles personnelles, persistées par navigateur ----------
function loadFavorites(){
  try{ return JSON.parse(localStorage.getItem("sanad_favorites") || "[]"); }
  catch(e){ return []; }
}
function saveFavoritesList(list){
  localStorage.setItem("sanad_favorites", JSON.stringify(list.slice(0, 8)));
}
function isFavorited(tab){
  return loadFavorites().some(f=>f.tab === tab);
}
function toggleCurrentFavorite(){
  const tab = document.querySelector(".nav-item.active:not([data-real-module])")?.dataset.tab;
  if(!tab) return;
  const navItem = document.querySelector(`.nav-item[data-tab="${tab}"]:not([data-real-module])`);
  const label = navItem ? navItem.textContent.trim().replace(/\s+/g, " ") : tab;
  const module = navItem ? navItem.dataset.module : currentModule;
  let list = loadFavorites();
  if(list.some(f=>f.tab === tab)){
    list = list.filter(f=>f.tab !== tab);
  } else {
    list.push({tab, label, module});
  }
  saveFavoritesList(list);
  renderFavorites();
  updateFavoriteBtn(tab);
}
function renderFavorites(){
  const group = document.getElementById("favoritesGroup");
  const title = document.getElementById("favoritesTitle");
  const list = loadFavorites();
  if(!list.length){ group.style.display = "none"; title.style.display = "none"; group.innerHTML = ""; return; }
  title.style.display = ""; group.style.display = "";
  // data-module="shared" : toujours visible quel que soit le volet actif (c'est le principe
  // même d'un raccourci) — le clic bascule quand même le bon module via data-real-module.
  group.innerHTML = list.map(f=>`
    <button class="nav-item" data-tab="${f.tab}" data-module="shared" data-real-module="${f.module}">
      <span class="nav-ico">⭐</span> ${esc(f.label)}
    </button>
  `).join("");
}
function updateFavoriteBtn(tab){
  const btn = document.getElementById("favoriteToggleBtn");
  if(!btn) return;
  btn.textContent = isFavorited(tab) ? "★" : "☆";
  btn.title = isFavorited(tab) ? "Retirer des favoris" : "Épingler cette vue en favori";
  btn.classList.toggle("is-favorited", isFavorited(tab));
}

function renderTab(tab){
  if(tab === "dashboard") renderDashboard();
  if(tab === "annual") loadAnnual();
  if(tab === "prompts") renderPrompts();
  if(tab === "admin") loadAdmin();
  if(tab === "m-dashboard") renderMDashboard();
  if(tab === "m-sources") renderMSources();
  if(tab === "m-analysis") renderMAnalysis();
  if(tab === "m-annual") loadMAnnual();
  if(tab === "m-guide") renderMGuide();
  if(tab === "users") loadUsers();
  if(tab === "notes") loadNotes();
  if(tab === "knowledge") loadKnowledge();
  if(tab === "p-dashboard") renderPDashboard();
  if(tab === "p-import") renderPImport();
  if(tab === "p-analysis") renderPAnalysis();
}

async function switchModule(mod){
  if(mod === currentModule) return;
  currentModule = mod;
  if(currentUser && currentUser.role !== "hotliner"){
    await loadGlobalSettings();
    await loadModuleSettings();
  }
  applyModuleBrand();
  if(currentUser && currentUser.role !== "hotliner"){
    if(mod === "tarkhiss"){
      await loadMonth(state.ym || defaultMonth());
    } else if(mod === "moussanada"){
      await loadMoussanadaMonth(mState.ym || defaultMonth());
    }
  }
  const activeTab = document.querySelector(".nav-item.active")?.dataset.tab;
  if(activeTab) renderTab(activeTab);
}

// ---------- Navigation ----------
// Les deux sous-volets de Tarkhiss (Hotline&Emails / Métier) étant visibles en même temps dans
// la sidebar (accordéon), un clic peut passer de l'un à l'autre sans repasser par le sélecteur
// de premier niveau : on bascule currentModule (routage /api/<module>/...) avant d'activer l'onglet.
// Délégation d'événement (plutôt qu'un attachement par élément) pour que les favoris (Lot G),
// ajoutés dynamiquement dans la sidebar, bénéficient du même comportement sans code dupliqué.
document.addEventListener("click", async (e)=>{
  const btn = e.target.closest(".nav-item");
  if(!btn || !document.contains(btn)) return;
  const mod = btn.dataset.realModule || btn.dataset.module;
  if(mod && mod !== "shared" && mod !== currentModule){
    await switchModule(mod);
  }
  activateTab(btn.dataset.tab);
});

// ---------- Month handling ----------
const monthInput = document.getElementById("monthSelect");

function sleep(ms){ return new Promise(r=>setTimeout(r, ms)); }

// Transition douce du contenu au changement de mois : fondu/glissement léger sur l'onglet
// actif pendant le rechargement, plutôt qu'un simple "saut" du contenu.
async function withMonthTransition(loaderFn, ym){
  const panel = document.querySelector(".tab-panel.active");
  if(panel){
    panel.classList.add("month-transitioning");
    await sleep(130);
  }
  try{
    await loaderFn(ym);
  } finally {
    if(panel){
      requestAnimationFrame(()=> panel.classList.remove("month-transitioning"));
    }
  }
}

function goToMonth(ym){
  monthInput.value = ym;
  const loader = currentModule === "tarkhiss" ? loadMonth : loadMoussanadaMonth;
  withMonthTransition(loader, ym);
}

monthInput.addEventListener("change", ()=>{
  withMonthTransition(currentModule === "tarkhiss" ? loadMonth : loadMoussanadaMonth, monthInput.value);
});

document.getElementById("monthPrevBtn").addEventListener("click", ()=>{
  goToMonth(prevYm(monthInput.value || defaultMonth()));
});
document.getElementById("monthNextBtn").addEventListener("click", ()=>{
  goToMonth(nextYm(monthInput.value || defaultMonth()));
});

function nextYm(ym){
  const [y, m] = ym.split("-").map(Number);
  return m === 12 ? `${y+1}-01` : `${y}-${pad(m+1)}`;
}


function defaultMonth(){
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}`;
}

async function fetchMonthData(ym){
  const res = await fetch(apiM(`/month/${ym}`));
  return res.json();
}

async function loadMonth(ym){
  state.ym = ym;
  monthInput.value = ym;
  const data = await fetchMonthData(ym);
  state.calls = data.calls || {};
  state.emails = data.emails || {};
  state.analysis = data.analysis || { problems:[], demandes:[], weekly:[], constats:[], recommandations:[], keywords:[], internal_notes:"" };

  const prevData = await fetchMonthData(prevYm(ym));
  prevState.calls = prevData.calls || {};
  prevState.emails = prevData.emails || {};
  prevState.analysis = prevData.analysis || { problems:[], demandes:[], weekly:[] };

  // Comparatif N vs N-1 (même mois, année précédente) — réservé admin/superviseur
  if(currentUser && (currentUser.role === "admin" || currentUser.role === "superviseur")){
    const yoyData = await fetchMonthData(prevYearYm(ym));
    yoyState.calls = yoyData.calls || {};
    yoyState.emails = yoyData.emails || {};
    yoyState.analysis = yoyData.analysis || { problems:[], demandes:[], weekly:[] };
  }

  // Rappel d'envoi mensuel : n'a de sens que pour qui peut réellement envoyer (admin/superviseur).
  // Le hotliner a un accès lecture seule au dashboard Tarkhiss et n'a pas droit à /send-log.
  if(!currentUser || currentUser.role !== "hotliner"){
    checkReminder();
  }

  renderCallsTable();
  renderEmailsTable();
  renderAnalysisForms();

  const activeTab = document.querySelector(".nav-item.active")?.dataset.tab;
  if(activeTab === "dashboard") renderDashboard();
  if(activeTab === "prompts") renderPrompts();
}

async function checkReminder(){
  const res = await fetch(apiM("/send-log"));
  const log = await res.json();
  const py = prevYm(state.ym);
  const sent = log.some(l=>l.ym === py);
  const banner = document.getElementById("reminderBanner");
  const reminderDay = moduleSettings.reminder_day || 5;
  const today = new Date();
  if(!sent && today.getDate() >= reminderDay){
    banner.style.display = "block";
    banner.textContent = `⚠️ Le rapport de ${monthLabel(py)} n'a pas encore été envoyé (rappel programmé à partir du ${reminderDay} du mois).`;
    notifyReminder("Tarkhiss", py);
  } else {
    banner.style.display = "none";
  }
}

async function checkMoussanadaReminder(){
  const res = await fetch("/api/moussanada/send-log");
  const log = await res.json();
  const py = prevYmM(mState.ym);
  const sent = log.some(l=>l.ym === py);
  const banner = document.getElementById("mReminderBanner");
  const reminderDay = moduleSettings.reminder_day || 5;
  const today = new Date();
  if(!sent && today.getDate() >= reminderDay){
    banner.style.display = "block";
    banner.textContent = `⚠️ Le rapport de ${monthLabel(py)} n'a pas encore été envoyé (rappel programmé à partir du ${reminderDay} du mois).`;
    notifyReminder("Moussanada", py);
  } else {
    banner.style.display = "none";
  }
}

let notifiedThisSession = {};
function notifyReminder(moduleLabel, py){
  const key = moduleLabel+py;
  if(notifiedThisSession[key]) return;
  notifiedThisSession[key] = true;
  notifyBrowser(`Rapport ${moduleLabel} en attente`, `Le rapport de ${monthLabel(py)} n'a pas encore été envoyé.`, false);
}

// Notification navigateur générique — utilisée pour le rappel de rapport (ci-dessus) et pour
// les opérations longues terminées pendant que l'onglet est en arrière-plan (imports).
// Ne redemande jamais la permission de façon intrusive : seule la première notification
// déclenche potentiellement une demande native du navigateur.
// onlyIfHidden=true : n'affiche que si l'onglet est actuellement en arrière-plan (évite les
// notifications redondantes avec le toast déjà visible quand l'utilisateur regarde l'écran).
function notifyBrowser(title, body, onlyIfHidden=true){
  if(!("Notification" in window)) return;
  if(Notification.permission === "granted"){
    if(!onlyIfHidden || document.hidden) new Notification(title, {body});
  } else if(Notification.permission !== "denied"){
    Notification.requestPermission();
  }
}

// ---------- Import CSV générique ----------
function parseCsv(text){
  const lines = text.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
  const delim = lines[0].includes(";") ? ";" : (lines[0].includes("\t") ? "\t" : ",");
  return lines.map(l=>l.split(delim).map(c=>c.trim()));
}

function normalizeDate(s){
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m) return `${m[3]}-${pad(m[1])}-${pad(m[2])}`;
  return null;
}

// ============================================================
// APPELS
// ============================================================
function computeCompletude(days, obj, field){
  const businessDays = days.filter(d=>!d.weekend);
  const filled = businessDays.filter(d=> obj[d.date] && obj[d.date][field] !== undefined && obj[d.date][field] !== null && obj[d.date][field] !== "").length;
  return { filled, total: businessDays.length };
}

function getDurationSec(rec){
  if(rec.duration_sec !== undefined) return rec.duration_sec || 0;
  if(rec.duration_min !== undefined) return (rec.duration_min || 0) * 60;
  return 0;
}

function renderCallsTable(){
  const tbody = document.querySelector("#callsTable tbody");
  tbody.innerHTML = "";
  const days = businessDaysOfMonth(state.ym);
  let total = 0, totalSec = 0;
  days.forEach(day=>{
    const rec = state.calls[day.date] || {};
    if(!day.weekend){ total += Number(rec.calls || 0); totalSec += getDurationSec(rec); }
    const missing = !day.weekend && (rec.calls === undefined || rec.calls === null || rec.calls === "");
    const tr = document.createElement("tr");
    if(day.weekend) tr.classList.add("weekend");
    if(missing) tr.classList.add("missing-day");
    const hmsVal = rec.calls !== undefined ? formatHMS(getDurationSec(rec)) : "";
    tr.innerHTML = `
      <td class="day-label">${day.label}</td>
      <td class="date-label">${day.date}</td>
      <td>${day.weekend ? "—" : `<input type="number" min="0" data-date="${day.date}" data-field="calls" value="${rec.calls ?? ''}">`}</td>
      <td>${day.weekend ? "—" : `<input type="text" placeholder="00:00:00" pattern="\\d{1,3}:\\d{2}:\\d{2}" data-date="${day.date}" data-field="duration_hms" value="${hmsVal}">`}</td>
    `;
    tbody.appendChild(tr);
  });
  document.getElementById("callsTotalHint").textContent = total;
  document.getElementById("callsDurationHint").textContent = formatHMS(totalSec);
  updateBadge("badgeCalls", computeCompletude(days, state.calls, "calls"));

  const autosave = debounce(()=>saveCalls(true), 1500);
  tbody.querySelectorAll("input").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      let t = 0, ts = 0;
      tbody.querySelectorAll('input[data-field="calls"]').forEach(i=>t += Number(i.value||0));
      tbody.querySelectorAll('input[data-field="duration_hms"]').forEach(i=>ts += parseHMS(i.value));
      document.getElementById("callsTotalHint").textContent = t;
      document.getElementById("callsDurationHint").textContent = formatHMS(ts);
      autosave();
    });
  });
}

function collectCallsRows(){
  const tbody = document.querySelector("#callsTable tbody");
  const rows = {};
  const dates = [...new Set([...tbody.querySelectorAll("input")].map(i=>i.dataset.date))];
  dates.forEach(date=>{
    const calls = tbody.querySelector(`input[data-date="${date}"][data-field="calls"]`);
    const hms = tbody.querySelector(`input[data-date="${date}"][data-field="duration_hms"]`);
    if(calls && calls.value !== ""){
      rows[date] = { calls: Number(calls.value||0), duration_sec: parseHMS(hms ? hms.value : "") };
    }
  });
  return rows;
}

async function saveCalls(silent){
  const rows = collectCallsRows();
  state.calls = rows;
  await fetch(apiM(`/calls/${state.ym}`), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(rows)});
  if(silent) pulseAutosave("callsAutosaveHint");
  updateBadge("badgeCalls", computeCompletude(businessDaysOfMonth(state.ym), rows, "calls"));
}

document.getElementById("saveCallsBtn").addEventListener("click", async ()=>{
  await saveCalls(false);
  saveIndicatorPulse();
  toast("Appels enregistrés");
});

document.getElementById("callsImportFile").addEventListener("change", (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = ()=>{
    const rows = parseCsv(reader.result);
    let count = 0;
    rows.forEach(cols=>{
      if(cols.length < 2) return;
      const date = normalizeDate(cols[0]);
      if(!date) return;
      const calls = Number(cols[1]||0);
      const durSec = cols[2] ? parseHMS(cols[2]) : 0;
      state.calls[date] = { calls, duration_sec: durSec };
      count++;
    });
    renderCallsTable();
    saveCalls(false);
    toast(`${count} jour(s) importé(s)`);
    e.target.value = "";
  };
  reader.readAsText(file, "UTF-8");
});

document.getElementById("callsRawImportFile").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const box = document.getElementById("callsRawImportResult");
  box.style.display = "";
  box.innerHTML = "Analyse du journal d'appels en cours...";
  showBusy("Import du journal d'appels en cours...");
  const fd = new FormData();
  fd.append("file", file);
  try{
    const res = await fetch("/api/tarkhiss/import-calls-raw", {method:"POST", body: fd});
    const data = await res.json();
    if(data.ok){
      const s = data.stats;
      box.innerHTML = `
        <h3 style="margin:0 0 8px;font-size:14px;color:var(--primary);">✅ Import réussi</h3>
        <p style="font-size:12.5px;color:var(--muted);margin:0 0 8px;">
          ${s.months_covered.length} mois resynchronisés (${s.months_covered.map(monthLabel).join(", ")}) — ${s.rows} lignes lues${s.skipped ? `, ${s.skipped} ignorée(s)` : ""}
        </p>
        <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);">
          <div class="kpi-card"><div class="kpi-label">Reçus (décrochés)</div><div class="kpi-value">${s.total_in}</div></div>
          <div class="kpi-card"><div class="kpi-label">Manqués</div><div class="kpi-value">${s.total_missed}</div></div>
          <div class="kpi-card"><div class="kpi-label">Émis</div><div class="kpi-value">${s.total_out}</div></div>
          <div class="kpi-card"><div class="kpi-label">Taux de décroché</div><div class="kpi-value">${s.answer_rate!=null?s.answer_rate+"%":"—"}</div></div>
        </div>`;
      toast("Import journal d'appels réussi — mois resynchronisés");
      if(state.ym) await loadMonth(state.ym);
    } else {
      box.innerHTML = `<span style="color:var(--danger);">❌ ${esc(data.error || "Erreur d'import")}</span>`;
      toast(data.error || "Erreur d'import");
    }
  } catch(err){
    box.innerHTML = `<span style="color:var(--danger);">❌ Erreur réseau lors de l'import.</span>`;
  } finally {
    hideBusy();
    e.target.value = "";
  }
});

// ============================================================
// EMAILS (week-ends éditables)
// ============================================================
function renderEmailsTable(){
  const tbody = document.querySelector("#emailsTable tbody");
  tbody.innerHTML = "";
  const days = businessDaysOfMonth(state.ym);
  let totalR = 0, totalS = 0;
  days.forEach(day=>{
    const rec = state.emails[day.date] || {};
    totalR += Number(rec.received||0);
    totalS += Number(rec.sent||0);
    const missing = !day.weekend && (rec.received === undefined || rec.received === null || rec.received === "");
    const tr = document.createElement("tr");
    if(day.weekend) tr.classList.add("weekend");
    if(missing) tr.classList.add("missing-day");
    tr.innerHTML = `
      <td class="day-label">${day.label}${day.weekend ? ' <span class="hint" style="font-size:10px;">(w-e)</span>' : ''}</td>
      <td class="date-label">${day.date}</td>
      <td><input type="number" min="0" data-date="${day.date}" data-field="received" value="${rec.received ?? ''}"></td>
      <td><input type="number" min="0" data-date="${day.date}" data-field="sent" value="${rec.sent ?? ''}"></td>
    `;
    tbody.appendChild(tr);
  });
  document.getElementById("emailsRecvHint").textContent = totalR;
  document.getElementById("emailsSentHint").textContent = totalS;
  updateBadge("badgeEmails", computeCompletude(days, state.emails, "received"));

  const autosave = debounce(()=>saveEmails(true), 1500);
  tbody.querySelectorAll("input").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      let r=0,s=0;
      tbody.querySelectorAll('input[data-field="received"]').forEach(i=>r+=Number(i.value||0));
      tbody.querySelectorAll('input[data-field="sent"]').forEach(i=>s+=Number(i.value||0));
      document.getElementById("emailsRecvHint").textContent = r;
      document.getElementById("emailsSentHint").textContent = s;
      autosave();
    });
  });
}

function collectEmailsRows(){
  const tbody = document.querySelector("#emailsTable tbody");
  const rows = {};
  const dates = [...new Set([...tbody.querySelectorAll("input")].map(i=>i.dataset.date))];
  dates.forEach(date=>{
    const rec = tbody.querySelector(`input[data-date="${date}"][data-field="received"]`);
    const sent = tbody.querySelector(`input[data-date="${date}"][data-field="sent"]`);
    if((rec && rec.value !== "") || (sent && sent.value !== "")){
      rows[date] = { received: Number(rec.value||0), sent: Number(sent.value||0) };
    }
  });
  return rows;
}

async function saveEmails(silent){
  const rows = collectEmailsRows();
  state.emails = rows;
  await fetch(apiM(`/emails/${state.ym}`), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(rows)});
  if(silent) pulseAutosave("emailsAutosaveHint");
  updateBadge("badgeEmails", computeCompletude(businessDaysOfMonth(state.ym), rows, "received"));
}

document.getElementById("saveEmailsBtn").addEventListener("click", async ()=>{
  await saveEmails(false);
  saveIndicatorPulse();
  toast("Emails enregistrés");
});

document.getElementById("emailsImportFile").addEventListener("change", (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = ()=>{
    const rows = parseCsv(reader.result);
    let count = 0;
    rows.forEach(cols=>{
      if(cols.length < 2) return;
      const date = normalizeDate(cols[0]);
      if(!date) return;
      state.emails[date] = { received: Number(cols[1]||0), sent: Number(cols[2]||0) };
      count++;
    });
    renderEmailsTable();
    saveEmails(false);
    toast(`${count} jour(s) importé(s)`);
    e.target.value = "";
  };
  reader.readAsText(file, "UTF-8");
});

document.getElementById("emailsRawImportFile").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const box = document.getElementById("emailsRawImportResult");
  box.style.display = "";
  box.innerHTML = "Analyse du fichier Outlook en cours...";
  showBusy("Import de l'export Outlook en cours (peut prendre quelques secondes sur un gros historique)...");
  const fd = new FormData();
  fd.append("file", file);
  try{
    const res = await fetch("/api/tarkhiss/import-emails-raw", {method:"POST", body: fd});
    const data = await res.json();
    if(data.ok){
      const s = data.stats;
      box.innerHTML = `
        <h3 style="margin:0 0 8px;font-size:14px;color:var(--primary);">✅ Import réussi</h3>
        <p style="font-size:12.5px;color:var(--muted);margin:0 0 8px;">
          ${s.months_covered.length} mois resynchronisés (${s.months_covered.map(monthLabel).join(", ")})
        </p>
        <div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);">
          <div class="kpi-card"><div class="kpi-label">Total reçus</div><div class="kpi-value">${s.total_received}</div></div>
          <div class="kpi-card"><div class="kpi-label">Total envoyés</div><div class="kpi-value">${s.total_sent}</div></div>
          <div class="kpi-card"><div class="kpi-label">Avec pièce jointe</div><div class="kpi-value">${s.pct_with_attachment}%</div></div>
          <div class="kpi-card"><div class="kpi-label">Importance haute</div><div class="kpi-value">${s.pct_high_importance}%</div></div>
        </div>`;
      toast("Import Outlook réussi — mois resynchronisés");
      if(state.ym) await loadMonth(state.ym);
    } else {
      box.innerHTML = `<span style="color:var(--danger);">❌ ${esc(data.error || "Erreur d'import")}</span>`;
      toast(data.error || "Erreur d'import");
    }
  } catch(err){
    box.innerHTML = `<span style="color:var(--danger);">❌ Erreur réseau lors de l'import.</span>`;
  } finally {
    hideBusy();
    e.target.value = "";
  }
});

function updateBadge(elId, completude){
  const el = document.getElementById(elId);
  if(!el) return;
  el.textContent = `${completude.filled}/${completude.total}`;
  el.classList.toggle("warn", completude.filled < completude.total);
}

// ============================================================
// ANALYSE MENSUELLE
// ============================================================
function catRow(tableBodySel, kind, item={}){
  const tbody = document.querySelector(tableBodySel);
  const tr = document.createElement("tr");
  if(kind === "weekly"){
    tr.innerHTML = `
      <td><input type="text" class="f-week" value="${item.week||''}" placeholder="ex: 1-7 juin"></td>
      <td><input type="number" class="f-bugs" value="${item.bugs??''}"></td>
      <td><input type="number" class="f-demandes" value="${item.demandes??''}"></td>
      <td><textarea class="f-obs" rows="1">${item.obs||''}</textarea></td>
      <td><button class="row-del">✕</button></td>
    `;
  } else if(kind === "problems"){
    tr.innerHTML = `
      <td><input type="text" class="f-label" value="${item.label||''}" placeholder="ex: Problèmes paiement"></td>
      <td><input type="number" class="f-count" value="${item.count??''}"></td>
      <td><textarea class="f-desc" rows="1">${item.desc||''}</textarea></td>
      <td><textarea class="f-action" rows="1" placeholder="Action / responsable">${item.action||''}</textarea></td>
      <td><button class="row-del">✕</button></td>
    `;
  } else if(kind === "keywords"){
    tr.innerHTML = `
      <td><input type="text" class="f-word" value="${item.word||''}" placeholder="ex: paiement"></td>
      <td><input type="number" class="f-kwcount" value="${item.count??''}"></td>
      <td><button class="row-del">✕</button></td>
    `;
  } else {
    tr.innerHTML = `
      <td><input type="text" class="f-label" value="${item.label||''}" placeholder="ex: Suivi de dossier"></td>
      <td><input type="number" class="f-count" value="${item.count??''}"></td>
      <td><textarea class="f-desc" rows="1">${item.desc||''}</textarea></td>
      <td><button class="row-del">✕</button></td>
    `;
  }
  tr.querySelector(".row-del").addEventListener("click", ()=>{ tr.remove(); scheduleAnalysisAutosave(); });
  tr.querySelectorAll("input,textarea").forEach(inp=>inp.addEventListener("input", scheduleAnalysisAutosave));
  tbody.appendChild(tr);
}

function renderAnalysisForms(){
  const a = state.analysis;
  document.querySelector("#problemsTable tbody").innerHTML = "";
  document.querySelector("#demandesTable tbody").innerHTML = "";
  document.querySelector("#weeklyTable tbody").innerHTML = "";
  document.querySelector("#keywordsTable tbody").innerHTML = "";
  (a.problems||[]).forEach(it=>catRow("#problemsTable tbody","problems",it));
  (a.demandes||[]).forEach(it=>catRow("#demandesTable tbody","demandes",it));
  (a.weekly||[]).forEach(it=>catRow("#weeklyTable tbody","weekly",it));
  (a.keywords||[]).forEach(it=>catRow("#keywordsTable tbody","keywords",it));
  document.getElementById("constatsBox").value = (a.constats||[]).join("\n");
  document.getElementById("recoBox").value = (a.recommandations||[]).join("\n");
  document.getElementById("internalNotesBox").value = a.internal_notes || "";
}

document.querySelectorAll(".btn-add[data-add]").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    const kind = btn.dataset.add;
    if(kind === "problems") catRow("#problemsTable tbody","problems");
    if(kind === "demandes") catRow("#demandesTable tbody","demandes");
    if(kind === "weekly") catRow("#weeklyTable tbody","weekly");
    if(kind === "keywords") catRow("#keywordsTable tbody","keywords");
  });
});

document.getElementById("duplicatePrevBtn").addEventListener("click", ()=>{
  const pa = prevState.analysis || {};
  if(!(pa.problems||[]).length && !(pa.demandes||[]).length){
    toast("Aucune catégorie trouvée pour le mois précédent");
    return;
  }
  document.querySelector("#problemsTable tbody").innerHTML = "";
  document.querySelector("#demandesTable tbody").innerHTML = "";
  document.querySelector("#weeklyTable tbody").innerHTML = "";
  (pa.problems||[]).forEach(it=>catRow("#problemsTable tbody","problems",{label:it.label, count:"", desc:it.desc, action:it.action}));
  (pa.demandes||[]).forEach(it=>catRow("#demandesTable tbody","demandes",{label:it.label, count:"", desc:it.desc}));
  (pa.weekly||[]).forEach(it=>catRow("#weeklyTable tbody","weekly",{week:it.week, bugs:"", demandes:"", obs:""}));
  toast("Catégories dupliquées — pensez à mettre à jour les valeurs");
});

function attachSearch(inputId, tableSel){
  const input = document.getElementById(inputId);
  if(!input) return;
  input.addEventListener("input", ()=>{
    const q = input.value.trim().toLowerCase();
    document.querySelectorAll(tableSel + " tbody tr").forEach(tr=>{
      const label = tr.querySelector(".f-label")?.value.toLowerCase() || "";
      tr.style.display = label.includes(q) ? "" : "none";
    });
  });
}
attachSearch("searchProblems", "#problemsTable");
attachSearch("searchDemandes", "#demandesTable");

function collectCatTable(sel, withAction){
  const out = [];
  document.querySelectorAll(sel + " tbody tr").forEach(tr=>{
    const label = tr.querySelector(".f-label").value.trim();
    const count = Number(tr.querySelector(".f-count").value || 0);
    const desc = tr.querySelector(".f-desc").value.trim();
    if(!label) return;
    const item = {label, count, desc};
    if(withAction) item.action = tr.querySelector(".f-action").value.trim();
    out.push(item);
  });
  return out;
}

function collectWeeklyTable(){
  const out = [];
  document.querySelectorAll("#weeklyTable tbody tr").forEach(tr=>{
    const week = tr.querySelector(".f-week").value.trim();
    const bugs = Number(tr.querySelector(".f-bugs").value || 0);
    const demandes = Number(tr.querySelector(".f-demandes").value || 0);
    const obs = tr.querySelector(".f-obs").value.trim();
    if(week) out.push({week, bugs, demandes, obs});
  });
  return out;
}

function collectKeywordsTable(){
  const out = [];
  document.querySelectorAll("#keywordsTable tbody tr").forEach(tr=>{
    const word = tr.querySelector(".f-word").value.trim();
    const count = Number(tr.querySelector(".f-kwcount").value || 0);
    if(word) out.push({word, count});
  });
  return out;
}

function collectAnalysisPayload(){
  return {
    month_label: monthLabel(state.ym),
    problems: collectCatTable("#problemsTable", true),
    demandes: collectCatTable("#demandesTable", false),
    weekly: collectWeeklyTable(),
    keywords: collectKeywordsTable(),
    internal_notes: document.getElementById("internalNotesBox").value,
    constats: document.getElementById("constatsBox").value.split("\n").map(s=>s.trim()).filter(Boolean),
    recommandations: document.getElementById("recoBox").value.split("\n").map(s=>s.trim()).filter(Boolean),
  };
}

async function saveAnalysis(silent){
  const payload = collectAnalysisPayload();
  state.analysis = payload;
  await fetch(apiM(`/analysis/${state.ym}`), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  if(silent) pulseAutosave("analysisAutosaveHint");
}

const scheduleAnalysisAutosave = debounce(()=>saveAnalysis(true), 1800);
document.getElementById("constatsBox").addEventListener("input", scheduleAnalysisAutosave);
document.getElementById("recoBox").addEventListener("input", scheduleAnalysisAutosave);
document.getElementById("internalNotesBox").addEventListener("input", scheduleAnalysisAutosave);

document.getElementById("saveAnalysisBtn").addEventListener("click", async ()=>{
  await saveAnalysis(false);
  saveIndicatorPulse();
  toast("Analyse enregistrée");
});

// ============================================================
// DASHBOARD
// ============================================================
function sumField(obj, field){
  return Object.values(obj).reduce((s,d)=>s+Number(d[field]||0),0);
}

function computeMonthMetrics(calls, emails, analysis){
  const problems = analysis.problems || [];
  const demandes = analysis.demandes || [];
  const durationSec = Object.values(calls).reduce((s,d)=>s+getDurationSec(d),0);
  return {
    received: sumField(emails, "received"),
    sent: sumField(emails, "sent"),
    calls: sumField(calls, "calls"),
    durationSec,
    bugs: problems.reduce((s,p)=>s+Number(p.count||0),0),
    demandes: demandes.reduce((s,d)=>s+Number(d.count||0),0),
  };
}

// polarity: "positive" (hausse = bien, vert), "negative" (hausse = mal, rouge), "neutral" (gris, pas de jugement)
function kpiDelta(current, previous, polarity="neutral"){
  if(previous === 0 && current === 0) return "";
  let txt, color;
  if(previous === 0){
    txt = "nouveau"; color = "var(--muted)";
  } else {
    const diff = ((current-previous)/previous)*100;
    const sign = diff >= 0 ? "+" : "";
    txt = `${sign}${diff.toFixed(0)}% vs préc.`;
    if(Math.abs(diff) < 1 || polarity === "neutral"){ color = "var(--muted)"; }
    else {
      const isUp = diff > 0;
      const isGood = polarity === "positive" ? isUp : !isUp;
      color = isGood ? "var(--success)" : "var(--danger)";
    }
  }
  return `<div class="kpi-delta" style="color:${color};">${txt}</div>`;
}

function compareChip(label, current, previous){
  if(previous === 0 && current === 0) return "";
  let cls = "flat", arrow = "→", pct = "0%";
  if(previous === 0){ cls="up"; arrow="↑"; pct="nouveau"; }
  else {
    const diff = ((current-previous)/previous)*100;
    if(Math.abs(diff) < 1){ cls="flat"; arrow="→"; pct="stable"; }
    else if(diff > 0){ cls="up"; arrow="↑"; pct=`+${diff.toFixed(0)}%`; }
    else { cls="down"; arrow="↓"; pct=`${diff.toFixed(0)}%`; }
  }
  return `<div class="compare-chip">${label} <span class="${cls}">${arrow} ${pct}</span> <span style="color:var(--muted);">vs ${monthLabel(prevYm(state.ym))}</span></div>`;
}

// Chip générique pour un comparatif à référence libre (utilisé pour N vs N-1)
function metricCompareChip(label, current, previous, refLabel){
  let cls = "flat", arrow = "→", pct = "0%";
  if(previous === 0 && current === 0){ pct = "—"; }
  else if(previous === 0){ cls="up"; arrow="↑"; pct="nouveau"; }
  else {
    const diff = ((current-previous)/previous)*100;
    if(Math.abs(diff) < 1){ cls="flat"; arrow="→"; pct="stable"; }
    else if(diff > 0){ cls="up"; arrow="↑"; pct=`+${diff.toFixed(0)}%`; }
    else { cls="down"; arrow="↓"; pct=`${diff.toFixed(0)}%`; }
  }
  return `<div class="yoy-chip">
    <div class="yoy-chip-label">${label}</div>
    <div class="yoy-chip-values"><strong>${current}</strong> <span style="color:var(--muted);">(${previous})</span></div>
    <div class="compare-chip"><span class="${cls}">${arrow} ${pct}</span> <span style="color:var(--muted);">${refLabel}</span></div>
  </div>`;
}

function groupDailyIntoWeeks(ym, calls, emails){
  const days = businessDaysOfMonth(ym);
  const buckets = [[],[],[],[],[]]; // jusqu'à 5 semaines de 7 jours calendaires
  days.forEach(d=>{
    const dayNum = Number(d.date.split("-")[2]);
    const idx = Math.min(4, Math.floor((dayNum-1)/7));
    buckets[idx].push(d);
  });
  return buckets.filter(b=>b.length).map((b,i)=>{
    const recv = b.reduce((s,d)=>s + Number((emails[d.date]||{}).received||0), 0);
    const sent = b.reduce((s,d)=>s + Number((emails[d.date]||{}).sent||0), 0);
    return { label: `Sem. ${i+1}`, recv, sent };
  });
}

function renderDashboard(){
  const info = modulesInfo[currentModule] || {label: currentModule};
  const a = state.analysis || {};
  const problems = a.problems || [];
  const demandes = a.demandes || [];
  const weekly = a.weekly || [];
  const keywords = a.keywords || [];

  const cur = computeMonthMetrics(state.calls, state.emails, a);
  const prev = computeMonthMetrics(prevState.calls, prevState.emails, prevState.analysis || {});
  const yoy = computeMonthMetrics(yoyState.calls, yoyState.emails, yoyState.analysis || {});
  const showYoy = currentUser && (currentUser.role === "admin" || currentUser.role === "superviseur");
  const ratio = cur.demandes ? (cur.bugs/cur.demandes) : 0;
  const avgCallSec = cur.calls ? cur.durationSec/cur.calls : 0;
  const channelTotal = cur.calls + cur.received;
  const channelCallsPct = channelTotal ? Math.round(cur.calls/channelTotal*100) : 0;
  const channelEmailPct = 100 - channelCallsPct;

  const tarkhissAlerts = computeAlerts(
    {backlog: cur.bugs, delayH: avgCallSec/3600, volumeCur: cur.received, volumePrev: prev.received},
    moduleSettings.alert_thresholds
  );
  renderAlertBanner("tarkhissAlertBanner", tarkhissAlerts);
  logAlertHistory("tarkhiss", state.ym, tarkhissAlerts);

  let critWeek = null;
  weekly.forEach(w=>{
    const t = Number(w.bugs||0)+Number(w.demandes||0);
    if(!critWeek || t > (Number(critWeek.bugs||0)+Number(critWeek.demandes||0))) critWeek = w;
  });

  document.getElementById("dashTitle").textContent = `Dashboard — ${info.label} — ${monthLabel(state.ym)}`;

  const logoImg = globalSettings.logo_filename ? `<img src="/static/uploads/${globalSettings.logo_filename}" class="print-only" style="height:34px;margin-bottom:10px;">` : "";

  const maxKw = Math.max(1, ...keywords.map(k=>k.count||0));
  const wordCloudHtml = keywords.length ? `
    <div class="dash-table-wrap">
      <h3>☁️ Nuage de mots-clés</h3>
      <div class="word-cloud">
        ${keywords.map(k=>{
          const size = 12 + (k.count/maxKw)*22;
          const opacity = 0.55 + (k.count/maxKw)*0.45;
          return `<span style="font-size:${size}px;opacity:${opacity};">${k.word}</span>`;
        }).join("")}
      </div>
    </div>` : "";

  const html = `
    <div class="print-only" style="margin-bottom:10px;">
      ${logoImg}
      <div style="font-size:11px;color:#55616B;">${globalSettings.agency_name || ""} · ${info.label} · ${monthLabel(state.ym)}</div>
    </div>

    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">📧 Emails reçus</div><div class="kpi-value">${cur.received}</div>${kpiDelta(cur.received, prev.received, "neutral")}</div>
      <div class="kpi-card"><div class="kpi-label">📤 Emails envoyés</div><div class="kpi-value">${cur.sent}</div>${kpiDelta(cur.sent, prev.sent, "neutral")}</div>
      <div class="kpi-card"><div class="kpi-label">📞 Total appels</div><div class="kpi-value">${cur.calls}</div>${kpiDelta(cur.calls, prev.calls, "neutral")}</div>
      <div class="kpi-card"><div class="kpi-label">⏱️ Temps comm. (hh:mm:ss)</div><div class="kpi-value" style="font-size:20px;">${formatHMS(cur.durationSec)}</div></div>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card danger"><div class="kpi-label">🔴 Problèmes techniques</div><div class="kpi-value">${cur.bugs}</div>${kpiDelta(cur.bugs, prev.bugs, "negative")}</div>
      <div class="kpi-card success"><div class="kpi-label">🟢 Demandes d'info</div><div class="kpi-value">${cur.demandes}</div>${kpiDelta(cur.demandes, prev.demandes, "neutral")}</div>
      <div class="kpi-card"><div class="kpi-label">Ratio Bugs / Demandes</div><div class="kpi-value">${ratio.toFixed(2)}</div></div>
      <div class="kpi-card accent"><div class="kpi-label">⚠️ Semaine critique</div><div class="kpi-value" style="font-size:16px;">${critWeek ? critWeek.week : "—"}</div></div>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">⏳ Durée moy./appel</div><div class="kpi-value" style="font-size:20px;">${formatHMS(avgCallSec)}</div></div>
      <div class="kpi-card"><div class="kpi-label">📊 Canal dominant</div><div class="kpi-value" style="font-size:15px;">☎ ${channelCallsPct}% / ✉ ${channelEmailPct}%</div></div>
    </div>

    ${showYoy ? `
    <div class="dash-table-wrap yoy-card">
      <h3>📆 Comparatif ${monthLabel(state.ym)} vs ${monthLabel(prevYearYm(state.ym))} (N-1)</h3>
      <div class="yoy-grid">
        ${metricCompareChip("Emails reçus", cur.received, yoy.received, "vs N-1")}
        ${metricCompareChip("Appels", cur.calls, yoy.calls, "vs N-1")}
        ${metricCompareChip("Problèmes tech.", cur.bugs, yoy.bugs, "vs N-1")}
        ${metricCompareChip("Demandes info", cur.demandes, yoy.demandes, "vs N-1")}
      </div>
    </div>` : ""}

    <div class="dash-table-wrap" id="emailMetaStatsCard" style="display:none;">
      <h3>📧 Statistiques email — historique complet</h3>
      <p style="font-size:11.5px;color:var(--muted);margin:0 0 8px;">Basé sur le dernier import Outlook (tous mois confondus).</p>
      <div class="yoy-grid" id="emailMetaStatsGrid"></div>
    </div>

    <div class="dash-table-wrap" id="callMetaStatsCard" style="display:none;">
      <h3>📞 Statistiques appels — historique complet</h3>
      <p style="font-size:11.5px;color:var(--muted);margin:0 0 8px;">Basé sur le dernier import du journal d'appels (tous mois confondus).</p>
      <div class="yoy-grid" id="callMetaStatsGrid"></div>
    </div>

    <div class="charts-row" id="hourlyHeatmapsRow" style="display:none;">
      <div class="chart-card">
        <h3>Appels reçus — jour × heure</h3>
        <p style="font-size:10.5px;color:var(--muted);margin:-6px 0 8px;">Jours ouvrés, 08h-17h (horaires Tarkhiss)</p>
        <img id="callHourlyHeatmapImg" style="width:100%;display:block;" alt="Heatmap horaire des appels">
      </div>
      <div class="chart-card">
        <h3>Emails reçus — jour × heure</h3>
        <p style="font-size:10.5px;color:var(--muted);margin:-6px 0 8px;">Jours ouvrés, 08h-17h (horaires Tarkhiss)</p>
        <img id="emailHourlyHeatmapImg" style="width:100%;display:block;" alt="Heatmap horaire des emails">
      </div>
    </div>

    <div class="dash-table-wrap" id="responseDelayCard" style="display:none;">
      <h3>⏱ Délai de première réponse (email, indicatif)</h3>
      <p style="font-size:11px;color:var(--muted);margin:0 0 8px;" id="responseDelayNote"></p>
      <div class="yoy-grid" id="responseDelayGrid"></div>
    </div>

    <div class="charts-row">
      <div class="chart-card">
        <h3>Évolution hebdomadaire — Bugs vs Demandes</h3>
        <canvas id="weeklyChart" height="140"></canvas>
      </div>
      <div class="chart-card">
        <h3>Répartition des problèmes techniques</h3>
        <canvas id="problemsChart" height="140"></canvas>
      </div>
    </div>

    <div class="charts-row">
      <div class="chart-card">
        <h3>Charge quotidienne (appels + emails reçus)</h3>
        <canvas id="dailyLoadChart" height="140"></canvas>
      </div>
      <div class="chart-card">
        <h3>Emails reçus/envoyés par semaine</h3>
        <canvas id="weeklyEmailsChart" height="140"></canvas>
      </div>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Heatmap de charge (jour × semaine)</h3>
      <p style="font-size:11.5px;color:var(--muted);margin:0 0 8px;">Granularité journalière — Tarkhiss n'enregistre pas l'heure des appels.</p>
      <img src="/api/tarkhiss/heatmap-png/${state.ym}?_=${Date.now()}" style="width:100%;max-width:720px;display:block;margin:0 auto;" alt="Heatmap de charge Tarkhiss">
    </div>

    ${wordCloudHtml}

    <div class="dash-table-wrap">
      <h3>🔴 Problèmes techniques signalés (${cur.bugs})</h3>
      <table class="dash-table">
        <thead><tr><th>Type</th><th>Nb</th><th>%</th><th>Description</th><th>Action corrective</th></tr></thead>
        <tbody>
          ${problems.map(p=>{
            const pct = cur.bugs ? (p.count/cur.bugs*100) : 0;
            return `<tr>
              <td>${esc(p.label)}</td><td><strong>${p.count}</strong></td>
              <td><span class="pct-bar-bg"><span class="pct-bar-fill" style="width:${pct}%;background:var(--danger);"></span></span>${pct.toFixed(1)}%</td>
              <td style="color:var(--muted);font-size:12.5px;">${esc(p.desc)||''}</td>
              <td style="color:var(--muted);font-size:12.5px;">${esc(p.action)||''}</td>
            </tr>`;
          }).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucune donnée — à saisir dans l'onglet Analyse mensuelle</td></tr>`}
        </tbody>
      </table>
    </div>

    <div class="dash-table-wrap">
      <h3>🟢 Demandes d'information reçues (${cur.demandes})</h3>
      <table class="dash-table">
        <thead><tr><th>Type</th><th>Nb</th><th>%</th><th>Description</th></tr></thead>
        <tbody>
          ${demandes.map(d=>{
            const pct = cur.demandes ? (d.count/cur.demandes*100) : 0;
            return `<tr>
              <td>${esc(d.label)}</td><td><strong>${d.count}</strong></td>
              <td><span class="pct-bar-bg"><span class="pct-bar-fill" style="width:${pct}%;background:var(--success);"></span></span>${pct.toFixed(1)}%</td>
              <td style="color:var(--muted);font-size:12.5px;">${esc(d.desc)||''}</td>
            </tr>`;
          }).join("") || `<tr><td colspan="4" style="color:var(--muted);">Aucune donnée — à saisir dans l'onglet Analyse mensuelle</td></tr>`}
        </tbody>
      </table>
    </div>

    <div class="dash-table-wrap">
      <h3>📅 Évolution hebdomadaire</h3>
      <table class="dash-table">
        <thead><tr><th>Semaine</th><th>Bugs</th><th>Demandes</th><th>Observations</th></tr></thead>
        <tbody>
          ${weekly.map(w=>{
            const isCrit = critWeek && w.week===critWeek.week;
            return `<tr class="${isCrit?'critical':''}">
              <td>${w.week}${isCrit?' ⚠️':''}</td><td style="color:var(--danger);font-weight:700;">${w.bugs}</td>
              <td style="color:var(--success);font-weight:700;">${w.demandes}</td>
              <td style="color:var(--muted);font-size:12.5px;">${w.obs||''}</td>
            </tr>`;
          }).join("") || `<tr><td colspan="4" style="color:var(--muted);">Aucune donnée</td></tr>`}
        </tbody>
      </table>
    </div>

    <div class="synth-cols">
      <div class="dash-table-wrap synth-card constats">
        <h3>🔑 Constats clés</h3>
        <ul>${(a.constats||[]).map(c=>`<li>${c}</li>`).join("") || '<li style="color:var(--muted);">—</li>'}</ul>
      </div>
      <div class="dash-table-wrap synth-card reco">
        <h3>✅ Recommandations</h3>
        <ul>${(a.recommandations||[]).map(c=>`<li>${c}</li>`).join("") || '<li style="color:var(--muted);">—</li>'}</ul>
      </div>
    </div>
  `;
  document.getElementById("dashboardContent").innerHTML = html;

  Object.values(charts).forEach(c=>c && c.destroy());
  [dailyLoadChart, weeklyEmailsChart].forEach(c=>c && c.destroy());

  const ctxW = document.getElementById("weeklyChart");
  charts.weekly = new Chart(ctxW, {
    type: "bar",
    data: {
      labels: weekly.map(w=>w.week),
      datasets: [
        { label:"Bugs", data: weekly.map(w=>w.bugs), backgroundColor:"#DC2828", borderRadius:5 },
        { label:"Demandes", data: weekly.map(w=>w.demandes), backgroundColor:"#25935F", borderRadius:5 },
      ]
    },
    options: { responsive:true, animation:{duration:700, easing:'easeOutQuart'}, plugins:{legend:{position:"bottom"}}, scales:{ x:{grid:{display:false}}, y:{beginAtZero:true} } }
  });

  const ctxP = document.getElementById("problemsChart");
  const palette = ["#0B4965","#DC2828","#F59F0A","#25935F","#55616B","#8E5AA6","#1794CF"];
  charts.problems = new Chart(ctxP, {
    type: "doughnut",
    data: {
      labels: problems.map(p=>p.label),
      datasets: [{ data: problems.map(p=>p.count), backgroundColor: problems.map((_,i)=>palette[i%palette.length]) }]
    },
    options: { responsive:true, animation:{duration:700, easing:'easeOutQuart'}, plugins:{legend:{position:"bottom", labels:{boxWidth:10, font:{size:11}}}} }
  });

  const days = businessDaysOfMonth(state.ym);
  const dailyLoadCtx = document.getElementById("dailyLoadChart");
  dailyLoadChart = new Chart(dailyLoadCtx, {
    type: "line",
    data: {
      labels: days.map(d=>d.date.split("-")[2]),
      datasets: [{
        label: "Charge (appels + emails reçus)",
        data: days.map(d=> Number((state.calls[d.date]||{}).calls||0) + Number((state.emails[d.date]||{}).received||0) ),
        borderColor:"#0B4965", backgroundColor:"#0B496522", fill:true, tension:.3
      }]
    },
    options: { responsive:true, animation:{duration:700}, plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}} }
  });

  const weeklyEmails = groupDailyIntoWeeks(state.ym, state.calls, state.emails);
  const weeklyEmailsCtx = document.getElementById("weeklyEmailsChart");
  weeklyEmailsChart = new Chart(weeklyEmailsCtx, {
    type: "bar",
    data: {
      labels: weeklyEmails.map(w=>w.label),
      datasets: [
        { label:"Reçus", data: weeklyEmails.map(w=>w.recv), backgroundColor:"#0B4965", stack:"e" },
        { label:"Envoyés", data: weeklyEmails.map(w=>w.sent), backgroundColor:"#F59F0A", stack:"e" },
      ]
    },
    options: { responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}}, scales:{ x:{stacked:true, grid:{display:false}}, y:{stacked:true, beginAtZero:true} } }
  });

  setupEmailPanel();
  addPngExportButtons(charts);
  refreshEmailMetaStats();
  refreshCallMetaStats();
  refreshHourlyHeatmaps();
  refreshResponseDelay();
}

async function refreshCallMetaStats(){
  const card = document.getElementById("callMetaStatsCard");
  const grid = document.getElementById("callMetaStatsGrid");
  if(!card || !grid) return;
  try{
    const res = await fetch("/api/tarkhiss/call-meta-stats");
    const s = await res.json();
    if(!s || !s.total_in){ card.style.display = "none"; return; }
    grid.innerHTML = `
      <div class="yoy-chip"><div class="yoy-chip-label">Total reçus (historique)</div><div class="yoy-chip-values">${s.total_in}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Total manqués</div><div class="yoy-chip-values">${s.total_missed}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Total émis</div><div class="yoy-chip-values">${s.total_out}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Taux de décroché</div><div class="yoy-chip-values">${s.answer_rate!=null?s.answer_rate+"%":"—"}</div></div>
    `;
    card.style.display = "";
  } catch(e){
    card.style.display = "none";
  }
}

async function refreshHourlyHeatmaps(){
  const row = document.getElementById("hourlyHeatmapsRow");
  const callImg = document.getElementById("callHourlyHeatmapImg");
  const emailImg = document.getElementById("emailHourlyHeatmapImg");
  if(!row) return;
  const [callOk, emailOk] = await Promise.all([
    fetch(`/api/tarkhiss/call-meta/${state.ym}`).then(r=>r.ok?r.json():null).then(d=>d && Object.keys(d).length>0),
    fetch(`/api/tarkhiss/email-meta/${state.ym}`).then(r=>r.ok?r.json():null).then(d=>d && Object.keys(d).length>0),
  ]);
  row.style.display = (callOk || emailOk) ? "" : "none";
  row.children[0].style.display = callOk ? "" : "none";
  row.children[1].style.display = emailOk ? "" : "none";
  if(callOk) callImg.src = `/api/tarkhiss/call-heatmap-hourly-png/${state.ym}?_=${Date.now()}`;
  if(emailOk) emailImg.src = `/api/tarkhiss/email-heatmap-hourly-png/${state.ym}?_=${Date.now()}`;
}

async function refreshResponseDelay(){
  const card = document.getElementById("responseDelayCard");
  const grid = document.getElementById("responseDelayGrid");
  const note = document.getElementById("responseDelayNote");
  if(!card || !grid) return;
  try{
    const res = await fetch(`/api/tarkhiss/email-meta/${state.ym}`);
    const m = await res.json();
    const r = m && m.response;
    if(!r || !r.replied){ card.style.display = "none"; return; }
    const fmtMin = v => v==null ? "—" : (v>=60 ? `${Math.floor(v/60)} h ${String(Math.round(v%60)).padStart(2,"0")}` : `${Math.round(v)} min`);
    grid.innerHTML = `
      <div class="yoy-chip"><div class="yoy-chip-label">Délai médian (heures ouvrées)</div><div class="yoy-chip-values">${fmtMin(r.median_biz_min)}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Délai P90</div><div class="yoy-chip-values">${fmtMin(r.p90_biz_min)}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Sans réponse sous 7 j</div><div class="yoy-chip-values">${m.unanswered}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Taux d'appariement</div><div class="yoy-chip-values">${r.match_rate!=null?r.match_rate+"%":"—"}</div></div>
    `;
    note.textContent = `${r.note} Base horaire : ${r.business_hours}.`;
    card.style.display = "";
  } catch(e){
    card.style.display = "none";
  }
}

async function refreshEmailMetaStats(){
  const card = document.getElementById("emailMetaStatsCard");
  const grid = document.getElementById("emailMetaStatsGrid");
  if(!card || !grid) return;
  try{
    const res = await fetch("/api/tarkhiss/email-meta-stats");
    const s = await res.json();
    if(!s || !s.total_received){ card.style.display = "none"; return; }
    const days = ["Lun","Mar","Mer","Jeu","Ven","Sam","Dim"];
    const topDay = s.weekday_counts ? days[s.weekday_counts.indexOf(Math.max(...s.weekday_counts))] : "—";
    grid.innerHTML = `
      <div class="yoy-chip"><div class="yoy-chip-label">Total reçus (historique)</div><div class="yoy-chip-values">${s.total_received}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Total envoyés (historique)</div><div class="yoy-chip-values">${s.total_sent}</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Avec pièce jointe</div><div class="yoy-chip-values">${s.pct_with_attachment}%</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Importance haute</div><div class="yoy-chip-values">${s.pct_high_importance}%</div></div>
      <div class="yoy-chip"><div class="yoy-chip-label">Jour le plus chargé</div><div class="yoy-chip-values">${topDay}</div></div>
    `;
    card.style.display = "";
  } catch(e){
    card.style.display = "none";
  }
}

// ---------- ONE-PAGER (image PNG condensée partageable) ----------
function roundRect(ctx, x, y, w, h, r){
  ctx.beginPath();
  ctx.moveTo(x+r, y);
  ctx.arcTo(x+w, y, x+w, y+h, r);
  ctx.arcTo(x+w, y+h, x, y+h, r);
  ctx.arcTo(x, y+h, x, y, r);
  ctx.arcTo(x, y, x+w, y, r);
  ctx.closePath();
}

function buildOnePagerCanvas(title, subtitle, kpis, chartCanvases){
  const W = 900;
  const cardH = 92, cardGap = 14, cols = 4;
  const kpiRows = Math.ceil(kpis.length / cols);
  const chartsH = chartCanvases.length ? 340*Math.ceil(chartCanvases.length/2) + 20 : 0;
  const H = 150 + kpiRows*(cardH+cardGap) + chartsH + 60;

  const canvas = document.createElement("canvas");
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#F5F7F9";
  ctx.fillRect(0,0,W,H);

  ctx.fillStyle = "#0B4965";
  ctx.fillRect(0,0,W,86);
  ctx.fillStyle = "#FFFFFF";
  ctx.font = "bold 24px Arial";
  ctx.fillText(title, 28, 40);
  ctx.font = "13px Arial";
  ctx.fillStyle = "#B9D3E0";
  ctx.fillText(subtitle, 28, 64);
  if(globalSettings.agency_name){
    ctx.textAlign = "right";
    ctx.fillText(globalSettings.agency_name, W-28, 64);
    ctx.textAlign = "left";
  }

  let y = 110;
  kpis.forEach((k,i)=>{
    const col = i % cols, row = Math.floor(i/cols);
    const cw = (W - 28*2 - cardGap*(cols-1)) / cols;
    const x = 28 + col*(cw+cardGap);
    const cy = y + row*(cardH+cardGap);
    ctx.fillStyle = "#FFFFFF";
    roundRect(ctx, x, cy, cw, cardH, 8); ctx.fill();
    ctx.strokeStyle = "#DAE0E7"; ctx.lineWidth = 1;
    roundRect(ctx, x, cy, cw, cardH, 8); ctx.stroke();
    ctx.fillStyle = "#55616B";
    ctx.font = "11px Arial";
    ctx.fillText(k.label, x+12, cy+22);
    ctx.fillStyle = k.color || "#0B4965";
    ctx.font = "bold 26px Arial";
    ctx.fillText(String(k.value), x+12, cy+56);
    if(k.delta){
      ctx.fillStyle = k.deltaColor || "#55616B";
      ctx.font = "bold 11px Arial";
      ctx.fillText(k.delta, x+12, cy+76);
    }
  });

  let chartY = y + kpiRows*(cardH+cardGap) + 20;
  chartCanvases.forEach((c,i)=>{
    if(!c) return;
    const col = i % 2, row = Math.floor(i/2);
    const cw = (W - 28*2 - 14) / 2;
    const ch = 320;
    const x = 28 + col*(cw+14);
    const cy = chartY + row*(ch+20);
    ctx.fillStyle = "#FFFFFF";
    roundRect(ctx, x, cy, cw, ch, 8); ctx.fill();
    ctx.strokeStyle = "#DAE0E7"; roundRect(ctx, x, cy, cw, ch, 8); ctx.stroke();
    const ratio = Math.min((cw-16)/c.width, (ch-16)/c.height);
    const dw = c.width*ratio, dh = c.height*ratio;
    ctx.drawImage(c, x + (cw-dw)/2, cy + (ch-dh)/2, dw, dh);
  });

  ctx.fillStyle = "#55616B";
  ctx.font = "10px Arial";
  ctx.fillText(`Généré le ${new Date().toLocaleDateString('fr-FR')}`, 28, H-16);

  return canvas;
}

function downloadCanvas(canvas, filename){
  const a = document.createElement("a");
  a.href = canvas.toDataURL("image/png");
  a.download = filename;
  a.click();
}

// ---------- PDF / Excel ----------
let _pdfPreviewBlobUrl = null;
async function downloadPdfFromServer(url, payload, filename){
  showBusy("Génération du PDF en cours...");
  let res;
  try{
    res = await fetch(url, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  } finally {
    hideBusy();
  }
  if(!res.ok){ toast("Erreur génération PDF"); return; }
  const blob = await res.blob();
  if(_pdfPreviewBlobUrl) URL.revokeObjectURL(_pdfPreviewBlobUrl);
  _pdfPreviewBlobUrl = URL.createObjectURL(blob);
  document.getElementById("pdfPreviewFrame").src = _pdfPreviewBlobUrl;
  document.getElementById("pdfPreviewName").textContent = filename;
  document.getElementById("pdfPreviewDownloadBtn").onclick = ()=>{
    const a = document.createElement("a");
    a.href = _pdfPreviewBlobUrl;
    a.download = filename;
    a.click();
  };
  document.getElementById("pdfPreviewModalOverlay").style.display = "flex";
}

document.getElementById("closePdfPreviewBtn").addEventListener("click", ()=>{
  document.getElementById("pdfPreviewModalOverlay").style.display = "none";
  document.getElementById("pdfPreviewFrame").src = "";
});
document.getElementById("pdfPreviewModalOverlay").addEventListener("click", (e)=>{
  if(e.target.id === "pdfPreviewModalOverlay"){
    document.getElementById("pdfPreviewModalOverlay").style.display = "none";
    document.getElementById("pdfPreviewFrame").src = "";
  }
});

document.getElementById("pdfBtn").addEventListener("click", ()=>{
  downloadPdfFromServer(apiM(`/export-pdf/${state.ym}`), {charts: collectTarkhissChartImages()}, `rapport_tarkhiss_${state.ym}.pdf`);
});
document.getElementById("onePagerBtn").addEventListener("click", ()=>{
  const cur = computeMonthMetrics(state.calls, state.emails, state.analysis||{});
  const kpis = [
    {label:"EMAILS REÇUS", value: cur.received},
    {label:"APPELS", value: cur.calls},
    {label:"PROBLÈMES TECH.", value: cur.bugs, color:"#DC2828"},
    {label:"DEMANDES INFO", value: cur.demandes, color:"#25935F"},
  ];
  const canvas = buildOnePagerCanvas(
    `Support Tarkhiss — ${monthLabel(state.ym)}`,
    "Synthèse mensuelle",
    kpis,
    [charts.weekly?.canvas, charts.problems?.canvas]
  );
  downloadCanvas(canvas, `onepager_tarkhiss_${state.ym}.png`);
});
document.getElementById("xlsxBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-xlsx/${state.ym}`);
});

// ---------- Panneau d'envoi email (générique, réutilisé par tous les volets) ----------
function buildGreetingSuggestion(toListId="toCheckList"){
  const toContacts = (moduleSettings.contacts||[]).filter(c=>document.querySelector(`#${toListId} input[value="${c.email}"]`)?.checked);
  const names = toContacts.map(c=>c.name).filter(Boolean);
  if(!names.length) return moduleSettings.greeting || "Bonjour,";
  if(names.length === 1) return `Bonjour ${names[0]},`;
  return `Bonjour ${names.slice(0,-1).join(", ")} et ${names[names.length-1]},`;
}

function renderContactCheckList(toListId="toCheckList", ccListId="ccCheckList", greetingId="greetingInput"){
  const contacts = moduleSettings.contacts || [];
  const toList = document.getElementById(toListId);
  const ccList = document.getElementById(ccListId);
  toList.innerHTML = contacts.map(c=>`
    <label class="contact-check-item"><input type="checkbox" value="${esc(c.email)}" ${c.role==='to'?'checked':''} data-kind="to"> ${c.name ? esc(c.name)+' — ' : ''}${esc(c.email)}</label>
  `).join("") || `<span class="hint">Aucun contact — ajoutez-en dans Administration</span>`;
  ccList.innerHTML = contacts.map(c=>`
    <label class="contact-check-item"><input type="checkbox" value="${esc(c.email)}" ${c.role==='cc'?'checked':''} data-kind="cc"> ${c.name ? esc(c.name)+' — ' : ''}${esc(c.email)}</label>
  `).join("") || `<span class="hint">Aucun contact</span>`;

  document.querySelectorAll(`#${toListId} input`).forEach(cb=>{
    cb.addEventListener("change", ()=>{ document.getElementById(greetingId).value = buildGreetingSuggestion(toListId); });
  });
}

function setupEmailPanel(toListId="toCheckList", ccListId="ccCheckList", greetingId="greetingInput"){
  renderContactCheckList(toListId, ccListId, greetingId);
  document.getElementById(greetingId).value = moduleSettings.greeting || buildGreetingSuggestion(toListId);
}

document.getElementById("mailToggleBtn").addEventListener("click", ()=>{
  const panel = document.getElementById("emailPanel");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
});

function getSelectedRecipients(toListId="toCheckList", ccListId="ccCheckList"){
  const to = [...document.querySelectorAll(`#${toListId} input:checked`)].map(i=>i.value);
  const cc = [...document.querySelectorAll(`#${ccListId} input:checked`)].map(i=>i.value);
  return {to, cc};
}

function openPreviewModal(subject, html){
  document.getElementById("previewSubject").textContent = subject;
  document.getElementById("previewFrame").srcdoc = html;
  document.getElementById("previewModalOverlay").style.display = "flex";
}

function collectTarkhissChartImages(){
  const map = {weekly:"weekly", problems:"problems"};
  const out = {};
  Object.entries(map).forEach(([key, chartKey])=>{
    const chart = charts[chartKey];
    if(chart && chart.toBase64Image){
      try{ out[key] = chart.toBase64Image('image/png', 1); }catch(e){}
    }
  });
  return out;
}

document.getElementById("previewMailBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients();
  const greeting = document.getElementById("greetingInput").value.trim();
  const chartsData = collectTarkhissChartImages();
  const res = await fetch(apiM(`/preview-email/${state.ym}`), {
    method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({to, cc, greeting, charts: chartsData})
  });
  const data = await res.json();
  openPreviewModal(data.subject, data.html);
});

document.getElementById("closePreviewBtn").addEventListener("click", ()=>{
  document.getElementById("previewModalOverlay").style.display = "none";
});
document.getElementById("previewModalOverlay").addEventListener("click", (e)=>{
  if(e.target.id === "previewModalOverlay") document.getElementById("previewModalOverlay").style.display = "none";
});

document.getElementById("confirmSendBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients();
  if(!to.length){ toast("Sélectionnez au moins un destinataire (À)"); return; }
  const greeting = document.getElementById("greetingInput").value.trim();
  const chartsData = collectTarkhissChartImages();
  toast("Génération du rapport...");
  const res = await fetch(apiM(`/send-email/${state.ym}`), {
    method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({to, cc, greeting, charts: chartsData})
  });
  const data = await res.json();
  toast(data.opened ? "Ouverture dans Outlook..." : "Fichier .eml généré : " + data.file);
  document.getElementById("emailPanel").style.display = "none";
  checkReminder();
});

// ============================================================
// VUE ANNUELLE
// ============================================================
function pad2(n){ return n.toString().padStart(2,"0"); }
function ymOf(date){ return `${date.getFullYear()}-${pad2(date.getMonth()+1)}`; }

function periodPresetRanges(){
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth(); // 0-indexed
  const qStartMonth = Math.floor(m/3)*3;
  const hStartMonth = Math.floor(m/6)*6;
  const twelveAgo = new Date(y, m-11, 1);
  return {
    quarter: { start: ymOf(new Date(y, qStartMonth, 1)), end: ymOf(new Date(y, qStartMonth+2, 1)), label: "Ce trimestre" },
    semester: { start: ymOf(new Date(y, hStartMonth, 1)), end: ymOf(new Date(y, hStartMonth+5, 1)), label: "Ce semestre" },
    year: { start: `${y}-01`, end: `${y}-12`, label: "Cette année" },
    last12: { start: ymOf(twelveAgo), end: ymOf(now), label: "12 derniers mois" },
  };
}

function setupPeriodPicker(presetsId, startId, endId, onChange){
  const presets = periodPresetRanges();
  const wrap = document.getElementById(presetsId);
  const startInput = document.getElementById(startId);
  const endInput = document.getElementById(endId);

  wrap.innerHTML = Object.entries(presets).map(([key,p])=>`<button data-key="${key}">${p.label}</button>`).join("");

  function applyPreset(key){
    const p = presets[key];
    startInput.value = p.start;
    endInput.value = p.end;
    wrap.querySelectorAll("button").forEach(b=>b.classList.toggle("active", b.dataset.key===key));
    onChange();
  }
  wrap.querySelectorAll("button").forEach(b=>{
    b.addEventListener("click", ()=>applyPreset(b.dataset.key));
  });
  [startInput, endInput].forEach(inp=>{
    inp.addEventListener("change", ()=>{
      wrap.querySelectorAll("button").forEach(b=>b.classList.remove("active"));
      onChange();
    });
  });
  applyPreset("year");
}

const HEATMAP_COLORS = ["#E7F5EE","#CFEBDD","#6FC49A","#25935F","#F59F0A","#DC2828","#0B4965"];
function heatColor(avg, maxAvg){
  if(!maxAvg) return "#E7F5EE";
  const ratio = avg / maxAvg;
  if(ratio < 0.15) return "#E7F5EE";
  if(ratio < 0.35) return "#CFEBDD";
  if(ratio < 0.55) return "#6FC49A";
  if(ratio < 0.75) return "#25935F";
  if(ratio < 0.9) return "#F59F0A";
  return "#DC2828";
}

async function loadAnnual(){
  const start = document.getElementById("periodStart").value || `${new Date().getFullYear()}-01`;
  const end = document.getElementById("periodEnd").value || `${new Date().getFullYear()}-12`;
  const year = start.split("-")[0];
  const res = await fetch(apiM(`/annual/${year}?start=${start}&end=${end}`));
  const data = await res.json();
  const months = data.months;
  const periodLabel = `${monthLabel(start)} — ${monthLabel(end)}`;

  const topHtml = data.top_categories.length ? `
    <div class="dash-table-wrap">
      <h3>🏆 Top 3 catégories récurrentes sur la période</h3>
      <table class="dash-table">
        <thead><tr><th>Catégorie</th><th>Total signalé sur la période</th></tr></thead>
        <tbody>
          ${data.top_categories.map((c,i)=>`<tr><td>${i===0?'🥇':i===1?'🥈':'🥉'} ${esc(c.label)}</td><td><strong>${c.count}</strong></td></tr>`).join("")}
        </tbody>
      </table>
    </div>` : `<div class="dash-table-wrap"><p class="hint">Aucune donnée d'analyse pour ${periodLabel}.</p></div>`;

  const kpiTotal = months.reduce((acc,m)=>{
    acc.received+=m.emails_received; acc.sent+=m.emails_sent; acc.calls+=m.calls; acc.bugs+=m.bugs; acc.demandes+=m.demandes;
    return acc;
  }, {received:0,sent:0,calls:0,bugs:0,demandes:0});

  const maxAvg = Math.max(...data.weekday_heatmap.map(h=>h.avg), 0.001);
  const heatmapHtml = `
    <div class="dash-table-wrap">
      <h3>🗓️ Charge moyenne par jour de la semaine (appels + emails reçus)</h3>
      <div class="heatmap-row">
        ${data.weekday_heatmap.map(h=>`
          <div class="heatmap-cell" style="background:${heatColor(h.avg, maxAvg)};">
            ${h.day.slice(0,3)}
            <span class="hm-val">${h.avg}</span>
          </div>
        `).join("")}
      </div>
    </div>`;

  document.getElementById("annualContent").innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">📧 Emails reçus (année)</div><div class="kpi-value">${kpiTotal.received}</div></div>
      <div class="kpi-card"><div class="kpi-label">📞 Appels (année)</div><div class="kpi-value">${kpiTotal.calls}</div></div>
      <div class="kpi-card danger"><div class="kpi-label">🔴 Bugs (année)</div><div class="kpi-value">${kpiTotal.bugs}</div></div>
      <div class="kpi-card success"><div class="kpi-label">🟢 Demandes (année)</div><div class="kpi-value">${kpiTotal.demandes}</div></div>
    </div>
    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Tendance mensuelle</h3>
      <canvas id="annualChart" height="90"></canvas>
    </div>
    ${heatmapHtml}
    ${topHtml}
    <div class="dash-table-wrap">
      <h3>Détail par mois</h3>
      <table class="dash-table">
        <thead><tr><th>Mois</th><th>Emails reçus</th><th>Emails envoyés</th><th>Appels</th><th>Bugs</th><th>Demandes</th></tr></thead>
        <tbody>
          ${months.map(m=>`<tr><td>${m.label}</td><td>${m.emails_received}</td><td>${m.emails_sent}</td><td>${m.calls}</td><td style="color:var(--danger);font-weight:700;">${m.bugs}</td><td style="color:var(--success);font-weight:700;">${m.demandes}</td></tr>`).join("") || `<tr><td colspan="6" style="color:var(--muted);">Aucune donnée</td></tr>`}
        </tbody>
      </table>
    </div>
  `;

  if(annualChart) annualChart.destroy();
  const ctx = document.getElementById("annualChart");
  annualChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: months.map(m=>m.label),
      datasets: [
        { label:"Emails reçus", data: months.map(m=>m.emails_received), borderColor:"#0B4965", backgroundColor:"#0B496522", tension:.3 },
        { label:"Appels", data: months.map(m=>m.calls), borderColor:"#F59F0A", backgroundColor:"#F59F0A22", tension:.3 },
        { label:"Bugs", data: months.map(m=>m.bugs), borderColor:"#DC2828", backgroundColor:"#DC282822", tension:.3 },
        { label:"Demandes", data: months.map(m=>m.demandes), borderColor:"#25935F", backgroundColor:"#25935F22", tension:.3 },
      ]
    },
    options: { responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}}, scales:{ y:{beginAtZero:true} } }
  });
  addPngExportButtons({annualChart});
}

// ============================================================
// PROMPTS COPILOT
// ============================================================
function renderPrompts(){
  const label = monthLabel(state.ym);
  document.getElementById("promptMonthLabel").textContent = label;

  document.getElementById("promptA").textContent =
`Dans la feuille "Mail", pour le mois de ${label} uniquement :
1. Utilise "DateTimeReceived" pour les emails reçus (colonne "Folder Path" = \\Inbox\\)
   et "DateTimeSent" pour les emails envoyés (autres dossiers / Sent Items).
2. Génère un tableau avec une ligne par jour du mois (y compris week-ends s'il y a des emails), avec :
   - Date
   - Nombre d'emails reçus ce jour
   - Nombre d'emails envoyés ce jour
3. Trie par date croissante.
Affiche uniquement ce tableau.`;

  document.getElementById("promptB").textContent =
`Analyse les emails du mois de ${label} dans la feuille "Mail" (colonnes Subject et Preview)
et génère un rapport de support avec :

1. STATISTIQUES GÉNÉRALES : total emails reçus/envoyés, nombre de problèmes techniques vs
   nombre de demandes d'information, ratio bugs/demandes, semaine la plus chargée.

2. CLASSIFICATION DES PROBLÈMES TECHNIQUES : identifie toi-même les catégories pertinentes
   ce mois-ci (elles peuvent différer du mois précédent, ne les fige pas à l'avance) parmi les
   emails signalant un dysfonctionnement. Pour chaque catégorie : Nombre, %, Description courte.

3. CLASSIFICATION DES DEMANDES D'INFORMATION : idem pour les emails de demande.
   Nombre, %, Description courte.

4. ÉVOLUTION HEBDOMADAIRE : découpe le mois en semaines calendaires, indique pour chacune :
   nombre de bugs, nombre de demandes, et une observation d'une phrase.

5. CONSTATS CLÉS : 4 à 6 phrases factuelles chiffrées.

6. RECOMMANDATIONS : 2 à 4 actions concrètes priorisées.

Présente le résultat sous forme de tableaux distincts, faciles à copier-coller.`;

  document.getElementById("promptC").textContent =
`Compare les emails de ${label} avec ceux de ${monthLabel(prevYm(state.ym))} dans la feuille "Mail" :
1. Variation en % du volume total d'emails reçus/envoyés entre les deux mois.
2. Catégories de problèmes techniques en hausse ou en baisse significative (>15%).
3. Nouvelles catégories apparues ce mois-ci, catégories disparues.
4. Une phrase de synthèse sur la tendance générale (amélioration ou dégradation du support).
Présente sous forme de tableau comparatif à 2 colonnes (mois précédent / mois courant).`;

  document.getElementById("promptF").textContent =
`Pour les emails du mois de ${label} dans la feuille "Mail" (colonnes Subject et Preview) :
1. Extrais les mots ou expressions clés les plus fréquents (hors mots vides comme "bonjour",
   "merci", "cordialement", articles, prépositions).
2. Regroupe les synonymes/variantes proches sous un seul terme représentatif (ex: "paiement",
   "payer", "paiements" → "paiement").
3. Donne la liste des 15 à 20 mots-clés les plus fréquents avec leur nombre d'occurrences,
   triée par fréquence décroissante.
Présente sous forme de tableau "Mot-clé | Fréquence", prêt à recopier dans le Nuage de
mots-clés de l'application (un mot-clé par ligne).`;
}

document.querySelectorAll(".copy-btn").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    const text = document.getElementById(btn.dataset.copy).textContent;
    navigator.clipboard.writeText(text).then(()=>toast("Prompt copié dans le presse-papier"));
  });
});

// ============================================================
// ADMINISTRATION
// ============================================================
const AUDIT_ACTION_LABELS = {
  login: "Connexion", login_failed: "Connexion échouée", logout: "Déconnexion",
  user_create: "Création utilisateur", user_update: "Modification utilisateur", user_delete: "Suppression utilisateur",
  import: "Import de données",
};

function auditLogRelativeDate(iso){
  const d = new Date(iso);
  return d.toLocaleString("fr-FR", {day:"2-digit", month:"2-digit", year:"numeric", hour:"2-digit", minute:"2-digit"});
}

async function loadAuditLog(){
  const tbody = document.querySelector("#auditLogTable tbody");
  if(!tbody) return;
  try{
    const res = await fetch("/api/admin/audit-log?limit=200");
    if(!res.ok) throw new Error("http " + res.status);
    const entries = await res.json();
    tbody.innerHTML = entries.map(e=>`
      <tr>
        <td style="white-space:nowrap;">${auditLogRelativeDate(e.ts)}</td>
        <td>${esc(e.user || "—")}</td>
        <td>${esc(AUDIT_ACTION_LABELS[e.action] || e.action)}</td>
        <td>${esc(e.module || "—")}</td>
        <td style="font-size:11.5px;color:var(--muted);">${esc(JSON.stringify(e.details || {}))}</td>
      </tr>
    `).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucun événement enregistré</td></tr>`;
  } catch(e){
    tbody.innerHTML = `<tr><td colspan="5" style="color:var(--danger);">Impossible de charger le journal d'audit.</td></tr>`;
  }
}

document.getElementById("refreshAuditLogBtn")?.addEventListener("click", loadAuditLog);

async function loadRawImports(){
  const tbody = document.querySelector("#rawImportsTable tbody");
  if(!tbody) return;
  tbody.innerHTML = `<tr><td colspan="6">Chargement...</td></tr>`;
  try{
    const res = await fetch("/api/tarkhiss/raw-imports");
    const rows = await res.json();
    if(!rows.length){ tbody.innerHTML = `<tr><td colspan="6">Aucun import archivé pour l'instant.</td></tr>`; return; }
    tbody.innerHTML = rows.map(r=>{
      const kindLabel = r.kind === "emails" ? "📧 Emails (Outlook)" : "📞 Appels";
      const vol = r.kind === "emails" ? `${r.total_received||0} reçus / ${r.total_sent||0} envoyés` : `${r.rows||0} lignes`;
      return `<tr>
        <td>${esc((r.uploaded_at||"").replace("T"," ").slice(0,16))}</td>
        <td>${kindLabel}</td>
        <td>${esc(r.original_name||"")}</td>
        <td>${(r.months||[]).map(monthLabel).join(", ")}</td>
        <td>${esc(vol)}</td>
        <td><a href="/api/tarkhiss/raw-imports/download/${encodeURIComponent(r.id)}" class="btn btn-outline btn-sm">⬇</a></td>
      </tr>`;
    }).join("");
  } catch(e){
    tbody.innerHTML = `<tr><td colspan="6">Erreur de chargement.</td></tr>`;
  }
}
document.getElementById("refreshRawImportsBtn")?.addEventListener("click", loadRawImports);

async function loadGlobalSettings(){
  const res = await fetch("/api/global-settings");
  globalSettings = await res.json();
  applyUiPalette(globalSettings.ui_palette || "ammps");
  applyAppBranding();
}

function applyAppBranding(){
  const appName = globalSettings.app_name || "SANAD";
  const appSubtitle = globalSettings.app_subtitle || "Plateforme de pilotage du Helpdesk SI";
  document.title = `${appName} — ${appSubtitle}`;
  const authTitle = document.getElementById("authBrandTitle");
  const authSubtitle = document.getElementById("authBrandSubtitle");
  if(authTitle) authTitle.textContent = appName;
  if(authSubtitle) authSubtitle.textContent = appSubtitle;
  const logoText = document.getElementById("sanadLogoText");
  if(logoText) logoText.textContent = appName;
}

function applyUiPalette(palette){
  const valid = ["ammps", "ocean", "emerald", "slate", "violet", "crimson"];
  const p = valid.includes(palette) ? palette : "ammps";
  if(p === "ammps") delete document.documentElement.dataset.palette;
  else document.documentElement.dataset.palette = p;
  document.querySelectorAll('input[name="uiPalette"]').forEach(input=>{
    input.checked = input.value === p;
  });
}

// Aperçu live de la palette : chaque clic l'applique immédiatement à l'écran entier, sans
// attendre "Enregistrer". Si l'admin quitte l'onglet sans enregistrer, on revient à la palette
// réellement sauvegardée (voir activateTab / revertThemePreviewIfUnsaved).
document.querySelectorAll('input[name="uiPalette"]').forEach(r=>{
  r.addEventListener("change", ()=> applyUiPalette(r.value));
});

function revertThemePreviewIfUnsaved(){
  const savedPalette = globalSettings.ui_palette || "ammps";
  if((document.documentElement.dataset.palette || "ammps") !== savedPalette){
    applyUiPalette(savedPalette);
  }
}

async function loadModuleSettings(){
  const res = await fetch(apiM("/settings"));
  moduleSettings = await res.json();
}

function renderContactsAdmin(){
  const tbody = document.querySelector("#contactsTable tbody");
  const contacts = moduleSettings.contacts || [];
  tbody.innerHTML = contacts.map((c,i)=>`
    <tr>
      <td>${esc(c.email)}</td>
      <td>${c.name||''}</td>
      <td><span class="role-pill ${c.role}">${c.role==='to'?'À':'Cc'}</span></td>
      <td><button class="admin-del-btn" data-i="${i}">Supprimer</button></td>
    </tr>
  `).join("") || `<tr><td colspan="4" style="color:var(--muted);">Aucun contact</td></tr>`;
  tbody.querySelectorAll(".admin-del-btn").forEach(b=>{
    b.addEventListener("click", async ()=>{
      moduleSettings.contacts.splice(Number(b.dataset.i), 1);
      await fetch(apiM("/settings"), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({contacts: moduleSettings.contacts})});
      renderContactsAdmin();
    });
  });
}

document.getElementById("addContactBtn").addEventListener("click", async ()=>{
  const email = document.getElementById("contactEmailInput").value.trim();
  const name = document.getElementById("contactNameInput").value.trim();
  const role = document.getElementById("contactRoleInput").value;
  if(!email) return;
  moduleSettings.contacts = moduleSettings.contacts || [];
  moduleSettings.contacts.push({email, name, role});
  await fetch(apiM("/settings"), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({contacts: moduleSettings.contacts})});
  document.getElementById("contactEmailInput").value = "";
  document.getElementById("contactNameInput").value = "";
  renderContactsAdmin();
  toast("Contact ajouté");
});

document.getElementById("saveGlobalSettingsBtn").addEventListener("click", async ()=>{
  globalSettings.agency_name = document.getElementById("agencyNameInput").value.trim();
  globalSettings.ui_palette = document.querySelector('input[name="uiPalette"]:checked')?.value || "ammps";
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({agency_name: globalSettings.agency_name, ui_palette: globalSettings.ui_palette})});
  applyUiPalette(globalSettings.ui_palette);
  toast("Paramètres généraux enregistrés");
});

document.getElementById("saveAdvancedEditingBtn").addEventListener("click", async ()=>{
  const payload = {
    app_name: document.getElementById("appNameInput").value.trim(),
    app_subtitle: document.getElementById("appSubtitleInput").value.trim(),
    report_title_tarkhiss: document.getElementById("reportTitleTarkhissInput").value.trim(),
    report_title_moussanada: document.getElementById("reportTitleMoussanadaInput").value.trim(),
    report_title_pchc: document.getElementById("reportTitlePchcInput").value.trim(),
    report_show_contact_names: document.getElementById("reportShowContactNamesInput").checked,
  };
  Object.assign(globalSettings, payload);
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  applyAppBranding();
  toast("Identité et grands titres enregistrés");
});

function randomToken(len=32){
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let out = "";
  for(let i=0;i<len;i++) out += chars[Math.floor(Math.random()*chars.length)];
  return out;
}

document.getElementById("regenerateTokenBtn").addEventListener("click", ()=>{
  document.getElementById("reminderTokenInput").value = randomToken();
});

document.getElementById("saveSmtpBtn").addEventListener("click", async ()=>{
  const payload = {
    smtp_host: document.getElementById("smtpHostInput").value.trim(),
    smtp_port: Number(document.getElementById("smtpPortInput").value || 587),
    smtp_user: document.getElementById("smtpUserInput").value.trim(),
    smtp_password: document.getElementById("smtpPasswordInput").value,
    smtp_from: document.getElementById("smtpFromInput").value.trim(),
    reminder_api_token: document.getElementById("reminderTokenInput").value.trim(),
  };
  Object.assign(globalSettings, payload);
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  toast("Configuration SMTP enregistrée");
});

document.getElementById("saveSecurityBtn").addEventListener("click", async ()=>{
  const payload = {
    session_timeout_minutes: Number(document.getElementById("sessionTimeoutInput").value || 240),
    password_min_length: Number(document.getElementById("pwMinLenInput").value || 8),
    password_require_digit: document.getElementById("pwRequireDigitInput").checked,
    password_require_upper: document.getElementById("pwRequireUpperInput").checked,
  };
  Object.assign(globalSettings, payload);
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  toast("Paramètres de sécurité enregistrés");
});

async function runRgpdSearch(){
  const q = document.getElementById("rgpdQueryInput").value.trim();
  const box = document.getElementById("rgpdResults");
  if(!q){ box.innerHTML = ""; return; }
  const res = await fetch(`/api/admin/rgpd/search?q=${encodeURIComponent(q)}`);
  const data = await res.json();
  if(!data.tickets_count && !data.pchc_count){
    box.innerHTML = `<p class="hint">Aucun résultat pour "${esc(q)}".</p>`;
    return;
  }
  box.innerHTML = `
    <p class="hint">${data.tickets_count} ticket(s) Moussanada, ${data.pchc_count} dossier(s) PCHC trouvés.</p>
    <button class="btn btn-outline" id="rgpdAnonymizeBtn" style="margin-bottom:10px;">🗑️ Anonymiser ces résultats</button>
    ${data.tickets.length ? `<table class="dash-table"><thead><tr><th>Ticket</th><th>Demandeur</th><th>Statut</th><th>Date</th></tr></thead><tbody>
      ${data.tickets.map(t=>`<tr><td>${esc(t.id)}</td><td>${esc(t.demandeur)}</td><td>${esc(t.statut)}</td><td>${esc((t.date_ouverture||"").slice(0,10))}</td></tr>`).join("")}
    </tbody></table>` : ""}
    ${data.pchc_records.length ? `<table class="dash-table" style="margin-top:10px;"><thead><tr><th>Référence</th><th>Entité</th><th>Statut</th></tr></thead><tbody>
      ${data.pchc_records.map(r=>`<tr><td>${esc(r.ref)}</td><td>${esc(r.entity)}</td><td>${esc(r.statut)}</td></tr>`).join("")}
    </tbody></table>` : ""}
  `;
  document.getElementById("rgpdAnonymizeBtn")?.addEventListener("click", async ()=>{
    if(!confirm(`Anonymiser définitivement toutes les occurrences de "${q}" ? Cette action est tracée dans le journal d'audit et n'est pas réversible depuis l'écran (seule une restauration de sauvegarde le permettrait).`)) return;
    const r = await fetch("/api/admin/rgpd/anonymize", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({q})});
    const d = await r.json();
    toast(`${d.tickets_touched} ticket(s) et ${d.pchc_touched} dossier(s) anonymisés`);
    runRgpdSearch();
  });
}
document.getElementById("rgpdSearchBtn").addEventListener("click", runRgpdSearch);
document.getElementById("rgpdExportBtn").addEventListener("click", ()=>{
  const q = document.getElementById("rgpdQueryInput").value.trim();
  if(!q) return toast("Saisir un nom à rechercher d'abord");
  window.location.href = `/api/admin/rgpd/export?q=${encodeURIComponent(q)}`;
});

document.getElementById("saveGlpiBtn").addEventListener("click", async ()=>{
  const payload = {
    glpi_url: document.getElementById("glpiUrlInput").value.trim(),
    glpi_app_token: document.getElementById("glpiAppTokenInput").value,
    glpi_user_token: document.getElementById("glpiUserTokenInput").value,
    glpi_entity_id: document.getElementById("glpiEntityIdInput").value.trim(),
  };
  Object.assign(globalSettings, payload);
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  toast("Configuration GLPI enregistrée");
});

document.getElementById("glpiTestConnectionBtn").addEventListener("click", async ()=>{
  const box = document.getElementById("glpiTestResult");
  box.textContent = "Test en cours...";
  box.style.color = "var(--muted)";
  const payload = {
    glpi_url: document.getElementById("glpiUrlInput").value.trim(),
    glpi_app_token: document.getElementById("glpiAppTokenInput").value,
    glpi_user_token: document.getElementById("glpiUserTokenInput").value,
    glpi_entity_id: document.getElementById("glpiEntityIdInput").value.trim(),
  };
  try{
    const res = await fetch("/api/admin/glpi/test-connection", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
    const data = await res.json();
    box.textContent = (data.ok ? "✅ " : "❌ ") + data.message;
    box.style.color = data.ok ? "var(--success)" : "var(--danger)";
  } catch(e){
    box.textContent = "❌ Erreur réseau lors du test.";
    box.style.color = "var(--danger)";
  }
});

async function loadGlpiProfilesSelect(){
  const select = document.getElementById("glpiProfileSelect");
  if(!select) return;
  try{
    const res = await fetch("/api/admin/glpi/profiles");
    const profiles = await res.json();
    select.innerHTML = profiles.map(p=>`<option value="${p.key}">${esc(p.label)}</option>`).join("");
  } catch(e){ /* silencieux : GLPI peut ne pas être configuré */ }
}

document.getElementById("glpiImportDirectBtn").addEventListener("click", async ()=>{
  const box = document.getElementById("glpiImportDirectResult");
  const profile = document.getElementById("glpiProfileSelect").value;
  const ym = document.getElementById("glpiImportYmInput").value;
  box.textContent = "Import en cours...";
  box.style.color = "var(--muted)";
  showBusy("Import direct depuis GLPI en cours...");
  try{
    const res = await fetch("/api/moussanada/glpi-import-direct", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({profile, ym})});
    const data = await res.json();
    if(data.ok){
      box.textContent = `✅ ${data.raw_tickets} ticket(s) GLPI -> ${data.records_stored} enregistrement(s) stocké(s).`;
      box.style.color = "var(--success)";
      toast("Import direct GLPI réussi");
    } else {
      box.textContent = "❌ " + data.error;
      box.style.color = "var(--danger)";
    }
  } catch(e){
    box.textContent = "❌ Erreur réseau lors de l'import.";
    box.style.color = "var(--danger)";
  } finally {
    hideBusy();
  }
});

async function loadTechnicienFilter(){
  try{
    const res = await fetch("/api/moussanada/technicien-filter");
    const data = await res.json();
    document.getElementById(data.mode === "whitelist" ? "techFilterWhitelistInput" : "techFilterExcludeInput").checked = true;
    document.getElementById("technicienFilterListInput").value = (data.technicians || []).join("\n");
  } catch(e){ /* silencieux */ }
}

document.getElementById("saveTechnicienFilterBtn").addEventListener("click", async ()=>{
  const mode = document.querySelector('input[name="technicienFilterMode"]:checked')?.value || "exclude";
  const technicians = document.getElementById("technicienFilterListInput").value.split("\n").map(s=>s.trim()).filter(Boolean);
  await fetch("/api/moussanada/technicien-filter", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({mode, technicians})});
  toast("Filtre techniciens enregistré — appliqué immédiatement aux rapports");
});

document.getElementById("loadTechnicienNamesBtn").addEventListener("click", async ()=>{
  const box = document.getElementById("technicienNamesKnown");
  box.textContent = "Chargement...";
  const res = await fetch("/api/moussanada/technicien-names");
  const names = await res.json();
  box.textContent = names.length ? `Noms connus : ${names.join(", ")}` : "Aucun technicien connu pour l'instant (importer des données Moussanada d'abord).";
});

document.getElementById("saveModuleSettingsBtn").addEventListener("click", async ()=>{
  moduleSettings.greeting = document.getElementById("defaultGreetingInput").value.trim();
  moduleSettings.signature_name = document.getElementById("sigNameInput").value.trim();
  moduleSettings.signature_function = document.getElementById("sigFunctionInput").value.trim();
  moduleSettings.signature_phone = document.getElementById("sigPhoneInput").value.trim();
  await fetch(apiM("/settings"), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(moduleSettings)});
  toast("Paramètres du volet enregistrés");
});

document.getElementById("logoInput").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("logo", file);
  fd.append("logo_type", "ammps");
  const res = await fetch("/api/upload-logo", {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){
    globalSettings.logo_filename = data.logo_filename;
    updateLogoPreview();
    toast("Logo AMMPS mis à jour");
  } else {
    toast(data.error || "Erreur import logo");
  }
});

document.getElementById("sanadLogoInput").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("logo", file);
  fd.append("logo_type", "sanad");
  const res = await fetch("/api/upload-logo", {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){
    globalSettings.sanad_logo_filename = data.logo_filename;
    updateLogoPreview();
    applySanadBrand();
    toast("Logo SANAD mis à jour");
  } else {
    toast(data.error || "Erreur import logo");
  }
});

function updateLogoPreview(){
  const img = document.getElementById("logoPreview");
  if(globalSettings.logo_filename){
    img.src = `/static/uploads/${globalSettings.logo_filename}?t=${Date.now()}`;
    img.classList.add("show");
  } else {
    img.classList.remove("show");
  }
  const simg = document.getElementById("sanadLogoPreview");
  if(globalSettings.sanad_logo_filename){
    simg.src = `/static/uploads/${globalSettings.sanad_logo_filename}?t=${Date.now()}`;
    simg.classList.add("show");
  } else {
    simg.classList.remove("show");
  }
}

function applySanadBrand(){
  const img = document.getElementById("sanadLogoImg");
  const txt = document.getElementById("sanadLogoText");
  if(globalSettings.sanad_logo_filename){
    img.src = `/static/uploads/${globalSettings.sanad_logo_filename}?t=${Date.now()}`;
    img.style.display = "block";
    txt.style.display = "none";
  } else {
    img.style.display = "none";
    txt.style.display = "inline";
  }
}

document.getElementById("exportAllBtn").addEventListener("click", ()=>{
  window.location.href = apiM("/export-all");
});

document.getElementById("importAllFile").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("backup", file);
  const res = await fetch(apiM("/import-all"), {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){
    toast("Sauvegarde restaurée");
    loadMonth(state.ym);
    loadAdmin();
  } else {
    toast(data.error || "Erreur import");
  }
  e.target.value = "";
});

document.getElementById("csvCallsBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-csv/calls/${state.ym}`);
});
document.getElementById("csvEmailsBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-csv/emails/${state.ym}`);
});
document.getElementById("csvProblemsBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-csv/problems/${state.ym}`);
});
document.getElementById("csvDemandesBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-csv/demandes/${state.ym}`);
});
document.getElementById("csvWeeklyBtn").addEventListener("click", ()=>{
  window.location.href = apiM(`/export-csv/weekly/${state.ym}`);
});

async function loadAdmin(){
  await loadGlobalSettings();
  await loadModuleSettings();
  document.getElementById("agencyNameInput").value = globalSettings.agency_name || "";
  document.getElementById("appNameInput").value = globalSettings.app_name || "";
  document.getElementById("appSubtitleInput").value = globalSettings.app_subtitle || "";
  document.getElementById("reportTitleTarkhissInput").value = globalSettings.report_title_tarkhiss || "";
  document.getElementById("reportTitleMoussanadaInput").value = globalSettings.report_title_moussanada || "";
  document.getElementById("reportTitlePchcInput").value = globalSettings.report_title_pchc || "";
  document.getElementById("reportShowContactNamesInput").checked = !!globalSettings.report_show_contact_names;
  document.getElementById("smtpHostInput").value = globalSettings.smtp_host || "";
  document.getElementById("smtpPortInput").value = globalSettings.smtp_port || 587;
  document.getElementById("smtpUserInput").value = globalSettings.smtp_user || "";
  document.getElementById("smtpFromInput").value = globalSettings.smtp_from || "";
  document.getElementById("reminderTokenInput").value = globalSettings.reminder_api_token || randomToken();
  document.getElementById("sessionTimeoutInput").value = globalSettings.session_timeout_minutes ?? 240;
  document.getElementById("pwMinLenInput").value = globalSettings.password_min_length ?? 8;
  document.getElementById("pwRequireDigitInput").checked = globalSettings.password_require_digit !== false;
  document.getElementById("pwRequireUpperInput").checked = globalSettings.password_require_upper !== false;
  document.getElementById("glpiUrlInput").value = globalSettings.glpi_url || "";
  document.getElementById("glpiAppTokenInput").value = globalSettings.glpi_app_token || "";
  document.getElementById("glpiUserTokenInput").value = globalSettings.glpi_user_token || "";
  document.getElementById("glpiEntityIdInput").value = globalSettings.glpi_entity_id || "";
  loadGlpiProfilesSelect();
  loadTechnicienFilter();
  updateLogoPreview();
  document.getElementById("defaultGreetingInput").value = moduleSettings.greeting || "";
  document.getElementById("sigNameInput").value = moduleSettings.signature_name || "";
  document.getElementById("sigFunctionInput").value = moduleSettings.signature_function || "";
  document.getElementById("sigPhoneInput").value = moduleSettings.signature_phone || "";
  const th = moduleSettings.alert_thresholds || {};
  document.getElementById("thResolutionMin").value = th.resolution_rate_min ?? "";
  document.getElementById("thBacklogMax").value = th.backlog_max ?? "";
  document.getElementById("thDelayMax").value = th.delay_max_h ?? "";
  document.getElementById("thVolumeVar").value = th.volume_variation_max ?? "";
  document.getElementById("reminderDayInput").value = moduleSettings.reminder_day ?? 5;

  document.getElementById("pchcThresholdsCard").style.display = currentModule === "pchc" ? "block" : "none";
  const pth = moduleSettings.pchc_thresholds || {};
  document.getElementById("pThHaut").value = pth.taux_haut ?? "";
  document.getElementById("pThBas").value = pth.taux_bas ?? "";
  document.getElementById("pThBacklog").value = pth.backlog_alert ?? "";
  document.getElementById("pThOld").value = pth.old_dossiers_alert ?? "";

  renderContactsAdmin();

  if(currentUser && currentUser.role === "admin"){
    await loadAuditLog();
  }

  const monthsRes = await fetch(apiM("/months-list"));
  const monthsData = await monthsRes.json();
  const tbody = document.querySelector("#monthsTable tbody");
  tbody.innerHTML = monthsData.map(m=>`
    <tr>
      <td>${m.label}</td>
      <td><span class="status-dot ${m.has_calls?'on':'off'}"></span></td>
      <td><span class="status-dot ${m.has_emails?'on':'off'}"></span></td>
      <td><span class="status-dot ${m.has_analysis?'on':'off'}"></span></td>
      <td><button class="admin-del-btn" data-ym="${m.ym}">Supprimer</button></td>
    </tr>
  `).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucune donnée enregistrée</td></tr>`;
  tbody.querySelectorAll(".admin-del-btn").forEach(b=>{
    b.addEventListener("click", async ()=>{
      if(!confirm(`Supprimer toutes les données de ${b.dataset.ym} ? Cette action est irréversible.`)) return;
      await fetch(apiM(`/delete-month/${b.dataset.ym}`), {method:"POST"});
      toast("Mois supprimé");
      loadAdmin();
    });
  });

  const logRes = await fetch(apiM("/send-log"));
  const logData = await logRes.json();
  const logBody = document.querySelector("#sendLogTable tbody");
  logBody.innerHTML = logData.map(l=>`
    <tr>
      <td>${new Date(l.date).toLocaleString('fr-FR')}</td>
      <td>${monthLabel(l.ym)}</td>
      <td style="font-size:12px;">${l.subject}</td>
      <td style="font-size:12px;">${(l.to||[]).join(", ") || "—"}</td>
      <td style="font-size:12px;">${(l.cc||[]).join(", ") || "—"}</td>
      <td>${l.file ? `<a href="/api/${currentModule}/send-log/download/${encodeURIComponent(l.file)}" class="btn-add" style="text-decoration:none;">⬇</a>` : ""}</td>
    </tr>
  `).join("") || `<tr><td colspan="6" style="color:var(--muted);">Aucun envoi enregistré</td></tr>`;

  try{
    const pchcLogRes = await fetch("/api/pchc/send-log");
    const pchcLogData = await pchcLogRes.json();
    const pchcLogBody = document.querySelector("#pchcSendLogTable tbody");
    if(pchcLogBody){
      pchcLogBody.innerHTML = pchcLogData.map(l=>`
        <tr>
          <td>${new Date(l.date).toLocaleString('fr-FR')}</td>
          <td style="font-size:12px;">${l.ym}</td>
          <td style="font-size:12px;">${l.subject}</td>
          <td style="font-size:12px;">${(l.to||[]).join(", ") || "—"}</td>
          <td style="font-size:12px;">${(l.cc||[]).join(", ") || "—"}</td>
          <td>${l.file ? `<a href="/api/pchc/send-log/download/${encodeURIComponent(l.file)}" class="btn-add" style="text-decoration:none;">⬇</a>` : ""}</td>
        </tr>
      `).join("") || `<tr><td colspan="6" style="color:var(--muted);">Aucun envoi enregistré</td></tr>`;
    }
  } catch(e){ /* silencieux */ }

  const alertHistRes = await fetch(apiM("/alert-history"));
  const alertHistData = await alertHistRes.json();
  const alertHistBody = document.querySelector("#alertHistoryTable tbody");
  if(alertHistBody){
    alertHistBody.innerHTML = alertHistData.map(h=>`
      <tr>
        <td>${new Date(h.date).toLocaleDateString('fr-FR')}</td>
        <td>${monthLabel(h.ym)}</td>
        <td style="font-size:12px;">${h.alerts.map(a=>esc(a)).join("<br>")}</td>
      </tr>
    `).join("") || `<tr><td colspan="3" style="color:var(--muted);">Aucun seuil déclenché à ce jour</td></tr>`;
  }
}

document.getElementById("saveThresholdsBtn").addEventListener("click", async ()=>{
  const payload = {
    alert_thresholds: {
      resolution_rate_min: document.getElementById("thResolutionMin").value === "" ? null : Number(document.getElementById("thResolutionMin").value),
      backlog_max: document.getElementById("thBacklogMax").value === "" ? null : Number(document.getElementById("thBacklogMax").value),
      delay_max_h: document.getElementById("thDelayMax").value === "" ? null : Number(document.getElementById("thDelayMax").value),
      volume_variation_max: document.getElementById("thVolumeVar").value === "" ? null : Number(document.getElementById("thVolumeVar").value),
    },
    reminder_day: Number(document.getElementById("reminderDayInput").value || 5),
  };
  moduleSettings = {...moduleSettings, ...payload};
  await fetch(apiM("/settings"), {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  toast("Seuils enregistrés");
});

document.getElementById("exportConfigBtn").addEventListener("click", ()=>{
  window.location.href = "/api/config-export";
});
document.getElementById("importConfigFile").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("config", file);
  const res = await fetch("/api/config-import", {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){ toast("Configuration importée"); loadAdmin(); }
  else toast(data.error || "Erreur");
  e.target.value = "";
});

// ============================================================
// MOUSSANADA
// ============================================================
function formatDH(hours){
  hours = Math.round((hours||0) * 100) / 100;
  return hours.toFixed(2).replace(".", ",") + " h";
}

async function loadMoussanadaMonth(ym){
  mState.ym = ym;
  monthInput.value = ym;
  const res = await fetch(`/api/moussanada/data/${ym}`);
  const data = await res.json();
  mState.month = data.month;
  mState.timeseries = data.timeseries;
  mState.notes = data.notes;
}

function prevYmM(ym){ return prevYm(ym); }

function mMetricsFor(ym){
  const row = mState.timeseries[ym] || {};
  return {
    ouverts: row.ouverts||0, resolus: row.resolus||0, en_retard: row.en_retard||0, clos: row.clos||0,
    resolution_h: row.resolution_h||0, cloture_h: row.cloture_h||0,
  };
}

// ---------- Agrégation Catégories/Services (racine ou nœud final) ----------
function aggKey(label){
  const parts = label.split(">").map(s=>s.trim());
  return mAggMode === "racine" ? parts[0] : parts[parts.length-1];
}

function aggregateItems(items){
  const map = {};
  items.forEach(it=>{
    const key = aggKey(it.label || "");
    if(!map[key]) map[key] = {label:key, ouverts:0, resolus:0, en_retard:0, fermes:0};
    map[key].ouverts += it.ouverts||0;
    map[key].resolus += it.resolus||0;
    map[key].en_retard += it.en_retard||0;
    map[key].fermes += it.fermes||0;
  });
  return Object.values(map).sort((a,b)=>b.ouverts-a.ouverts);
}

// ============================================================
// M-DASHBOARD
// ============================================================
function mCompareChip(label, current, previous){
  if(previous === 0 && current === 0) return "";
  let cls="flat", arrow="→", pctTxt="stable";
  if(previous === 0){ cls="up"; arrow="↑"; pctTxt="nouveau"; }
  else {
    const diff = ((current-previous)/previous)*100;
    if(Math.abs(diff) < 1){ cls="flat"; arrow="→"; pctTxt="stable"; }
    else if(diff > 0){ cls="up"; arrow="↑"; pctTxt=`+${diff.toFixed(0)}%`; }
    else { cls="down"; arrow="↓"; pctTxt=`${diff.toFixed(0)}%`; }
  }
  return `<div class="compare-chip">${label} <span class="${cls}">${arrow} ${pctTxt}</span> <span style="color:var(--muted);">vs ${monthLabel(prevYmM(mState.ym))}</span></div>`;
}

function renderMDashboard(){
  if(!mState.month) return;
  const cur = mMetricsFor(mState.ym);
  const prev = mMetricsFor(prevYmM(mState.ym));
  const yoy = mMetricsFor(prevYearYm(mState.ym));
  const showYoy = currentUser && (currentUser.role === "admin" || currentUser.role === "superviseur");
  const cat = aggregateItems(mState.month.categories);
  const srv = aggregateItems(mState.month.services);
  const techs = [...mState.month.techniciens].sort((a,b)=>b.ouverts-a.ouverts);

  const tauxResolution = cur.ouverts ? Math.round(cur.resolus/cur.ouverts*100) : 0;
  const tauxCloture = cur.ouverts ? Math.round(cur.clos/cur.ouverts*100) : 0;
  const prevTauxResolution = prev.ouverts ? Math.round(prev.resolus/prev.ouverts*100) : 0;
  const prevTauxCloture = prev.ouverts ? Math.round(prev.clos/prev.ouverts*100) : 0;

  const mAlerts = computeAlerts(
    {resolutionRate: tauxResolution, backlog: cur.en_retard, delayH: cur.resolution_h, volumeCur: cur.ouverts, volumePrev: prev.ouverts},
    moduleSettings.alert_thresholds
  );
  renderAlertBanner("moussanadaAlertBanner", mAlerts);
  logAlertHistory("moussanada", mState.ym, mAlerts);

  checkMoussanadaReminder();

  const teamAvgOuverts = techs.length ? techs.reduce((s,t)=>s+t.ouverts,0)/techs.length : 0;
  const teamAvgResolus = techs.length ? techs.reduce((s,t)=>s+t.resolus,0)/techs.length : 0;

  document.getElementById("mDashTitle").textContent = `Dashboard Moussanada — ${monthLabel(mState.ym)}`;

  const info = modulesInfo.moussanada || {label:"Moussanada"};
  const logoImg = globalSettings.logo_filename ? `<img src="/static/uploads/${globalSettings.logo_filename}" class="print-only" style="height:34px;margin-bottom:10px;">` : "";

  const html = `
    <div class="print-only" style="margin-bottom:10px;">
      ${logoImg}
      <div style="font-size:11px;color:#55616B;">${globalSettings.agency_name || ""} · ${info.label} · ${monthLabel(mState.ym)}</div>
    </div>

    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">🎫 Tickets ouverts</div><div class="kpi-value">${cur.ouverts}</div>${kpiDelta(cur.ouverts, prev.ouverts, "neutral")}</div>
      <div class="kpi-card success"><div class="kpi-label">✅ Tickets résolus</div><div class="kpi-value">${cur.resolus}</div>${kpiDelta(cur.resolus, prev.resolus, "positive")}</div>
      <div class="kpi-card"><div class="kpi-label">📦 Tickets clos</div><div class="kpi-value">${cur.clos}</div>${kpiDelta(cur.clos, prev.clos, "positive")}</div>
      <div class="kpi-card danger"><div class="kpi-label">⚠️ En retard</div><div class="kpi-value">${cur.en_retard}</div>${kpiDelta(cur.en_retard, prev.en_retard, "negative")}</div>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card success"><div class="kpi-label">Taux de résolution</div><div class="kpi-value">${tauxResolution}%</div>${kpiDelta(tauxResolution, prevTauxResolution, "positive")}</div>
      <div class="kpi-card success"><div class="kpi-label">Taux de clôture</div><div class="kpi-value">${tauxCloture}%</div>${kpiDelta(tauxCloture, prevTauxCloture, "positive")}</div>
      <div class="kpi-card"><div class="kpi-label">⏱️ Délai moy. résolution</div><div class="kpi-value" style="font-size:20px;">${formatDH(cur.resolution_h)}</div>${kpiDelta(cur.resolution_h, prev.resolution_h, "negative")}</div>
      <div class="kpi-card"><div class="kpi-label">⏱️ Délai moy. clôture</div><div class="kpi-value" style="font-size:20px;">${formatDH(cur.cloture_h)}</div>${kpiDelta(cur.cloture_h, prev.cloture_h, "negative")}</div>
    </div>

    ${showYoy ? `
    <div class="dash-table-wrap yoy-card">
      <h3>📆 Comparatif ${monthLabel(mState.ym)} vs ${monthLabel(prevYearYm(mState.ym))} (N-1)</h3>
      <div class="yoy-grid">
        ${metricCompareChip("Tickets ouverts", cur.ouverts, yoy.ouverts, "vs N-1")}
        ${metricCompareChip("Tickets résolus", cur.resolus, yoy.resolus, "vs N-1")}
        ${metricCompareChip("En retard", cur.en_retard, yoy.en_retard, "vs N-1")}
        ${metricCompareChip("Taux résolution", tauxResolution, (yoy.ouverts ? Math.round(yoy.resolus/yoy.ouverts*100) : 0), "vs N-1")}
      </div>
    </div>` : ""}


    <div class="charts-row">
      <div class="chart-card">
        <h3>Comparatif Volumes — ${monthLabel(prevYmM(mState.ym))} / ${monthLabel(mState.ym)}</h3>
        <canvas id="mVolumeChart" height="150"></canvas>
      </div>
      <div class="chart-card">
        <h3>Comparatif Délais moyens (heures)</h3>
        <canvas id="mDelayChart" height="150"></canvas>
      </div>
    </div>

    <div class="chart-card" id="mHeatmapCard" style="margin-bottom:16px;display:none;">
      <h3>Heatmap de charge (jour × heure)</h3>
      <p style="font-size:11.5px;color:var(--muted);margin:0 0 8px;">Basée sur les tickets bruts GLPI importés (Administration → Import tickets bruts).</p>
      <img id="mHeatmapImg" style="width:100%;max-width:760px;display:block;margin:0 auto;" alt="Heatmap de charge Moussanada">
    </div>

    <div class="charts-row" id="mTicketAnalyticsRow" style="display:none;">
      <div class="chart-card">
        <h3>Répartition par Type</h3>
        <canvas id="mTypeChart" height="150"></canvas>
      </div>
      <div class="chart-card">
        <h3>Distribution des délais de résolution</h3>
        <canvas id="mDelayDistChart" height="150"></canvas>
      </div>
    </div>

    <div class="dash-table-wrap" id="mSouffranceWrap" style="display:none;">
      <h3>⚠️ Tickets en souffrance (non résolus/clos)</h3>
      <table class="dash-table">
        <thead><tr><th>ID</th><th>Titre</th><th>Statut</th><th>Ouvert le</th><th>Technicien</th></tr></thead>
        <tbody id="mSouffranceBody"></tbody>
      </table>
    </div>

    <div class="charts-row">
      <div class="chart-card">
        <h3>Tendance ouverts/résolus/clos (12 mois)</h3>
        <canvas id="mTrendChart" height="140"></canvas>
      </div>
      <div class="chart-card">
        <h3>Évolution du stock en retard (backlog)</h3>
        <canvas id="mBacklogChart2" height="140"></canvas>
      </div>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
        <h3 style="margin:0;">Top 10 catégories</h3>
        <div class="agg-toggle" id="mAggToggleCat">
          <button data-mode="racine" class="${mAggMode==='racine'?'active':''}">Racine</button>
          <button data-mode="leaf" class="${mAggMode==='leaf'?'active':''}">Nœud final</button>
        </div>
      </div>
      <canvas id="mCatChart" height="130"></canvas>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Top 10 services demandeurs</h3>
      <canvas id="mSrvChart" height="130"></canvas>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Analyse Pareto — catégories (80% des tickets)</h3>
      <canvas id="mParetoChart" height="140"></canvas>
    </div>

    <div class="charts-row">
      <div class="chart-card">
        <h3>Charge par technicien (ouverts vs résolus)</h3>
        <canvas id="mTechChart" height="150"></canvas>
      </div>
      <div class="chart-card">
        <h3>Performance techniciens</h3>
        <canvas id="mRadarChart" height="150"></canvas>
      </div>
    </div>

    <div class="dash-table-wrap">
      <h3>👤 Comparaison technicien vs moyenne équipe</h3>
      <table class="dash-table">
        <thead><tr><th>Technicien</th><th>Ouverts</th><th>Résolus</th><th>vs Moy. ouverts</th><th>vs Moy. résolus</th></tr></thead>
        <tbody>
          ${techs.map(t=>{
            const dO = teamAvgOuverts ? Math.round(((t.ouverts-teamAvgOuverts)/teamAvgOuverts)*100) : 0;
            const dR = teamAvgResolus ? Math.round(((t.resolus-teamAvgResolus)/teamAvgResolus)*100) : 0;
            return `<tr>
              <td>${esc(t.label)}</td><td>${t.ouverts}</td><td>${t.resolus}</td>
              <td style="color:${dO>=0?'var(--success)':'var(--danger)'};font-weight:700;">${dO>=0?'+':''}${dO}%</td>
              <td style="color:${dR>=0?'var(--success)':'var(--danger)'};font-weight:700;">${dR>=0?'+':''}${dR}%</td>
            </tr>`;
          }).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucune donnée</td></tr>`}
          <tr style="background:var(--primary-pale);font-weight:700;">
            <td>Moyenne équipe</td><td>${teamAvgOuverts.toFixed(1)}</td><td>${teamAvgResolus.toFixed(1)}</td><td></td><td></td>
          </tr>
        </tbody>
      </table>
    </div>

    ${MOUSSANADA_SECTIONS.map(([key,title])=>{
      const text = (mState.notes.sections||{})[key];
      if(!text || !text.trim()) return "";
      return `<div class="dash-table-wrap synth-card">
        <h3>📝 ${title}</h3>
        ${text.split("\n").filter(p=>p.trim()).map(p=>`<p style="font-size:13px;color:var(--ink);line-height:1.6;">${p}</p>`).join("")}
      </div>`;
    }).join("")}

    <div class="synth-cols">
      <div class="dash-table-wrap synth-card constats">
        <h3>🔑 Constats clés</h3>
        <ul>${(mState.notes.constats||[]).map(c=>`<li>${c}</li>`).join("") || '<li style="color:var(--muted);">—</li>'}</ul>
      </div>
      <div class="dash-table-wrap synth-card reco">
        <h3>✅ Recommandations</h3>
        <ul>${(mState.notes.recommandations||[]).map(c=>`<li>${c}</li>`).join("") || '<li style="color:var(--muted);">—</li>'}</ul>
      </div>
    </div>
  `;
  document.getElementById("mDashboardContent").innerHTML = html;

  document.querySelectorAll("#mAggToggleCat button").forEach(b=>{
    b.addEventListener("click", ()=>{ mAggMode = b.dataset.mode; renderMDashboard(); });
  });

  Object.values(mCharts).forEach(c=>c && c.destroy());

  // Comparatif volumes
  mCharts.volume = new Chart(document.getElementById("mVolumeChart"), {
    type:"bar",
    data:{
      labels:["Ouverts","Résolus","En retard","Clos"],
      datasets:[
        {label: monthLabel(prevYmM(mState.ym)), data:[prev.ouverts,prev.resolus,prev.en_retard,prev.clos], backgroundColor:"#A9C7D6", borderRadius:5},
        {label: monthLabel(mState.ym), data:[cur.ouverts,cur.resolus,cur.en_retard,cur.clos], backgroundColor:"#0B4965", borderRadius:5},
      ]
    },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:true, anchor:'end', align:'top', font:{weight:'bold', size:11}, color:'#0D1926'}}, scales:{y:{beginAtZero:true, grace:'12%'}}}
  });

  mCharts.delay = new Chart(document.getElementById("mDelayChart"), {
    type:"bar",
    data:{
      labels:["Délai résolution","Délai clôture"],
      datasets:[
        {label: monthLabel(prevYmM(mState.ym)), data:[prev.resolution_h,prev.cloture_h], backgroundColor:"#BEE3D0", borderRadius:5},
        {label: monthLabel(mState.ym), data:[cur.resolution_h,cur.cloture_h], backgroundColor:"#25935F", borderRadius:5},
      ]
    },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:true, anchor:'end', align:'top', font:{weight:'bold', size:11}, color:'#0D1926', formatter:(v)=>v.toFixed(1)+'h'}}, scales:{y:{beginAtZero:true, grace:'12%', title:{display:true,text:"heures"}}}}
  });

  // Tendance 12 mois
  const months = Object.keys(mState.timeseries).sort();
  mCharts.trend = new Chart(document.getElementById("mTrendChart"), {
    type:"line",
    data:{
      labels: months.map(m=>monthLabel(m)),
      datasets:[
        {label:"Ouverts", data: months.map(m=>mState.timeseries[m].ouverts||0), borderColor:"#0B4965", backgroundColor:"#0B496522", tension:.3},
        {label:"Résolus", data: months.map(m=>mState.timeseries[m].resolus||0), borderColor:"#25935F", backgroundColor:"#25935F22", tension:.3},
        {label:"Clos", data: months.map(m=>mState.timeseries[m].clos||0), borderColor:"#F59F0A", backgroundColor:"#F59F0A22", tension:.3},
      ]
    },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}}, scales:{y:{beginAtZero:true}}}
  });

  mCharts.backlog = new Chart(document.getElementById("mBacklogChart2"), {
    type:"bar",
    data:{ labels: months.map(m=>monthLabel(m)), datasets:[{label:"En retard", data: months.map(m=>mState.timeseries[m].en_retard||0), backgroundColor:"#DC2828", borderRadius:5}] },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{display:false}, datalabels:{display:true, anchor:'end', align:'top', font:{weight:'bold',size:10}, color:'#0D1926'}}, scales:{y:{beginAtZero:true, grace:'15%'}}}
  });

  const topCat = cat.slice(0,10), topSrv = srv.slice(0,10);
  mCharts.cat = new Chart(document.getElementById("mCatChart"), {
    type:"bar",
    data:{ labels: topCat.map(c=>c.label), datasets:[
      {label:"Ouverts", data: topCat.map(c=>c.ouverts), backgroundColor:"#0B4965", borderRadius:4},
      {label:"Résolus", data: topCat.map(c=>c.resolus), backgroundColor:"#25935F", borderRadius:4},
    ]},
    options:{indexAxis:"y", responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:true, anchor:'end', align:'end', font:{weight:'bold',size:10}, color:'#0D1926'}}, scales:{x:{beginAtZero:true, grace:'15%'}}}
  });

  mCharts.srv = new Chart(document.getElementById("mSrvChart"), {
    type:"bar",
    data:{ labels: topSrv.map(c=>c.label), datasets:[
      {label:"Ouverts", data: topSrv.map(c=>c.ouverts), backgroundColor:"#1794CF", borderRadius:4},
      {label:"Résolus", data: topSrv.map(c=>c.resolus), backgroundColor:"#25935F", borderRadius:4},
    ]},
    options:{indexAxis:"y", responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:true, anchor:'end', align:'end', font:{weight:'bold',size:10}, color:'#0D1926'}}, scales:{x:{beginAtZero:true, grace:'15%'}}}
  });

  // Pareto
  const paretoSorted = [...cat].sort((a,b)=>b.ouverts-a.ouverts).slice(0,12);
  const totalCat = cat.reduce((s,c)=>s+c.ouverts,0) || 1;
  let cum = 0;
  const cumPct = paretoSorted.map(c=>{ cum += c.ouverts; return Math.round(cum/totalCat*100); });
  mCharts.pareto = new Chart(document.getElementById("mParetoChart"), {
    data:{
      labels: paretoSorted.map(c=>c.label),
      datasets:[
        {type:"bar", label:"Ouverts", data: paretoSorted.map(c=>c.ouverts), backgroundColor:"#0B4965", borderRadius:4, order:2},
        {type:"line", label:"% cumulé", data: cumPct, borderColor:"#F59F0A", backgroundColor:"#F59F0A", yAxisID:"y1", tension:.3, order:1},
      ]
    },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}},
      scales:{ y:{beginAtZero:true}, y1:{position:"right", min:0, max:100, grid:{drawOnChartArea:false}, ticks:{callback:v=>v+"%"}} }}
  });

  // Charge techniciens
  mCharts.tech = new Chart(document.getElementById("mTechChart"), {
    type:"bar",
    data:{ labels: techs.map(t=>t.label), datasets:[
      {label:"Ouverts", data: techs.map(t=>t.ouverts), backgroundColor:"#0B4965", borderRadius:4},
      {label:"Résolus", data: techs.map(t=>t.resolus), backgroundColor:"#25935F", borderRadius:4},
    ]},
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:true, anchor:'end', align:'top', font:{weight:'bold',size:10}, color:'#0D1926'}}, scales:{y:{beginAtZero:true, grace:'15%'}}}
  });

  // Radar techniciens (top 6)
  const radarTechs = techs.slice(0,6);
  const maxOuverts = Math.max(1, ...radarTechs.map(t=>t.ouverts));
  const maxDelai = Math.max(1, ...radarTechs.map(t=>t.delai_resolution_s||0));
  const palette = ["#0B4965","#DC2828","#F59F0A","#25935F","#8E5AA6","#1794CF"];
  mCharts.radar = new Chart(document.getElementById("mRadarChart"), {
    type:"radar",
    data:{
      labels:["Ouverts","Résolus","Taux résolution","Rapidité"],
      datasets: radarTechs.map((t,i)=>({
        label: t.label,
        data:[
          Math.round(t.ouverts/maxOuverts*100),
          t.ouverts ? Math.round(t.resolus/t.ouverts*100) : 0,
          t.ouverts ? Math.round(t.resolus/t.ouverts*100) : 0,
          t.delai_resolution_s ? Math.round(100 - (t.delai_resolution_s/maxDelai*100)) : 100,
        ],
        borderColor: palette[i%palette.length], backgroundColor: palette[i%palette.length]+"22",
      }))
    },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom", labels:{boxWidth:10,font:{size:10}}}}, scales:{r:{min:0,max:100}}}
  });

  setupEmailPanel("mToCheckList","mCcCheckList","mGreetingInput");
  addPngExportButtons(mCharts);
  refreshMoussanadaHeatmap();
}

async function refreshMoussanadaHeatmap(){
  const card = document.getElementById("mHeatmapCard");
  const img = document.getElementById("mHeatmapImg");
  if(!card || !img) return;
  try{
    const res = await fetch(`/api/moussanada/heatmap?start=${mState.ym}&end=${mState.ym}`);
    const data = await res.json();
    if(data.total_tickets > 0){
      img.src = `/api/moussanada/heatmap.png?start=${mState.ym}&end=${mState.ym}&_=${Date.now()}`;
      card.style.display = "";
    } else {
      card.style.display = "none";
    }
  } catch(e){
    card.style.display = "none";
  }
  await refreshMoussanadaTicketAnalytics();
}

let mTypeChart = null, mDelayDistChart = null;
async function refreshMoussanadaTicketAnalytics(){
  const row = document.getElementById("mTicketAnalyticsRow");
  const souffranceWrap = document.getElementById("mSouffranceWrap");
  if(!row) return;
  try{
    const res = await fetch(`/api/moussanada/tickets-analytics?start=${mState.ym}&end=${mState.ym}`);
    const data = await res.json();
    if(!data.total_tickets){
      row.style.display = "none";
      souffranceWrap.style.display = "none";
      return;
    }
    row.style.display = "";
    const typeEntries = Object.entries(data.type_counts);
    if(mTypeChart) mTypeChart.destroy();
    mTypeChart = new Chart(document.getElementById("mTypeChart"), {
      type:"doughnut",
      data:{ labels: typeEntries.map(([k])=>k), datasets:[{ data: typeEntries.map(([,v])=>v), backgroundColor:["#0B4965","#F59F0A","#25935F","#DC2828"] }] },
      options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}}}
    });

    const delayEntries = Object.entries(data.delay_buckets);
    if(mDelayDistChart) mDelayDistChart.destroy();
    mDelayDistChart = new Chart(document.getElementById("mDelayDistChart"), {
      type:"bar",
      data:{ labels: delayEntries.map(([k])=>k), datasets:[{ label:"Tickets", data: delayEntries.map(([,v])=>v), backgroundColor:"#1794CF", borderRadius:4 }] },
      options:{responsive:true, animation:{duration:700}, plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}}}
    });

    if(data.tickets_en_souffrance && data.tickets_en_souffrance.length){
      souffranceWrap.style.display = "";
      document.getElementById("mSouffranceBody").innerHTML = data.tickets_en_souffrance.map(t=>`
        <tr>
          <td>${esc(t.id)}</td>
          <td>${esc(t.titre)}</td>
          <td>${esc(t.statut)}</td>
          <td>${esc((t.date_ouverture||"").slice(0,10))}</td>
          <td>${esc(t.technicien)}</td>
        </tr>
      `).join("");
    } else {
      souffranceWrap.style.display = "none";
    }
  } catch(e){
    row.style.display = "none";
    if(souffranceWrap) souffranceWrap.style.display = "none";
  }
}

document.getElementById("mPdfBtn").addEventListener("click", ()=>{
  downloadPdfFromServer(`/api/moussanada/export-pdf/${mState.ym}`, {charts: collectMoussanadaChartImages()}, `rapport_moussanada_${mState.ym}.pdf`);
});
document.getElementById("mOnePagerBtn").addEventListener("click", ()=>{
  const cur = mMetricsFor(mState.ym);
  const kpis = [
    {label:"TICKETS OUVERTS", value: cur.ouverts},
    {label:"TICKETS RÉSOLUS", value: cur.resolus, color:"#25935F"},
    {label:"EN RETARD", value: cur.en_retard, color:"#DC2828"},
    {label:"DÉLAI RÉSOLUTION", value: formatDH(cur.resolution_h)},
  ];
  const canvas = buildOnePagerCanvas(
    `Support Moussanada — ${monthLabel(mState.ym)}`,
    "Synthèse mensuelle — Helpdesk GLPI",
    kpis,
    [mCharts.volume?.canvas, mCharts.delay?.canvas]
  );
  downloadCanvas(canvas, `onepager_moussanada_${mState.ym}.png`);
});
document.getElementById("mXlsxBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-xlsx/${mState.ym}`;
});
document.getElementById("mMailToggleBtn").addEventListener("click", ()=>{
  const panel = document.getElementById("mEmailPanel");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
});
function collectMoussanadaChartImages(){
  const map = {volume:"volume", delay:"delay", trend:"trend", cat:"cat", srv:"srv", tech:"tech", backlog:"backlog", pareto:"pareto", radar:"radar"};
  const out = {};
  Object.entries(map).forEach(([key, chartKey])=>{
    const chart = mCharts[chartKey];
    if(chart && chart.toBase64Image){
      try{ out[key] = chart.toBase64Image('image/png', 1); }catch(e){}
    }
  });
  return out;
}

document.getElementById("mPreviewMailBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients("mToCheckList","mCcCheckList");
  const greeting = document.getElementById("mGreetingInput").value.trim();
  const charts = collectMoussanadaChartImages();
  const res = await fetch(`/api/moussanada/preview-email/${mState.ym}`, {
    method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({to, cc, greeting, charts})
  });
  const data = await res.json();
  openPreviewModal(data.subject, data.html);
});
document.getElementById("mConfirmSendBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients("mToCheckList","mCcCheckList");
  if(!to.length){ toast("Sélectionnez au moins un destinataire (À)"); return; }
  const greeting = document.getElementById("mGreetingInput").value.trim();
  const charts = collectMoussanadaChartImages();
  toast("Génération du rapport...");
  const res = await fetch(`/api/moussanada/send-email/${mState.ym}`, {
    method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({to, cc, greeting, charts})
  });
  const data = await res.json();
  toast(data.opened ? "Ouverture dans Outlook..." : "Fichier .eml généré : " + data.file);
  document.getElementById("mEmailPanel").style.display = "none";
});

// ============================================================
// M-SOURCES
// ============================================================
function renderMSources(){
  if(!mState.month) return;
  document.getElementById("mSourcesMonthLabel").textContent = monthLabel(mState.ym);
  const t = mState.month.tickets, d = mState.month.durations;
  document.querySelector("#mTicketsPreviewTable tbody").innerHTML = `<tr>
    <td>${t.ouverts}</td><td>${t.resolus}</td><td>${t.en_retard}</td><td>${t.clos}</td>
    <td>${formatDH(d.resolution_h)}</td><td>${formatDH(d.cloture_h)}</td>
  </tr>`;

  const fillTable = (sel, items)=>{
    document.querySelector(sel+" tbody").innerHTML = items.map(it=>`
      <tr><td>${esc(it.label)}</td><td>${it.ouverts}</td><td>${it.resolus}</td><td>${it.en_retard}</td><td>${it.fermes}</td></tr>
    `).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucune donnée — importez un export GLPI</td></tr>`;
  };
  fillTable("#mCategoriesTable", mState.month.categories);
  fillTable("#mServicesTable", mState.month.services);
  fillTable("#mTechTable", mState.month.techniciens);
  document.getElementById("mCatCount").textContent = `${mState.month.categories.length} lignes`;
  document.getElementById("mSrvCount").textContent = `${mState.month.services.length} lignes`;
}

document.getElementById("mImportXlsx").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("file", file);
  document.getElementById("mImportHint").textContent = "Import en cours...";
  document.getElementById("mImportHint").classList.add("show");
  showBusy("Import du fichier GLPI en cours...");
  let res, data;
  try{
    res = await fetch(`/api/moussanada/import/${mState.ym}`, {method:"POST", body: fd});
    data = await res.json();
  } finally {
    hideBusy();
  }
  if(data.ok){
    toast("Import réussi");
    document.getElementById("mImportHint").textContent = `✓ Importé : ${data.result.tickets_months} mois de tickets, ${data.result.categories} catégories, ${data.result.services} services, ${data.result.techniciens} techniciens`;
    const detected = data.result.detected_month;
    const banner = document.getElementById("mDetectedMonthBanner");
    if(detected && detected !== mState.ym){
      banner.style.display = "block";
      banner.innerHTML = `📅 Mois détecté dans le fichier : <strong>${monthLabel(detected)}</strong> (mois actuellement sélectionné : ${monthLabel(mState.ym)}).
        <button class="btn-add" id="useDetectedMonthBtn" style="margin-left:8px;">Utiliser ${monthLabel(detected)}</button>`;
      document.getElementById("useDetectedMonthBtn").addEventListener("click", async ()=>{
        await loadMoussanadaMonth(detected);
        renderMSources();
        banner.style.display = "none";
        toast(`Mois basculé sur ${monthLabel(detected)}`);
      });
    } else {
      banner.style.display = "none";
    }
    await loadMoussanadaMonth(mState.ym);
    renderMSources();
  } else {
    toast(data.error || "Erreur d'import");
    document.getElementById("mImportHint").textContent = "";
  }
  e.target.value = "";
});

document.getElementById("mCsvImportBtn").addEventListener("click", async ()=>{
  const file = document.getElementById("mCsvFile").files[0];
  const kind = document.getElementById("mUnifiedKind").value;
  if(!file){ toast("Choisissez un fichier CSV"); return; }
  const doImport = async (force)=>{
    const fd = new FormData();
    fd.append("file", file);
    fd.append("kind", kind);
    if(force) fd.append("force", "1");
    showBusy("Import du fichier CSV en cours...");
    let res, data;
    try{
      res = await fetch(`/api/moussanada/import/${mState.ym}`, {method:"POST", body: fd});
      data = await res.json();
    } finally {
      hideBusy();
    }
    const box = document.getElementById("mUnifiedResult");
    if(data.conflict && !force){
      if(confirm(data.message + "\n\nContinuer quand même ?")) await doImport(true);
      return;
    }
    if(data.ok){
      box.textContent = `✅ ${data.result.rows} ligne(s) importée(s).`;
      box.style.color = "var(--success)";
      toast("Import CSV réussi");
      await loadMoussanadaMonth(mState.ym);
      renderMSources();
    } else {
      box.textContent = "❌ " + (data.error || "Erreur d'import");
      box.style.color = "var(--danger)";
      toast(data.error || "Erreur d'import");
    }
  };
  await doImport(false);
});

// ---------- Mode unifié Direct/CSV (Lot E) ----------
// Profils GLPI disponibles en Mode Direct (les autres kinds — durées, services — restent
// CSV uniquement tant qu'aucun profil GLPI n'a été défini pour eux, voir glpi_profiles.py).
const MOUSSANADA_DIRECT_PROFILES = { tickets: null, categories: "ym", demandeurs: "ym", techniciens: "ym" };

function updateUnifiedImportUI(){
  const kind = document.getElementById("mUnifiedKind").value;
  const directAvailable = kind in MOUSSANADA_DIRECT_PROFILES;
  document.getElementById("mModeDirectInput").disabled = !directAvailable;
  if(!directAvailable && document.getElementById("mModeDirectInput").checked){
    document.getElementById("mModeCsvInput").checked = true;
  }
  const mode = document.querySelector('input[name="mImportMode"]:checked').value;
  document.getElementById("mModeCsvBlock").style.display = mode === "csv" ? "" : "none";
  document.getElementById("mModeDirectBlock").style.display = mode === "direct" ? "" : "none";
  document.getElementById("mUnifiedDirectHint").textContent = directAvailable
    ? (MOUSSANADA_DIRECT_PROFILES[kind] === "ym" ? "Le mois cible ci-dessus sera utilisé pour cet import." : "Récupère l'ensemble de la série mensuelle disponible côté GLPI.")
    : "Mode Direct pas encore disponible pour ce type de données — utiliser le Mode CSV.";
  document.getElementById("mUnifiedYmInput").value = mState.ym || "";
}
document.getElementById("mUnifiedKind").addEventListener("change", updateUnifiedImportUI);
document.querySelectorAll('input[name="mImportMode"]').forEach(r=> r.addEventListener("change", updateUnifiedImportUI));

document.getElementById("mUnifiedDirectBtn").addEventListener("click", async ()=>{
  const kind = document.getElementById("mUnifiedKind").value;
  const ym = document.getElementById("mUnifiedYmInput").value;
  const box = document.getElementById("mUnifiedResult");
  if(MOUSSANADA_DIRECT_PROFILES[kind] === "ym" && !ym){
    box.textContent = "❌ Choisir un mois cible pour ce type de données.";
    box.style.color = "var(--danger)";
    return;
  }
  const doImport = async (force)=>{
    box.textContent = "Import en cours...";
    box.style.color = "var(--muted)";
    showBusy("Import direct depuis GLPI en cours...");
    try{
      const res = await fetch("/api/moussanada/glpi-import-direct", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({profile: kind, ym, force})});
      const data = await res.json();
      if(data.conflict && !force){
        hideBusy();
        if(confirm(data.message + "\n\nContinuer quand même ?")) await doImport(true);
        return;
      }
      if(data.ok){
        box.textContent = `✅ ${data.raw_tickets} ticket(s) GLPI -> ${data.records_stored} enregistrement(s) stocké(s).`;
        box.style.color = "var(--success)";
        toast("Import direct GLPI réussi");
        await loadMoussanadaMonth(mState.ym);
        renderMSources();
      } else {
        box.textContent = "❌ " + data.error;
        box.style.color = "var(--danger)";
      }
    } catch(e){
      box.textContent = "❌ Erreur réseau lors de l'import.";
      box.style.color = "var(--danger)";
    } finally {
      hideBusy();
    }
  };
  await doImport(false);
});
updateUnifiedImportUI();

document.getElementById("mCsvCatBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/categories/${mState.ym}`;
});
document.getElementById("mCsvSrvBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/services/${mState.ym}`;
});
document.getElementById("mCsvTechBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/techniciens/${mState.ym}`;
});
document.getElementById("mCsvTicketsBtn").addEventListener("click", ()=>{
  window.location.href = "/api/moussanada/export-csv-tickets";
});

document.getElementById("mImportRawTickets").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const hint = document.getElementById("mImportRawTicketsHint");
  hint.textContent = "Import en cours...";
  showBusy("Import des tickets bruts GLPI en cours...");
  const fd = new FormData();
  fd.append("file", file);
  try{
    const res = await fetch("/api/moussanada/import-tickets", {method:"POST", body: fd});
    const data = await res.json();
    if(data.ok){
      const r = data.result;
      hint.textContent = r.rows > 0 ? `${r.rows} ticket(s) importé(s) — ${r.total} au total.` : "Aucun ticket détecté.";
      if(data.warnings && data.warnings.length) toast(data.warnings[0]);
      else toast("Import des tickets bruts réussi");
      notifyBrowser("Import terminé", `${r.rows} ticket(s) importé(s) sur Moussanada.`);
      refreshMoussanadaHeatmap();
    } else {
      hint.textContent = "";
      toast(data.error || "Erreur d'import");
    }
  } catch(err){
    hint.textContent = "";
    toast("Erreur d'import");
  } finally {
    hideBusy();
  }
  e.target.value = "";
});

// ============================================================
// M-ANALYSE & COPILOT
// ============================================================
function renderMPrompt(){
  const label = monthLabel(mState.ym);
  const prevLabel = monthLabel(prevYmM(mState.ym));
  const cur = mMetricsFor(mState.ym);
  const prev = mMetricsFor(prevYmM(mState.ym));
  const cat = aggregateItems(mState.month ? mState.month.categories : []).slice(0,10);
  const srv = aggregateItems(mState.month ? mState.month.services : []).slice(0,10);
  const techs = mState.month ? mState.month.techniciens : [];

  const catLines = cat.map(c=>`- ${c.label} : ${c.ouverts} ouverts, ${c.resolus} résolus`).join("\n");
  const srvLines = srv.map(c=>`- ${c.label} : ${c.ouverts} ouverts, ${c.resolus} résolus`).join("\n");
  const techLines = techs.map(t=>`- ${t.label} : ${t.ouverts} ouverts, ${t.resolus} résolus, ${t.en_retard} en retard`).join("\n");

  document.getElementById("mPromptMonthLabel").textContent = label;
  document.getElementById("mPrompt").textContent =
`Tu es un analyste Helpdesk. Rédige une synthèse mensuelle professionnelle et un texte d'email
Direction pour le support Moussanada (plateforme GLPI), à partir des données suivantes :

MOIS ANALYSÉ : ${label}
MOIS PRÉCÉDENT : ${prevLabel}

VOLUMES
- Tickets ouverts : ${cur.ouverts} (mois précédent : ${prev.ouverts})
- Tickets résolus : ${cur.resolus} (mois précédent : ${prev.resolus})
- Tickets clos : ${cur.clos} (mois précédent : ${prev.clos})
- Tickets en retard : ${cur.en_retard} (mois précédent : ${prev.en_retard})

DÉLAIS MOYENS
- Délai moyen de résolution : ${cur.resolution_h.toFixed(2)} h (mois précédent : ${prev.resolution_h.toFixed(2)} h)
- Délai moyen de clôture : ${cur.cloture_h.toFixed(2)} h (mois précédent : ${prev.cloture_h.toFixed(2)} h)

TOP 10 CATÉGORIES (ouverts/résolus)
${catLines || "- (aucune donnée importée)"}

TOP 10 SERVICES DEMANDEURS (ouverts/résolus)
${srvLines || "- (aucune donnée importée)"}

CHARGE PAR TECHNICIEN
${techLines || "- (aucune donnée importée)"}

CONSIGNES DE RÉDACTION :
Rédige EXACTEMENT 7 paragraphes, chacun précédé de son titre en majuscules tel quel ci-dessous
(pour pouvoir copier chaque paragraphe dans la case correspondante de l'application) :

### 1. VOLUME GLOBAL DES TICKETS
(évolution du volume, lecture métier — hausse/baisse et ce que cela traduit)

### 2. ÉVOLUTION DES DÉLAIS DE TRAITEMENT
(amélioration ou dégradation des délais, contexte)

### 3. PRINCIPALES CATÉGORIES DE DEMANDES
(lecture métier des catégories dominantes, ce que cela révèle des besoins)

### 4. SERVICES LES PLUS DEMANDEURS
(quels services sollicitent le plus, pourquoi c'est notable)

### 5. CHARGE ET MOBILISATION DE L'ÉQUIPE SI
(répartition de la charge entre techniciens, équilibre ou déséquilibre)

### 6. ANALYSE COMPARATIVE ${prevLabel.toUpperCase()} / ${label.toUpperCase()}
(synthèse comparative des deux mois, tendance générale)

### 7. SYNTHÈSE FINALE ET CONCLUSION
(2-3 phrases de conclusion orientées Direction, aide à la décision)

Chaque paragraphe : texte uniquement (pas de tableau, pas de puces), ton sobre et factuel,
prêt à être collé directement dans l'application.`;
}

const MOUSSANADA_SECTIONS = [
  ["volume", "1. Volume global des tickets"],
  ["delays", "2. Évolution des délais de traitement"],
  ["categories", "3. Principales catégories de demandes"],
  ["services", "4. Services les plus demandeurs"],
  ["team", "5. Charge et mobilisation de l'équipe SI"],
  ["comparative", "6. Analyse comparative M-1 / M"],
  ["synthesis", "7. Synthèse finale et conclusion"],
];

function renderMSectionsForm(){
  const wrap = document.getElementById("mSectionsContainer");
  const sections = mState.notes.sections || {};
  wrap.innerHTML = MOUSSANADA_SECTIONS.map(([key,title])=>`
    <div class="card">
      <h2>${title}</h2>
      <p class="hint">Paragraphe Copilot pour cette section — inséré tel quel dans le rapport, sous les données correspondantes.</p>
      <textarea class="m-section-box" data-key="${key}" rows="4" placeholder="Coller ici le paragraphe Copilot...">${sections[key]||''}</textarea>
    </div>
  `).join("");
  wrap.querySelectorAll(".m-section-box").forEach(ta=>{
    ta.addEventListener("input", scheduleMNotesAutosave);
  });
}

function renderMAnalysis(){
  renderMPrompt();
  renderMSectionsForm();
  document.getElementById("mConstatsBox").value = (mState.notes.constats||[]).join("\n");
  document.getElementById("mRecoBox").value = (mState.notes.recommandations||[]).join("\n");
}

async function saveMNotes(silent){
  const sections = {};
  document.querySelectorAll(".m-section-box").forEach(ta=>{ sections[ta.dataset.key] = ta.value; });
  const payload = {
    sections,
    constats: document.getElementById("mConstatsBox").value.split("\n").map(s=>s.trim()).filter(Boolean),
    recommandations: document.getElementById("mRecoBox").value.split("\n").map(s=>s.trim()).filter(Boolean),
  };
  mState.notes = {...mState.notes, ...payload};
  await fetch(`/api/moussanada/notes/${mState.ym}`, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  if(silent) pulseAutosave("mAnalysisAutosaveHint");
}
const scheduleMNotesAutosave = debounce(()=>saveMNotes(true), 1800);
["mConstatsBox","mRecoBox"].forEach(id=>{
  document.getElementById(id).addEventListener("input", scheduleMNotesAutosave);
});
document.getElementById("mSaveNotesBtn").addEventListener("click", async ()=>{
  await saveMNotes(false);
  saveIndicatorPulse();
  toast("Analyse enregistrée");
});

document.querySelectorAll('.copy-btn[data-copy="mPrompt"]').forEach(btn=>{
  btn.addEventListener("click", ()=>{
    navigator.clipboard.writeText(document.getElementById("mPrompt").textContent).then(()=>toast("Prompt copié"));
  });
});

// ============================================================
// M-ANNUAL
// ============================================================
async function loadMAnnual(){
  const start = document.getElementById("mPeriodStart").value || `${new Date().getFullYear()}-01`;
  const end = document.getElementById("mPeriodEnd").value || `${new Date().getFullYear()}-12`;
  const year = start.split("-")[0];
  const res = await fetch(`/api/moussanada/annual/${year}?start=${start}&end=${end}`);
  const data = await res.json();
  const months = data.months;

  const totals = months.reduce((acc,m)=>{
    acc.ouverts+=m.ouverts; acc.resolus+=m.resolus; acc.clos+=m.clos; acc.en_retard+=m.en_retard;
    return acc;
  }, {ouverts:0,resolus:0,clos:0,en_retard:0});

  document.getElementById("mAnnualContent").innerHTML = `
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">🎫 Ouverts (période)</div><div class="kpi-value">${totals.ouverts}</div></div>
      <div class="kpi-card success"><div class="kpi-label">✅ Résolus (période)</div><div class="kpi-value">${totals.resolus}</div></div>
      <div class="kpi-card"><div class="kpi-label">📦 Clos (période)</div><div class="kpi-value">${totals.clos}</div></div>
      <div class="kpi-card danger"><div class="kpi-label">⚠️ En retard (cumul)</div><div class="kpi-value">${totals.en_retard}</div></div>
    </div>
    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Tendance annuelle</h3>
      <canvas id="mAnnualChartCanvas" height="90"></canvas>
    </div>
    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Évolution du backlog</h3>
      <canvas id="mBacklogAnnualCanvas" height="90"></canvas>
    </div>
    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Tendance des catégories (Top 8, sur la période)</h3>
      <canvas id="mCategoryTrendCanvas" height="100"></canvas>
    </div>
    <div class="dash-table-wrap">
      <h3>Détail par mois</h3>
      <table class="dash-table">
        <thead><tr><th>Mois</th><th>Ouverts</th><th>Résolus</th><th>En retard</th><th>Clos</th><th>Délai résolution</th><th>Délai clôture</th></tr></thead>
        <tbody>
          ${months.map(m=>`<tr><td>${m.label}</td><td>${m.ouverts}</td><td>${m.resolus}</td><td style="color:var(--danger);font-weight:700;">${m.en_retard}</td><td>${m.clos}</td><td>${formatDH(m.resolution_h)}</td><td>${formatDH(m.cloture_h)}</td></tr>`).join("") || `<tr><td colspan="7" style="color:var(--muted);">Aucune donnée</td></tr>`}
        </tbody>
      </table>
    </div>
  `;

  if(mAnnualChart) mAnnualChart.destroy();
  mAnnualChart = new Chart(document.getElementById("mAnnualChartCanvas"), {
    type:"line",
    data:{ labels: months.map(m=>m.label), datasets:[
      {label:"Ouverts", data: months.map(m=>m.ouverts), borderColor:"#0B4965", backgroundColor:"#0B496522", tension:.3},
      {label:"Résolus", data: months.map(m=>m.resolus), borderColor:"#25935F", backgroundColor:"#25935F22", tension:.3},
      {label:"Clos", data: months.map(m=>m.clos), borderColor:"#F59F0A", backgroundColor:"#F59F0A22", tension:.3},
    ]},
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}}, scales:{y:{beginAtZero:true}}}
  });

  if(mBacklogChart) mBacklogChart.destroy();
  mBacklogChart = new Chart(document.getElementById("mBacklogAnnualCanvas"), {
    type:"bar",
    data:{ labels: months.map(m=>m.label), datasets:[{label:"En retard", data: months.map(m=>m.en_retard), backgroundColor:"#DC2828", borderRadius:5}] },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}}}
  });
  // Tendance des catégories — endpoint dédié (s'appuie sur les sources mensuelles déjà stockées)
  const trendRes = await fetch(`/api/moussanada/category-trend?start=${start}&end=${end}`);
  const trend = await trendRes.json();
  const trendColors = ["#0B4965","#F59F0A","#25935F","#DC2828","#1794CF","#8B5CF6","#F2994A","#94A3AD"];
  if(mCategoryTrendChart) mCategoryTrendChart.destroy();
  mCategoryTrendChart = new Chart(document.getElementById("mCategoryTrendCanvas"), {
    type:"line",
    data:{ labels: trend.month_labels, datasets: trend.categories.map((cat,i)=>({
      label: cat, data: trend.series[cat], borderColor: trendColors[i % trendColors.length],
      backgroundColor: trendColors[i % trendColors.length] + "18", tension:.3, fill:false,
    })) },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom", labels:{boxWidth:10,font:{size:10.5}}}}, scales:{y:{beginAtZero:true}}}
  });

  addPngExportButtons({mAnnualChart, mBacklogChart, mCategoryTrendChart});
}

// ============================================================
// M-GUIDE (Guide de collecte GLPI)
// ============================================================
const GLPI_GUIDE = [
  {title:"Volume mensuel des tickets", path:"GLPI > Assistance > Statistiques > Tickets", filter:"Filtrer par période (mois en cours), grouper par statut (Ouvert/Résolu/Clos)", format:"Export Excel — feuille 'rawData - Tickets' (une ligne par mois)"},
  {title:"Délais de résolution", path:"GLPI > Rapports > Temps de traitement", filter:"Période = mois analysé, indicateur = durée moyenne de résolution/clôture", format:"Export Excel — feuille 'rawData - Durée moyenne - Heure'"},
  {title:"Catégories", path:"GLPI > Assistance > Tickets > Statistiques par catégorie", filter:"Période = mois analysé, arborescence complète des catégories", format:"Export Excel — feuille 'rawData - Ticket - Catégories'"},
  {title:"Services demandeurs", path:"GLPI > Assistance > Tickets > Groupes / Entités", filter:"Période = mois analysé, regrouper par entité/service demandeur", format:"Export Excel — feuille 'rawData - Ticket - Services'"},
  {title:"Techniciens", path:"GLPI > Assistance > Tickets > Statistiques par technicien", filter:"Période = mois analysé, technicien assigné", format:"Export Excel — feuille 'rawData - Technicien Assigné'"},
  {title:"Demandeurs", path:"GLPI > Assistance > Tickets > Statistiques par demandeur", filter:"Période = mois analysé", format:"Export Excel — feuille 'rawData - Ticket - Demandeur'"},
];

function renderMGuide(){
  document.getElementById("mGuideContent").innerHTML = `
    <div class="card">
      ${GLPI_GUIDE.map(g=>`
        <div class="guide-item">
          <h4>${g.title}</h4>
          <div class="g-path">${g.path}</div>
          <p><strong>Filtre :</strong> ${g.filter}</p>
          <p><strong>Format conseillé :</strong> ${g.format}</p>
        </div>
      `).join("")}
      <p class="hint">Une fois les exports réalisés, importez-les dans l'onglet <strong>Données sources</strong> — soit
      le fichier consolidé unique (.xlsx), soit chaque export séparément (.csv).</p>
    </div>
  `;
}

// ============================================================
// AUTHENTIFICATION
// ============================================================
function showAuthOverlay(setupRequired){
  document.getElementById("authOverlay").style.display = "flex";
  document.getElementById("setupForm").style.display = setupRequired ? "block" : "none";
  document.getElementById("loginForm").style.display = setupRequired ? "none" : "block";
}
function hideAuthOverlay(){
  document.getElementById("authOverlay").style.display = "none";
}

document.getElementById("setupSubmitBtn").addEventListener("click", async ()=>{
  const payload = {
    name: document.getElementById("setupName").value.trim(),
    username: document.getElementById("setupUsername").value.trim(),
    password: document.getElementById("setupPassword").value,
  };
  const res = await fetch("/api/auth/setup", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  const data = await res.json();
  if(data.ok){ currentUser = data.user; hideAuthOverlay(); startApp(); }
  else document.getElementById("setupError").textContent = data.error || "Erreur";
});

document.getElementById("loginSubmitBtn").addEventListener("click", async ()=>{
  const payload = {
    username: document.getElementById("loginUsername").value.trim(),
    password: document.getElementById("loginPassword").value,
  };
  const res = await fetch("/api/auth/login", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  const data = await res.json();
  if(data.ok){ currentUser = data.user; hideAuthOverlay(); startApp(); }
  else document.getElementById("loginError").textContent = data.error || "Erreur";
});

document.getElementById("logoutBtn").addEventListener("click", async ()=>{
  await fetch("/api/auth/logout", {method:"POST"});
  window.location.reload();
});

const ROLE_LABELS = {admin:"Administrateur", hotliner:"Hotliner", superviseur:"Superviseur"};

function applyUserMenu(){
  document.getElementById("userMenuName").textContent = currentUser.name;
  document.getElementById("userMenuRole").textContent = ROLE_LABELS[currentUser.role] || currentUser.role;
  document.getElementById("shareSanadBtn").style.display = currentUser.role === "admin" ? "" : "none";
}

async function startApp(){
  applyUserMenu();
  await loadModules();
  await loadGlobalSettings();
  applySanadBrand();
  if(currentUser.role !== "hotliner"){
    await loadModuleSettings();
    setupPeriodPicker("periodPresets", "periodStart", "periodEnd", ()=>{ if(document.getElementById("tab-annual").classList.contains("active")) loadAnnual(); });
    setupPeriodPicker("mPeriodPresets", "mPeriodStart", "mPeriodEnd", ()=>{ if(document.getElementById("tab-m-annual").classList.contains("active")) loadMAnnual(); });
    setupPeriodPickerDate("pPeriodPresets", "pPeriodStart", "pPeriodEnd", ()=>{ if(document.getElementById("tab-p-dashboard").classList.contains("active")) renderPDashboard(); });
    if(currentModule === "moussanada"){
      await loadMoussanadaMonth(defaultMonth());
    } else {
      await loadMonth(defaultMonth());
    }
  }
  activateTab(defaultTabForRole());
  maybeShowOnboarding();
}

// ---------- Onboarding contextuel par rôle (Lot 5) ----------
const ONBOARDING_CONTENT = {
  admin: {
    title: "Bienvenue, Administrateur",
    html: `<p>Vous avez accès complet à SANAD. Points de départ utiles :</p>
      <ul style="padding-left:18px;margin:10px 0;">
        <li><strong>Saisie quotidienne</strong> : onglets Appels/Emails (Tarkhiss) ou Import GLPI (Moussanada)</li>
        <li><strong>Rapports</strong> : chaque dashboard propose Excel, PDF, aperçu avant envoi, et envoi direct par email</li>
        <li><strong>Administration</strong> : gestion des comptes, personnalisation des titres/logos, sécurité, RGPD, et journal d'audit</li>
        <li><strong>Reporting Métier</strong> se trouve désormais comme sous-volet de Tarkhiss dans la sidebar</li>
      </ul>
      <p style="color:var(--muted);font-size:12.5px;">Vous pouvez revoir cette aide à tout moment via le bouton "? Aide" en haut à droite.</p>`,
  },
  superviseur: {
    title: "Bienvenue, Superviseur",
    html: `<p>Vous disposez d'un accès en lecture seule à l'ensemble des dashboards et exports.</p>
      <ul style="padding-left:18px;margin:10px 0;">
        <li>Consultez les <strong>Dashboards</strong> et <strong>Vues annuelles</strong> de chaque volet</li>
        <li>Téléchargez les rapports en <strong>Excel</strong> ou <strong>PDF</strong> librement</li>
        <li>Les comparatifs N-1 et le tableau "Détail par mois" sont accessibles à votre rôle</li>
        <li>La saisie, l'envoi d'email et l'administration restent réservés à l'Administrateur</li>
      </ul>`,
  },
  hotliner: {
    title: "Bienvenue",
    html: `<p>Votre accès est limité à la <strong>Base de connaissances</strong> et aux <strong>Notes</strong>, sur les volets qui vous ont été assignés.</p>
      <ul style="padding-left:18px;margin:10px 0;">
        <li>Consultez la Base de connaissances pour les procédures courantes</li>
        <li>Utilisez les Notes pour échanger avec l'Administrateur (ex. pendant un appel)</li>
      </ul>`,
  },
};

function maybeShowOnboarding(){
  const role = currentUser?.role;
  if(!role || !ONBOARDING_CONTENT[role]) return;
  const key = `sanad_onboarding_seen_${role}`;
  if(localStorage.getItem(key)) return;
  showOnboarding(role);
  localStorage.setItem(key, "1");
}

function showOnboarding(role){
  const content = ONBOARDING_CONTENT[role || currentUser?.role];
  if(!content) return;
  document.getElementById("onboardingTitle").textContent = content.title;
  document.getElementById("onboardingBody").innerHTML = content.html;
  document.getElementById("onboardingOverlay").style.display = "flex";
}

document.getElementById("closeOnboardingBtn").addEventListener("click", ()=>{
  document.getElementById("onboardingOverlay").style.display = "none";
});
document.getElementById("onboardingOverlay").addEventListener("click", (e)=>{
  if(e.target.id === "onboardingOverlay") document.getElementById("onboardingOverlay").style.display = "none";
});
document.getElementById("helpBtn").addEventListener("click", ()=> showOnboarding());

// ---------- Densité compacte (Lot 5+) ----------
function applyDensity(density){
  const d = density === "compact" ? "compact" : "comfortable";
  if(d === "compact") document.documentElement.dataset.density = "compact";
  else delete document.documentElement.dataset.density;
  const btn = document.getElementById("densityToggleBtn");
  if(btn) btn.textContent = d === "compact" ? "▤ Confortable" : "▤ Compact";
}
applyDensity(localStorage.getItem("sanad_density") || "comfortable");
renderFavorites();
document.getElementById("favoriteToggleBtn").addEventListener("click", toggleCurrentFavorite);
document.getElementById("densityToggleBtn").addEventListener("click", ()=>{
  const next = document.documentElement.dataset.density === "compact" ? "comfortable" : "compact";
  localStorage.setItem("sanad_density", next);
  applyDensity(next);
});

// ============================================================
// UTILISATEURS (admin)
// ============================================================
function renderNewUserModules(){
  const wrap = document.getElementById("newUserModulesList");
  wrap.innerHTML = Object.entries(modulesInfo).map(([key,info])=>`
    <label style="margin-right:14px;font-size:12.5px;"><input type="checkbox" class="newUserModuleCb" value="${key}" checked> ${info.label}</label>
  `).join("");
}

document.getElementById("newUserRole").addEventListener("change", (e)=>{
  document.getElementById("newUserModulesRow").style.display = e.target.value === "hotliner" ? "block" : "none";
});

document.getElementById("createUserBtn").addEventListener("click", async ()=>{
  const role = document.getElementById("newUserRole").value;
  const modules = [...document.querySelectorAll(".newUserModuleCb:checked")].map(c=>c.value);
  const payload = {
    name: document.getElementById("newUserName").value.trim(),
    username: document.getElementById("newUserUsername").value.trim(),
    password: document.getElementById("newUserPassword").value,
    role, modules: role === "hotliner" ? modules : Object.keys(modulesInfo),
  };
  const res = await fetch("/api/users", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  const data = await res.json();
  if(data.ok){
    toast("Compte créé");
    document.getElementById("newUserName").value = "";
    document.getElementById("newUserUsername").value = "";
    document.getElementById("newUserPassword").value = "";
    loadUsers();
  } else {
    toast(data.error || "Erreur");
  }
});

async function loadUsers(){
  renderNewUserModules();
  const res = await fetch("/api/users");
  const users = await res.json();
  document.querySelector("#usersTable tbody").innerHTML = users.map(u=>`
    <tr>
      <td>${u.name}</td>
      <td>${u.username}</td>
      <td>${ROLE_LABELS[u.role]||u.role}</td>
      <td style="font-size:12px;">${u.role==='hotliner' ? (u.modules||[]).map(m=>modulesInfo[m]?.label||m).join(", ") : "Tous"}</td>
      <td><span class="status-dot ${u.active?'on':'off'}"></span>${u.active?'Actif':'Désactivé'}</td>
      <td>
        ${u.id !== currentUser.id ? `<button class="admin-del-btn" data-action="toggle" data-id="${u.id}" data-active="${u.active}">${u.active?'Désactiver':'Réactiver'}</button>` : ''}
        <button class="admin-del-btn" data-action="reset" data-id="${u.id}" style="color:var(--primary);">Réinit. mdp</button>
      </td>
    </tr>
  `).join("") || `<tr><td colspan="6" style="color:var(--muted);">Aucun compte</td></tr>`;

  document.querySelectorAll('#usersTable [data-action="toggle"]').forEach(b=>{
    b.addEventListener("click", async ()=>{
      await fetch(`/api/users/${b.dataset.id}`, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({active: b.dataset.active !== "true"})});
      loadUsers();
    });
  });
  document.querySelectorAll('#usersTable [data-action="reset"]').forEach(b=>{
    b.addEventListener("click", async ()=>{
      const pw = prompt("Nouveau mot de passe (4 caractères min.) :");
      if(!pw || pw.length < 4) return;
      await fetch(`/api/users/${b.dataset.id}`, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({new_password: pw})});
      toast("Mot de passe réinitialisé");
    });
  });
}

// ============================================================
// NOTES (Hotliner <-> Admin)
// ============================================================
function renderNoteBubbles(thread){
  document.getElementById("notesThread").innerHTML = thread.map(n=>`
    <div class="note-bubble ${n.from}">
      ${esc(n.text)}
      <div class="note-meta">${n.from === 'admin' ? 'Responsable' : 'Hotliner'} · ${new Date(n.date).toLocaleString('fr-FR')}</div>
    </div>
  `).join("") || `<p class="hint">Aucune note pour le moment.</p>`;
  const t = document.getElementById("notesThread");
  t.scrollTop = t.scrollHeight;
}

let notesSelectedHotlinerId = null;

async function loadNotes(){
  if(currentUser.role === "hotliner"){
    document.getElementById("notesHotlinerSelectCard").style.display = "none";
    const res = await fetch("/api/notes/mine");
    renderNoteBubbles(await res.json());
  } else {
    document.getElementById("notesHotlinerSelectCard").style.display = "block";
    const res = await fetch("/api/notes/hotliners");
    const hotliners = await res.json();
    const sel = document.getElementById("notesHotlinerSelect");
    sel.innerHTML = hotliners.map(h=>`<option value="${h.id}">${h.name}</option>`).join("") || `<option value="">Aucun hotliner</option>`;
    if(!notesSelectedHotlinerId && hotliners.length) notesSelectedHotlinerId = hotliners[0].id;
    sel.value = notesSelectedHotlinerId || "";
    sel.onchange = async ()=>{ notesSelectedHotlinerId = Number(sel.value); await loadAdminNotesThread(); };
    await loadAdminNotesThread();
  }
}

async function loadAdminNotesThread(){
  if(!notesSelectedHotlinerId){ renderNoteBubbles([]); return; }
  const res = await fetch(`/api/notes/user/${notesSelectedHotlinerId}`);
  renderNoteBubbles(await res.json());
}

document.getElementById("notesSendBtn").addEventListener("click", async ()=>{
  const text = document.getElementById("notesInput").value.trim();
  if(!text) return;
  if(currentUser.role === "hotliner"){
    const res = await fetch("/api/notes/mine", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({text})});
    const data = await res.json();
    renderNoteBubbles(data.thread);
  } else {
    if(!notesSelectedHotlinerId){ toast("Choisissez un hotliner"); return; }
    const res = await fetch(`/api/notes/user/${notesSelectedHotlinerId}`, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({text})});
    const data = await res.json();
    renderNoteBubbles(data.thread);
  }
  document.getElementById("notesInput").value = "";
});

// ============================================================
// BASE DE CONNAISSANCES
// ============================================================
async function loadKnowledge(){
  const modSel = document.getElementById("knowledgeModule");
  const allowed = (currentUser.role === "hotliner") ? currentUser.modules : Object.keys(modulesInfo);
  modSel.innerHTML = Object.entries(modulesInfo).filter(([k])=>allowed.includes(k)).map(([k,i])=>`<option value="${k}">${i.label}</option>`).join("");
  document.getElementById("knowledgeAddCard").style.display = currentUser.role === "admin" ? "block" : "none";
  await renderKnowledgeList();
}

async function renderKnowledgeList(filterText=""){
  const res = await fetch("/api/knowledge");
  let items = await res.json();
  const allowed = (currentUser.role === "hotliner") ? currentUser.modules : Object.keys(modulesInfo);
  items = items.filter(i=>allowed.includes(i.module));
  if(filterText){
    const q = filterText.toLowerCase();
    items = items.filter(i=> i.title.toLowerCase().includes(q) || i.content.toLowerCase().includes(q));
  }
  document.getElementById("knowledgeList").innerHTML = items.map(i=>`
    <div class="kb-card">
      <div class="kb-mod">${modulesInfo[i.module]?.label || i.module}</div>
      <h4>${esc(i.title)}</h4>
      <p>${esc(i.content)}</p>
      ${currentUser.role === "admin" ? `
      <div class="kb-actions">
        <button data-action="edit" data-id="${i.id}">✏️ Modifier</button>
        <button data-action="delete" data-id="${i.id}">🗑️ Supprimer</button>
      </div>` : ""}
    </div>
  `).join("") || `<p class="hint">Aucun article.</p>`;

  document.querySelectorAll('#knowledgeList [data-action="edit"]').forEach(b=>{
    b.addEventListener("click", async ()=>{
      const item = items.find(i=>i.id == b.dataset.id);
      document.getElementById("knowledgeEditId").value = item.id;
      document.getElementById("knowledgeModule").value = item.module;
      document.getElementById("knowledgeTitle").value = item.title;
      document.getElementById("knowledgeContent").value = item.content;
      document.getElementById("knowledgeCancelBtn").style.display = "inline-block";
      window.scrollTo(0,0);
    });
  });
  document.querySelectorAll('#knowledgeList [data-action="delete"]').forEach(b=>{
    b.addEventListener("click", async ()=>{
      if(!confirm("Supprimer cet article ?")) return;
      await fetch(`/api/knowledge/${b.dataset.id}`, {method:"DELETE"});
      renderKnowledgeList(document.getElementById("knowledgeSearch").value);
    });
  });
}

document.getElementById("knowledgeSearch").addEventListener("input", (e)=> renderKnowledgeList(e.target.value));

document.getElementById("knowledgeSaveBtn").addEventListener("click", async ()=>{
  const payload = {
    id: document.getElementById("knowledgeEditId").value ? Number(document.getElementById("knowledgeEditId").value) : null,
    module: document.getElementById("knowledgeModule").value,
    title: document.getElementById("knowledgeTitle").value.trim(),
    content: document.getElementById("knowledgeContent").value.trim(),
  };
  if(!payload.title){ toast("Titre requis"); return; }
  await fetch("/api/knowledge", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  document.getElementById("knowledgeEditId").value = "";
  document.getElementById("knowledgeTitle").value = "";
  document.getElementById("knowledgeContent").value = "";
  document.getElementById("knowledgeCancelBtn").style.display = "none";
  toast("Article enregistré");
  renderKnowledgeList();
});
document.getElementById("knowledgeCancelBtn").addEventListener("click", ()=>{
  document.getElementById("knowledgeEditId").value = "";
  document.getElementById("knowledgeTitle").value = "";
  document.getElementById("knowledgeContent").value = "";
  document.getElementById("knowledgeCancelBtn").style.display = "none";
});

// ============================================================
// RECHERCHE GLOBALE
// ============================================================
async function runGlobalSearch(q){
  const results = document.getElementById("globalSearchResults");
  if(!q || q.length < 2){ results.classList.remove("show"); return; }
  const ql = q.toLowerCase();
  const items = [];

  if(currentUser.role !== "hotliner"){
    if(currentModule === "tarkhiss"){
      (state.analysis?.problems||[]).forEach(p=>{ if(p.label.toLowerCase().includes(ql)) items.push({cat:"Problème", label:p.label, action:()=>{activateTab("analysis");}}); });
      (state.analysis?.demandes||[]).forEach(p=>{ if(p.label.toLowerCase().includes(ql)) items.push({cat:"Demande", label:p.label, action:()=>{activateTab("analysis");}}); });
    } else {
      (mState.month?.categories||[]).forEach(c=>{ if(c.label.toLowerCase().includes(ql)) items.push({cat:"Catégorie GLPI", label:c.label, action:()=>{activateTab("m-sources");}}); });
      (mState.month?.services||[]).forEach(c=>{ if(c.label.toLowerCase().includes(ql)) items.push({cat:"Service GLPI", label:c.label, action:()=>{activateTab("m-sources");}}); });
      (mState.month?.techniciens||[]).forEach(c=>{ if(c.label.toLowerCase().includes(ql)) items.push({cat:"Technicien", label:c.label, action:()=>{activateTab("m-sources");}}); });
    }
    (moduleSettings.contacts||[]).forEach(c=>{
      if((c.name||"").toLowerCase().includes(ql) || c.email.toLowerCase().includes(ql)) items.push({cat:"Contact", label:`${c.name||''} ${c.email}`, action:()=>{activateTab("admin");}});
    });
    try{
      const monthsRes = await fetch(apiM("/months-list"));
      const months = await monthsRes.json();
      months.forEach(m=>{ if(m.label.toLowerCase().includes(ql)) items.push({cat:"Mois", label:m.label, action:()=>{ if(currentModule==='tarkhiss'){loadMonth(m.ym);activateTab("dashboard");} else {loadMoussanadaMonth(m.ym);activateTab("m-dashboard");} }}); });
    } catch(e){}
  }

  const kbRes = await fetch("/api/knowledge");
  const kb = await kbRes.json();
  kb.forEach(k=>{ if(k.title.toLowerCase().includes(ql)) items.push({cat:"Base de connaissances", label:k.title, action:()=>{activateTab("knowledge"); document.getElementById("knowledgeSearch").value=k.title; renderKnowledgeList(k.title);}}); });

  try{
    const notesRes = await fetch(`/api/notes/search?q=${encodeURIComponent(q)}`);
    const notes = await notesRes.json();
    notes.forEach(n=>{
      const preview = n.text.length > 60 ? n.text.slice(0, 60) + "…" : n.text;
      items.push({cat:`Note — ${n.hotliner_name}`, label: preview, action:()=>{
        activateTab("notes");
        if(currentUser.role !== "hotliner"){ notesSelectedHotlinerId = n.hotliner_id; loadNotes(); }
      }});
    });
  } catch(e){ /* silencieux */ }

  results.innerHTML = items.slice(0,12).map((it,i)=>`<div class="gsr-item" data-i="${i}"><div class="gsr-cat">${it.cat}</div>${it.label}</div>`).join("") || `<div class="gsr-item" style="color:var(--muted);">Aucun résultat</div>`;
  results.querySelectorAll(".gsr-item[data-i]").forEach((el,i)=>{
    el.addEventListener("click", ()=>{ items[i].action(); results.classList.remove("show"); document.getElementById("globalSearch").value=""; });
  });
  results.classList.add("show");
}
document.getElementById("globalSearch").addEventListener("input", (e)=> runGlobalSearch(e.target.value));
document.addEventListener("click", (e)=>{
  if(!e.target.closest(".global-search-wrap")) document.getElementById("globalSearchResults").classList.remove("show");
});

// ============================================================
// EXPORT PNG DES GRAPHIQUES
// ============================================================
function addPngExportButtons(chartRegistry){
  document.querySelectorAll(".chart-card canvas").forEach(canvas=>{
    const card = canvas.closest(".chart-card");
    if(!card || card.querySelector(".chart-png-btn")) return;
    const btn = document.createElement("button");
    btn.className = "chart-png-btn";
    btn.textContent = "⬇ PNG";
    btn.addEventListener("click", (e)=>{
      e.stopPropagation();
      const chart = chartRegistry[canvas.id] || Object.values(chartRegistry).find(c=>c && c.canvas && c.canvas.id === canvas.id);
      if(chart && chart.toBase64Image){
        const a = document.createElement("a");
        a.href = chart.toBase64Image();
        a.download = (canvas.id || "graphique") + ".png";
        a.click();
      }
    });
    card.appendChild(btn);
  });
}

// ============================================================
// SEUILS D'ALERTE
// ============================================================
function computeAlerts(metrics, thresholds){
  const alerts = [];
  const th = thresholds || {};
  if(th.resolution_rate_min != null && metrics.resolutionRate != null && metrics.resolutionRate < th.resolution_rate_min){
    alerts.push(`Taux de résolution (${metrics.resolutionRate.toFixed(0)}%) sous le seuil de ${th.resolution_rate_min}%`);
  }
  if(th.backlog_max != null && metrics.backlog != null && metrics.backlog > th.backlog_max){
    alerts.push(`Stock en retard (${metrics.backlog}) au-dessus du seuil de ${th.backlog_max}`);
  }
  if(th.delay_max_h != null && metrics.delayH != null && metrics.delayH > th.delay_max_h){
    alerts.push(`Délai moyen (${metrics.delayH.toFixed(1)} h) au-dessus du seuil de ${th.delay_max_h} h`);
  }
  if(th.volume_variation_max != null && metrics.volumeCur != null && metrics.volumePrev){
    const diff = Math.abs(((metrics.volumeCur - metrics.volumePrev) / metrics.volumePrev) * 100);
    if(diff > th.volume_variation_max){
      const sign = metrics.volumeCur > metrics.volumePrev ? "hausse" : "baisse";
      alerts.push(`Variation de volume anormale : ${sign} de ${diff.toFixed(0)}% (seuil : ${th.volume_variation_max}%)`);
    }
  }
  return alerts;
}

async function logAlertHistory(module, ym, alerts){
  if(!alerts || !alerts.length) return;
  try{
    await fetch(`/api/${module}/alert-history`, {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ym, alerts})});
  }catch(e){}
}

function renderAlertBanner(containerId, alerts){
  const el = document.getElementById(containerId);
  if(!el) return;
  if(alerts.length){
    el.style.display = "block";
    el.className = "banner danger";
    el.innerHTML = "⚠️ <strong>Seuils dépassés :</strong> " + alerts.join(" · ");
  } else {
    el.style.display = "none";
  }
}

// ============================================================
// DOSSIERS PCHC
// ============================================================
let pState = { start: null, end: null, data: null };
let pCharts = {};

function periodPresetRangesDate(){
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth();
  const todayStr = now.toISOString().slice(0,10);
  const qStartMonth = Math.floor(m/3)*3;
  const hStartMonth = Math.floor(m/6)*6;
  return {
    month: { start: `${y}-${pad2(m+1)}-01`, end: todayStr, label: "Ce mois-ci" },
    quarter: { start: `${y}-${pad2(qStartMonth+1)}-01`, end: todayStr, label: "Ce trimestre" },
    semester: { start: `${y}-${pad2(hStartMonth+1)}-01`, end: todayStr, label: "Ce semestre" },
    year: { start: `${y}-01-01`, end: todayStr, label: "Cette année" },
    all: { start: "2000-01-01", end: todayStr, label: "Tout" },
  };
}

function setupPeriodPickerDate(presetsId, startId, endId, onChange){
  const presets = periodPresetRangesDate();
  const wrap = document.getElementById(presetsId);
  const startInput = document.getElementById(startId);
  const endInput = document.getElementById(endId);
  wrap.innerHTML = Object.entries(presets).map(([key,p])=>`<button data-key="${key}">${p.label}</button>`).join("");
  function applyPreset(key){
    const p = presets[key];
    startInput.value = p.start;
    endInput.value = p.end;
    wrap.querySelectorAll("button").forEach(b=>b.classList.toggle("active", b.dataset.key===key));
    onChange();
  }
  wrap.querySelectorAll("button").forEach(b=> b.addEventListener("click", ()=>applyPreset(b.dataset.key)));
  [startInput, endInput].forEach(inp=>{
    inp.addEventListener("change", ()=>{
      wrap.querySelectorAll("button").forEach(b=>b.classList.remove("active"));
      onChange();
    });
  });
  applyPreset("year");
}

async function loadPchcData(){
  const start = document.getElementById("pPeriodStart").value || "2000-01-01";
  const end = document.getElementById("pPeriodEnd").value || new Date().toISOString().slice(0,10);
  pState.start = start; pState.end = end;
  const res = await fetch(`/api/pchc/data?start=${start}&end=${end}`);
  pState.data = await res.json();
}

const VOYANT_HEX = PCHC_PALETTE;
const PCHC_CAT_LABELS = {
  identification:"Identification Opérateur", declaration:"Déclaration d'Activité",
  ce:"Certificat d'Enregistrement", clv:"Certificat de Libre Vente",
  aimp:"Autorisation d'Importation Matières Premières",
};

async function renderPDashboard(){
  await loadPchcData();
  const d = pState.data;
  const s = d.summary;
  document.getElementById("pDashTitle").textContent = `Reporting Métier — PCHC — ${pState.start} au ${pState.end}`;
  renderSpecializationSwitch();

  renderAlertBanner("pAlertBanner", d.alerts);

  const catCards = Object.entries(d.categories).map(([key,cat])=>`
    <div class="dash-table-wrap pchc-cat-card" style="border-left-color:${VOYANT_HEX[cat.voyant]};">
      <h3>${cat.voyant==='green'?'🟢':cat.voyant==='yellow'?'🟡':'🔴'} ${esc(cat.label)}</h3>
      <div class="kpi-grid" style="margin-bottom:0;">
        <div class="kpi-card"><div class="kpi-label">Total</div><div class="kpi-value" style="font-size:20px;">${cat.total}</div></div>
        <div class="kpi-card success"><div class="kpi-label">Délivrés</div><div class="kpi-value" style="font-size:20px;">${cat.delivered}</div></div>
        <div class="kpi-card"><div class="kpi-label">En cours</div><div class="kpi-value" style="font-size:20px;color:var(--info);">${cat.encours}</div></div>
        <div class="kpi-card danger"><div class="kpi-label">Bloqués</div><div class="kpi-value" style="font-size:20px;">${cat.bloque}</div></div>
      </div>
      <div style="margin-top:10px;font-size:13px;color:var(--muted);">Taux de délivrance : <strong style="color:var(--ink);">${cat.taux}%</strong></div>
    </div>
  `).join("");

  const logoImg = globalSettings.logo_filename ? `<img src="/static/uploads/${globalSettings.logo_filename}" class="print-only" style="height:34px;margin-bottom:10px;">` : "";

  document.getElementById("pDashboardContent").innerHTML = `
    <div class="print-only" style="margin-bottom:10px;">
      ${logoImg}
      <div style="font-size:11px;color:#5B6472;">${globalSettings.agency_name || ""} · Reporting Métier (PCHC) · ${pState.start} au ${pState.end}</div>
    </div>

    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">📁 Total dossiers</div><div class="kpi-value">${s.total}</div></div>
      <div class="kpi-card success"><div class="kpi-label">✅ Délivrés / Validés</div><div class="kpi-value">${s.delivered}</div></div>
      <div class="kpi-card"><div class="kpi-label">⏳ En cours</div><div class="kpi-value" style="color:var(--info);">${s.encours}</div></div>
      <div class="kpi-card danger"><div class="kpi-label">🔴 Bloqués / Refusés</div><div class="kpi-value">${s.bloque}</div></div>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">Taux de délivrance global</div><div class="kpi-value">${s.taux}%</div></div>
      <div class="kpi-card"><div class="kpi-label">Dossiers &gt;90 jours</div><div class="kpi-value" style="color:var(--danger);">${d.backlog['>90']}</div></div>
      <div class="kpi-card"><div class="kpi-label">Dernier import</div><div class="kpi-value" style="font-size:14px;">${d.last_import ? new Date(d.last_import).toLocaleString('fr-FR') : '—'}</div></div>
    </div>

    <div class="charts-row">
      <div class="chart-card">
        <h3>Répartition globale des statuts</h3>
        <canvas id="pSummaryChart" height="150"></canvas>
      </div>
      <div class="chart-card">
        <h3>Ancienneté du backlog (dossiers en cours)</h3>
        <canvas id="pBacklogChart" height="150"></canvas>
      </div>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Synthèse par module — volume &amp; taux de délivrance</h3>
      <canvas id="pCategoriesChart" height="120"></canvas>
    </div>

    <div class="chart-card" style="margin-bottom:16px;">
      <h3>Top 10 opérateurs par volume de dossiers</h3>
      <canvas id="pTop10Chart" height="140"></canvas>
    </div>

    ${catCards}

    ${pState.notesText ? `<div class="dash-table-wrap synth-card"><h3>📝 Analyse</h3>${pState.notesText.split("\n").filter(p=>p.trim()).map(p=>`<p style="font-size:13px;color:var(--ink);line-height:1.6;">${esc(p)}</p>`).join("")}</div>` : ""}
  `;

  Object.values(pCharts).forEach(c=>c && c.destroy());

  const voyantColors = PCHC_PALETTE;
  const globalStatusCounts = {green:0,blue:0,yellow:0,orange:0,red:0};
  Object.values(d.categories).forEach(cat=>{
    Object.values(cat.status_counts).forEach(sc=>{ globalStatusCounts[sc.color] = (globalStatusCounts[sc.color]||0) + sc.count; });
  });
  pCharts.summary = new Chart(document.getElementById("pSummaryChart"), {
    type:"doughnut",
    data:{ labels:["Délivré/Validé","En cours (bleu)","Non conforme","Attente","Irrecevable/Rejeté"],
      datasets:[{ data:[globalStatusCounts.green,globalStatusCounts.blue,globalStatusCounts.yellow,globalStatusCounts.orange,globalStatusCounts.red],
        backgroundColor:[voyantColors.green,voyantColors.blue,voyantColors.yellow,voyantColors.orange,voyantColors.red] }] },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom",labels:{boxWidth:10,font:{size:10.5}}}}}
  });

  pCharts.backlog = new Chart(document.getElementById("pBacklogChart"), {
    type:"bar",
    data:{ labels:["0-30j","31-60j","61-90j",">90j"], datasets:[{label:"Dossiers en cours", data:[d.backlog["0-30"],d.backlog["31-60"],d.backlog["61-90"],d.backlog[">90"]], backgroundColor:[voyantColors.green,voyantColors.yellow,voyantColors.orange,voyantColors.red], borderRadius:5}] },
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{display:false}, datalabels:{display:true,anchor:'end',align:'top',font:{weight:'bold',size:11},color:'#0D1926'}}, scales:{y:{beginAtZero:true,grace:'15%'}}}
  });

  const catEntries = Object.entries(d.categories);
  pCharts.categories = new Chart(document.getElementById("pCategoriesChart"), {
    data:{ labels: catEntries.map(([k,c])=>PCHC_CAT_LABELS[k]||k), datasets:[
      {type:"bar", label:"Délivrés", data: catEntries.map(([k,c])=>c.delivered), backgroundColor:voyantColors.green, borderRadius:4, yAxisID:"y"},
      {type:"bar", label:"En cours", data: catEntries.map(([k,c])=>c.encours), backgroundColor:voyantColors.blue, borderRadius:4, yAxisID:"y"},
      {type:"bar", label:"Bloqués", data: catEntries.map(([k,c])=>c.bloque), backgroundColor:voyantColors.red, borderRadius:4, yAxisID:"y"},
      {type:"line", label:"Taux de délivrance", data: catEntries.map(([k,c])=>c.taux), borderColor:"#F59F0A", backgroundColor:"#F59F0A", borderWidth:2.5, pointRadius:4, pointBackgroundColor:"#F59F0A", tension:.3, yAxisID:"y1"},
    ]},
    options:{responsive:true, animation:{duration:700}, plugins:{legend:{position:"bottom"}, datalabels:{display:false}},
      scales:{
        y:{beginAtZero:true, position:"left", title:{display:true, text:"Nombre de dossiers"}},
        y1:{beginAtZero:true, max:100, position:"right", grid:{drawOnChartArea:false}, title:{display:true, text:"Taux (%)"}, ticks:{callback:v=>v+"%"}},
      }}
  });

  const top10 = d.top10.total.slice(0,10);
  pCharts.top10 = new Chart(document.getElementById("pTop10Chart"), {
    type:"bar",
    data:{ labels: top10.map(t=>t.entity), datasets:[{label:"Dossiers", data: top10.map(t=>t.count), backgroundColor:"#0B4965", borderRadius:4}] },
    options:{indexAxis:"y", responsive:true, animation:{duration:700}, plugins:{legend:{display:false}, datalabels:{display:true,anchor:'end',align:'end',font:{weight:'bold',size:10},color:'#0D1926'}}, scales:{x:{beginAtZero:true,grace:'15%'}}}
  });

  setupEmailPanel("pToCheckList","pCcCheckList","pGreetingInput");
  addPngExportButtons(pCharts);
}

document.getElementById("pPdfBtn").addEventListener("click", ()=>{
  downloadPdfFromServer("/api/pchc/export-pdf", {start: pState.start, end: pState.end, charts: collectPchcChartImages()}, `rapport_pchc_${pState.start}_${pState.end}.pdf`);
});
document.getElementById("pXlsxBtn").addEventListener("click", ()=>{
  window.location.href = `/api/pchc/export-xlsx?start=${pState.start}&end=${pState.end}`;
});
document.getElementById("pCsvRawBtn").addEventListener("click", ()=>{
  window.location.href = `/api/pchc/export-csv-raw?start=${pState.start}&end=${pState.end}`;
});

function collectPchcChartImages(){
  const map = {summary:"summary", backlog:"backlog", top10:"top10", categories:"categories"};
  const out = {};
  Object.entries(map).forEach(([key, chartKey])=>{
    const chart = pCharts[chartKey];
    if(chart && chart.toBase64Image){
      try{ out[key] = chart.toBase64Image('image/png', 1); }catch(e){}
    }
  });
  return out;
}

document.getElementById("pPptxBtn").addEventListener("click", async ()=>{
  toast("Génération de la présentation...");
  const charts = collectPchcChartImages();
  const res = await fetch("/api/pchc/export-pptx", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({start: pState.start, end: pState.end, charts})
  });
  if(!res.ok){ toast("Erreur génération PPTX"); return; }
  const blob = await res.blob();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `SANAD_PCHC_RevueDirection_${pState.end}.pptx`;
  a.click();
});

document.getElementById("pMailToggleBtn").addEventListener("click", ()=>{
  const panel = document.getElementById("pEmailPanel");
  panel.style.display = panel.style.display === "none" ? "block" : "none";
});
document.getElementById("pPreviewMailBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients("pToCheckList","pCcCheckList");
  const greeting = document.getElementById("pGreetingInput").value.trim();
  const charts = collectPchcChartImages();
  const res = await fetch("/api/pchc/preview-email", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({to, cc, greeting, start: pState.start, end: pState.end, charts})
  });
  const data = await res.json();
  openPreviewModal(data.subject, data.html);
});
document.getElementById("pConfirmSendBtn").addEventListener("click", async ()=>{
  const {to, cc} = getSelectedRecipients("pToCheckList","pCcCheckList");
  if(!to.length){ toast("Sélectionnez au moins un destinataire (À)"); return; }
  const greeting = document.getElementById("pGreetingInput").value.trim();
  const charts = collectPchcChartImages();
  toast("Génération du rapport et de la pièce jointe...");
  const res = await fetch("/api/pchc/send-email", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({to, cc, greeting, start: pState.start, end: pState.end, charts})
  });
  const data = await res.json();
  toast(data.opened ? "Ouverture dans Outlook..." : "Fichier .eml généré : " + data.file);
  document.getElementById("pEmailPanel").style.display = "none";
});

// ---------- P-IMPORT ----------
let pSpecializations = null;
async function renderSpecializationSwitch(){
  if(!pSpecializations){
    const res = await fetch("/api/pchc/specializations");
    pSpecializations = await res.json();
  }
  const html = Object.entries(pSpecializations).map(([key,spec])=>
    `<button ${spec.ready?'class="active"':'disabled title="Bientôt disponible"'} data-key="${key}">${spec.label}${spec.ready?'':' 🔒'}</button>`
  ).join("");
  ["pSpecializationSwitch","pSpecializationSwitchImport"].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.innerHTML = html;
  });
}

async function renderPImport(){
  const res = await fetch("/api/pchc/data?start=2000-01-01&end=2100-01-01");
  const d = await res.json();
  document.getElementById("pLastImport").textContent = d.last_import ? new Date(d.last_import).toLocaleString('fr-FR') : "Aucun import";
  await renderSpecializationSwitch();
  await renderPStatusColorsTable();
}

function showDetectedStatuses(data){
  const banner = document.getElementById("pNewStatusBanner");
  if(!data.statuses_found || !data.statuses_found.length){
    banner.style.display = "none";
    return;
  }
  const foundList = data.statuses_found.map(s=>esc(s)).join(", ");
  if(data.new_statuses && data.new_statuses.length){
    banner.className = "banner danger";
    banner.innerHTML = `⚠️ <strong>${data.new_statuses.length} nouveau(x) statut(s) détecté(s)</strong> dans le fichier importé, non encore mappé(s) (affichés en gris par défaut) :
      <strong>${data.new_statuses.map(s=>esc(s)).join(", ")}</strong>.
      Rendez-vous dans le tableau "Statut → Couleur" ci-dessous pour les classer.
      <div class="hint" style="margin-top:6px;">Tous les statuts détectés (${data.statuses_found.length}) : ${foundList}</div>`;
  } else {
    banner.className = "banner info";
    banner.innerHTML = `✓ ${data.statuses_found.length} statut(s) détecté(s), tous déjà mappés : ${foundList}`;
  }
  banner.style.display = "block";
}

document.getElementById("pImportXlsx").addEventListener("change", async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const fd = new FormData();
  fd.append("file", file);
  document.getElementById("pImportHint").textContent = "Import en cours...";
  document.getElementById("pImportHint").classList.add("show");
  showBusy("Import du fichier PCHC en cours...");
  let res, data;
  try{
    res = await fetch("/api/pchc/import", {method:"POST", body: fd});
    data = await res.json();
  } finally {
    hideBusy();
  }
  if(data.ok){
    toast("Import réussi");
    document.getElementById("pImportHint").textContent = "✓ " + Object.entries(data.result).map(([k,v])=>`${PCHC_CAT_LABELS[k]||k}: ${v}`).join(" · ");
    showDetectedStatuses(data);
    renderPImport();
  } else {
    toast(data.error || "Erreur d'import");
    document.getElementById("pImportHint").textContent = "";
  }
  e.target.value = "";
});

document.getElementById("pCsvImportBtn").addEventListener("click", async ()=>{
  const file = document.getElementById("pCsvFile").files[0];
  const kind = document.getElementById("pCsvKind").value;
  if(!file){ toast("Choisissez un fichier CSV"); return; }
  const fd = new FormData();
  fd.append("file", file);
  fd.append("kind", kind);
  const res = await fetch("/api/pchc/import", {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){ toast("Import CSV réussi"); showDetectedStatuses(data); renderPImport(); }
  else toast(data.error || "Erreur d'import");
});

async function renderPStatusColorsTable(){
  const res = await fetch("/api/pchc/status-colors");
  const data = await res.json();
  const palette = data.palette;
  const tbody = document.querySelector("#pStatusColorsTable tbody");
  tbody.innerHTML = Object.entries(data.mapping).sort().map(([statut,color])=>`
    <tr>
      <td>${esc(statut)}</td>
      <td>
        <select class="status-color-select" data-statut="${esc(statut)}">
          ${Object.keys(palette).map(c=>`<option value="${c}" ${c===color?'selected':''} style="color:${palette[c]};">${c}</option>`).join("")}
        </select>
      </td>
    </tr>
  `).join("") || `<tr><td colspan="2" style="color:var(--muted);">Importez des données pour voir les statuts observés</td></tr>`;
}

document.getElementById("pSaveStatusColorsBtn").addEventListener("click", async ()=>{
  const mapping = {};
  document.querySelectorAll(".status-color-select").forEach(sel=>{ mapping[sel.dataset.statut] = sel.value; });
  await fetch("/api/pchc/status-colors", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({mapping})});
  toast("Correspondance enregistrée");
});

// ---------- P-ANALYSIS ----------
async function renderPAnalysis(){
  const periodLabel = `${pState.start||''} — ${pState.end||''}`;
  document.getElementById("pPromptPeriodLabel").textContent = periodLabel;
  const d = pState.data || (await (await fetch(`/api/pchc/data?start=${pState.start||'2000-01-01'}&end=${pState.end||new Date().toISOString().slice(0,10)}`)).json());
  const s = d.summary;
  const catLines = Object.entries(d.categories).map(([k,c])=>`- ${c.label} : ${c.total} dossiers, ${c.delivered} délivrés (${c.taux}%), ${c.encours} en cours, ${c.bloque} bloqués/refusés`).join("\n");
  const top10Lines = d.top10.total.slice(0,10).map(t=>`- ${t.entity} : ${t.count} dossiers`).join("\n");

  document.getElementById("pPrompt").textContent =
`Tu es un analyste réglementaire. Rédige une synthèse exécutive professionnelle pour le
reporting métier PCHC (AMMPS), à partir des données suivantes :

PÉRIODE : du ${pState.start} au ${pState.end}

SYNTHÈSE GLOBALE
- Total dossiers : ${s.total}
- Délivrés/Validés : ${s.delivered}
- En cours : ${s.encours}
- Bloqués/Refusés : ${s.bloque}
- Taux de délivrance global : ${s.taux}%
- Dossiers en cours depuis plus de 90 jours : ${d.backlog['>90']}

DÉTAIL PAR MODULE
${catLines}

TOP 10 OPÉRATEURS (volume de dossiers)
${top10Lines || '- (aucune donnée)'}

ALERTES DÉTECTÉES
${d.alerts.length ? d.alerts.map(a=>'- '+a).join('\n') : '- Aucune'}

CONSIGNES DE RÉDACTION :
1. Une introduction (1 paragraphe) présentant la période et le volume global.
2. Une analyse du taux de délivrance par module, en identifiant les modules en difficulté.
3. Une analyse du backlog et de son ancienneté — points d'attention prioritaires.
4. Une lecture des opérateurs les plus actifs.
5. Une conclusion synthétique orientée Direction (2-3 phrases), avec recommandations.

Format : texte uniquement (pas de tableau, pas de puces), un paragraphe par point, prêt à
être collé dans l'application. Ton sobre, factuel, orienté pilotage.`;

  const notesRes = await fetch("/api/pchc/notes");
  const notes = await notesRes.json();
  document.getElementById("pCopilotText").value = notes.copilot_text || "";
  pState.notesText = notes.copilot_text || "";
}

async function savePchcNotes(silent){
  const text = document.getElementById("pCopilotText").value;
  pState.notesText = text;
  await fetch("/api/pchc/notes", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({copilot_text: text})});
  if(silent) pulseAutosave("pAnalysisAutosaveHint");
}
document.getElementById("pCopilotText").addEventListener("input", debounce(()=>savePchcNotes(true), 1800));
document.getElementById("pSaveNotesBtn").addEventListener("click", async ()=>{
  await savePchcNotes(false);
  saveIndicatorPulse();
  toast("Analyse enregistrée");
});
document.querySelectorAll('.copy-btn[data-copy="pPrompt"]').forEach(btn=>{
  btn.addEventListener("click", ()=>{
    navigator.clipboard.writeText(document.getElementById("pPrompt").textContent).then(()=>toast("Prompt copié"));
  });
});

document.getElementById("pSaveThresholdsBtn").addEventListener("click", async ()=>{
  const payload = {
    pchc_thresholds: {
      taux_haut: document.getElementById("pThHaut").value === "" ? null : Number(document.getElementById("pThHaut").value),
      taux_bas: document.getElementById("pThBas").value === "" ? null : Number(document.getElementById("pThBas").value),
      backlog_alert: document.getElementById("pThBacklog").value === "" ? null : Number(document.getElementById("pThBacklog").value),
      old_dossiers_alert: document.getElementById("pThOld").value === "" ? null : Number(document.getElementById("pThOld").value),
    }
  };
  await fetch("/api/pchc/settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(payload)});
  toast("Seuils enregistrés");
});

// ============================================================
// INIT
// ============================================================
(async function init(){
  await loadGlobalSettings();
  const res = await fetch("/api/auth/me");
  const data = await res.json();
  if(data.setup_required){ showAuthOverlay(true); return; }
  if(!data.authenticated){ showAuthOverlay(false); return; }
  currentUser = data.user;
  hideAuthOverlay();
  await startApp();
})();

// ---------- Drag & drop générique sur les zones d'import (Lot 5+) ----------
// Toute zone .file-btn (bouton stylé enrobant un <input type=file hidden>) devient une zone
// de dépôt : on simule la sélection de fichier puis on redéclenche l'événement 'change' déjà
// géré par les handlers existants — aucune duplication de logique d'import nécessaire.
document.querySelectorAll(".file-btn").forEach(label=>{
  const input = label.querySelector('input[type="file"]');
  if(!input) return;
  const zone = label.closest(".card") || label;
  zone.classList.add("dropzone-target");
  ["dragenter", "dragover"].forEach(evt=>{
    zone.addEventListener(evt, (e)=>{ e.preventDefault(); e.stopPropagation(); zone.classList.add("drag-over"); });
  });
  ["dragleave", "drop"].forEach(evt=>{
    zone.addEventListener(evt, (e)=>{ e.preventDefault(); e.stopPropagation(); zone.classList.remove("drag-over"); });
  });
  zone.addEventListener("drop", (e)=>{
    if(e.dataTransfer.files && e.dataTransfer.files.length){
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event("change", {bubbles:true}));
    }
  });
});

// ---------- Raccourcis clavier (Lot 5+) ----------
document.addEventListener("keydown", (e)=>{
  const tag = (e.target.tagName || "").toLowerCase();
  const typing = tag === "input" || tag === "textarea" || e.target.isContentEditable;

  // "/" focus la recherche globale (sauf si déjà en train de taper ailleurs)
  if(e.key === "/" && !typing){
    const search = document.getElementById("globalSearch");
    if(search){ e.preventDefault(); search.focus(); }
    return;
  }
  // Échap ferme les modales ouvertes (aperçu email/PDF, onboarding)
  if(e.key === "Escape"){
    ["previewModalOverlay", "pdfPreviewModalOverlay", "onboardingOverlay"].forEach(id=>{
      const el = document.getElementById(id);
      if(el && el.style.display !== "none") el.style.display = "none";
    });
    document.getElementById("globalSearchResults")?.classList.remove("show");
    return;
  }
  // Ctrl/Cmd+S sur un dashboard déclenche l'export PDF du volet actif (au lieu de la boîte
  // de dialogue "Enregistrer la page" du navigateur, peu pertinente ici)
  if((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s" && !typing){
    e.preventDefault();
    const activeTab = document.querySelector(".nav-item.active")?.dataset.tab;
    if(activeTab === "dashboard") document.getElementById("pdfBtn")?.click();
    else if(activeTab === "m-dashboard") document.getElementById("mPdfBtn")?.click();
    else if(activeTab === "p-dashboard") document.getElementById("pPdfBtn")?.click();
  }
});

// ============================================================
// PARTAGE SANAD SUR LE RÉSEAU LOCAL
// ============================================================
document.getElementById("shareSanadBtn")?.addEventListener("click", async ()=>{
  document.getElementById("shareSanadOverlay").style.display = "flex";
  document.getElementById("shareRestartHint").style.display = "none";
  try{
    const res = await fetch("/api/system/network-info");
    const info = await res.json();
    document.getElementById("shareModeLocal").checked = !info.network_enabled;
    document.getElementById("shareModeNetwork").checked = info.network_enabled;
    document.getElementById("sharePortInput").value = info.configured_port;
    updateShareLinkBox(info.network_enabled, info.share_url || `http://${info.lan_ip}:${info.configured_port}`);
  } catch(e){
    toast("Impossible de récupérer les informations réseau");
  }
});

document.getElementById("closeShareSanadBtn")?.addEventListener("click", ()=>{
  document.getElementById("shareSanadOverlay").style.display = "none";
});

function updateShareLinkBox(enabled, url){
  const box = document.getElementById("shareLinkBox");
  const input = document.getElementById("shareLinkValue");
  box.style.display = enabled ? "" : "none";
  if(enabled) input.value = url;
}

document.querySelectorAll('input[name="shareMode"]').forEach(r=>{
  r.addEventListener("change", async ()=>{
    const enabled = document.getElementById("shareModeNetwork").checked;
    if(enabled){
      try{
        const res = await fetch("/api/system/network-info");
        const info = await res.json();
        const port = document.getElementById("sharePortInput").value || info.configured_port;
        updateShareLinkBox(true, `http://${info.lan_ip}:${port}`);
      } catch(e){ /* silencieux */ }
    } else {
      updateShareLinkBox(false, "");
    }
  });
});

document.getElementById("sharePortInput")?.addEventListener("input", async ()=>{
  if(!document.getElementById("shareModeNetwork").checked) return;
  try{
    const res = await fetch("/api/system/network-info");
    const info = await res.json();
    const port = document.getElementById("sharePortInput").value || info.configured_port;
    updateShareLinkBox(true, `http://${info.lan_ip}:${port}`);
  } catch(e){ /* silencieux */ }
});

document.getElementById("copyShareLinkBtn")?.addEventListener("click", ()=>{
  const input = document.getElementById("shareLinkValue");
  input.select();
  navigator.clipboard?.writeText(input.value).then(()=> toast("Lien copié")).catch(()=>{
    document.execCommand("copy"); toast("Lien copié");
  });
});

document.getElementById("saveShareSanadBtn")?.addEventListener("click", async ()=>{
  const host = document.getElementById("shareModeNetwork").checked ? "0.0.0.0" : "127.0.0.1";
  const port = parseInt(document.getElementById("sharePortInput").value, 10) || 5050;
  const btn = document.getElementById("saveShareSanadBtn");
  const hint = document.getElementById("shareRestartHint");
  btn.disabled = true;
  try{
    const res = await fetch("/api/system/network-config", {
      method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({host, port}),
    });
    const data = await res.json();
    if(!data.ok){
      toast(data.error || "Erreur lors de l'enregistrement");
      btn.disabled = false;
      return;
    }
    if(data.share_url) updateShareLinkBox(true, data.share_url); else updateShareLinkBox(false, "");
    if(data.restart_scheduled){
      hint.style.display = "";
      hint.textContent = "🔄 Redémarrage automatique en cours… (quelques secondes)";
      await waitForServerRestart(host === "0.0.0.0" ? window.location.hostname : "127.0.0.1", port);
      hint.textContent = "✅ Redémarré — nouvelle page en cours de chargement…";
      window.location.href = `http://${host === "0.0.0.0" ? window.location.hostname : "127.0.0.1"}:${port}/`;
    } else {
      hint.style.display = "";
      hint.textContent = "✅ Appliqué immédiatement — aucun redémarrage nécessaire.";
      toast("Paramètres réseau mis à jour");
      btn.disabled = false;
    }
  } catch(e){
    toast("Erreur réseau lors de l'enregistrement");
    btn.disabled = false;
  }
});

// Sonde le nouveau port jusqu'à ce que le serveur (redémarré) réponde à nouveau, pour ne
// rediriger le navigateur qu'une fois l'application vraiment prête (évite un écran d'erreur
// de connexion pendant les ~1-2 secondes de redémarrage).
async function waitForServerRestart(hostname, port, maxWaitMs=20000){
  const start = Date.now();
  await new Promise(r=>setTimeout(r, 1500));  // laisse le temps au process de quitter avant de sonder
  while(Date.now() - start < maxWaitMs){
    try{
      const res = await fetch(`http://${hostname}:${port}/`, {mode: "no-cors"});
      return true;
    } catch(e){
      await new Promise(r=>setTimeout(r, 800));
    }
  }
  return false;
}
