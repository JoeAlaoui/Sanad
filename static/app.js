// ============================================================
// Helpdesk Dashboard - logique front-end (multi-volets)
// ============================================================
const JOURS = ["Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"];
const MOIS = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];

function esc(s){
  if(s === null || s === undefined) return "";
  return String(s).replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}

if(window.Chart && window.ChartDataLabels){
  Chart.register(ChartDataLabels);
  Chart.defaults.set('plugins.datalabels', {display:false});
}

let currentModule = "tarkhiss";
let currentUser = null;
let modulesInfo = {};
let state = { ym: null, calls: {}, emails: {}, analysis: {} };
let prevState = { calls: {}, emails: {}, analysis: {} };
let globalSettings = { agency_name: "", logo_filename: null };
let moduleSettings = { contacts: [], greeting: "Bonjour,", signature_name: "", signature_function: "", signature_phone: "" };
let charts = {};
let annualChart = null, dailyLoadChart = null, weeklyEmailsChart = null, channelChart = null;

// Moussanada
let mState = { ym: null, month: null, timeseries: {}, notes: {} };
let mCharts = {};
let mAggMode = "racine"; // "racine" | "leaf"
let mAnnualChart = null, mBacklogChart = null;

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

function saveIndicatorPulse(){
  const el = document.getElementById("saveIndicator");
  el.textContent = "✓ Enregistré";
  el.classList.add("show");
  setTimeout(()=>el.classList.remove("show"), 1800);
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
async function loadModules(){
  const res = await fetch("/api/modules");
  modulesInfo = await res.json();
  const allowedModules = (currentUser?.modules) || Object.keys(modulesInfo);
  const sw = document.getElementById("moduleSwitch");
  sw.innerHTML = Object.entries(modulesInfo).filter(([key])=>allowedModules.includes(key)).map(([key,info])=>`
    <button data-module="${key}" class="${key===currentModule?'active':''}">${info.label}</button>
  `).join("");
  sw.querySelectorAll("button").forEach(b=>{
    b.addEventListener("click", ()=>switchModule(b.dataset.module));
  });
  if(!allowedModules.includes(currentModule)) currentModule = allowedModules[0] || "tarkhiss";
  applyModuleBrand();
}

function defaultTabForRole(){
  if(!currentUser) return "dashboard";
  if(currentUser.role === "hotliner") return "knowledge";
  return currentModule === "moussanada" ? "m-dashboard" : "dashboard";
}

function applyModuleBrand(){
  const info = modulesInfo[currentModule] || {label: currentModule, subtitle:"", ready:true};
  document.getElementById("brandMark").textContent = info.label[0];
  document.getElementById("brandTitle").textContent = info.label;
  document.getElementById("brandSub").textContent = info.subtitle;
  document.querySelectorAll(".module-switch button").forEach(b=>{
    b.classList.toggle("active", b.dataset.module === currentModule);
  });
  ["adminModuleLabel","adminModuleLabel2","adminModuleLabel3","adminModuleLabel4","adminModuleLabel5","adminModuleLabel6","adminModuleLabel7","adminModuleLabel8"].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.textContent = info.label;
  });
  document.getElementById("moduleNotReadyBanner").style.display = info.ready ? "none" : "block";
  document.getElementById("monthPickerWrap").style.display = (currentUser && currentUser.role === "hotliner") ? "none" : "";

  // Filtre les items de nav selon le volet actif ET le rôle de l'utilisateur
  document.querySelectorAll(".nav-item").forEach(btn=>{
    const m = btn.dataset.module;
    const modOk = (m === "shared" || m === currentModule);
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
  document.querySelectorAll(".nav-item").forEach(b=>b.classList.toggle("active", b.dataset.tab === tab));
  document.querySelectorAll(".tab-panel").forEach(p=>p.classList.remove("active"));
  document.getElementById("tab-"+tab).classList.add("active");
  renderTab(tab);
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
    } else {
      await loadMoussanadaMonth(mState.ym || defaultMonth());
    }
  }
  const activeTab = document.querySelector(".nav-item.active")?.dataset.tab;
  if(activeTab) renderTab(activeTab);
}

// ---------- Navigation ----------
document.querySelectorAll(".nav-item").forEach(btn=>{
  btn.addEventListener("click", ()=> activateTab(btn.dataset.tab));
});

// ---------- Month handling ----------
const monthInput = document.getElementById("monthSelect");
monthInput.addEventListener("change", ()=>{
  if(currentModule === "tarkhiss") loadMonth(monthInput.value);
  else loadMoussanadaMonth(monthInput.value);
});

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

  renderCallsTable();
  renderEmailsTable();
  renderAnalysisForms();
  checkReminder();

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
  if("Notification" in window){
    if(Notification.permission === "granted"){
      new Notification(`Rapport ${moduleLabel} en attente`, {body:`Le rapport de ${monthLabel(py)} n'a pas encore été envoyé.`});
    } else if(Notification.permission !== "denied"){
      Notification.requestPermission();
    }
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
      <div style="font-size:11px;color:#67737E;">${globalSettings.agency_name || ""} · ${info.label} · ${monthLabel(state.ym)}</div>
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
  const palette = ["#0B4965","#DC2828","#F59F0A","#25935F","#67737E","#8E5AA6","#1794CF"];
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
    ctx.fillStyle = "#67737E";
    ctx.font = "11px Arial";
    ctx.fillText(k.label, x+12, cy+22);
    ctx.fillStyle = k.color || "#0B4965";
    ctx.font = "bold 26px Arial";
    ctx.fillText(String(k.value), x+12, cy+56);
    if(k.delta){
      ctx.fillStyle = k.deltaColor || "#67737E";
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

  ctx.fillStyle = "#67737E";
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
document.getElementById("pdfBtn").addEventListener("click", ()=> window.print());
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
async function loadGlobalSettings(){
  const res = await fetch("/api/global-settings");
  globalSettings = await res.json();
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
  await fetch("/api/global-settings", {method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({agency_name: globalSettings.agency_name})});
  toast("Paramètres généraux enregistrés");
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

async function loadAdmin(){
  await loadGlobalSettings();
  await loadModuleSettings();
  document.getElementById("agencyNameInput").value = globalSettings.agency_name || "";
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
  renderContactsAdmin();

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
    </tr>
  `).join("") || `<tr><td colspan="5" style="color:var(--muted);">Aucun envoi enregistré</td></tr>`;

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
      <div style="font-size:11px;color:#67737E;">${globalSettings.agency_name || ""} · ${info.label} · ${monthLabel(mState.ym)}</div>
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
}

document.getElementById("mPdfBtn").addEventListener("click", ()=> window.print());
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
  const res = await fetch(`/api/moussanada/import/${mState.ym}`, {method:"POST", body: fd});
  const data = await res.json();
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
  const kind = document.getElementById("mCsvKind").value;
  if(!file){ toast("Choisissez un fichier CSV"); return; }
  const fd = new FormData();
  fd.append("file", file);
  fd.append("kind", kind);
  const res = await fetch(`/api/moussanada/import/${mState.ym}`, {method:"POST", body: fd});
  const data = await res.json();
  if(data.ok){
    toast("Import CSV réussi");
    await loadMoussanadaMonth(mState.ym);
    renderMSources();
  } else {
    toast(data.error || "Erreur d'import");
  }
});

document.getElementById("mCsvCatBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/categories/${mState.ym}`;
});
document.getElementById("mCsvSrvBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/services/${mState.ym}`;
});
document.getElementById("mCsvTechBtn").addEventListener("click", ()=>{
  window.location.href = `/api/moussanada/export-csv-source/techniciens/${mState.ym}`;
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
  addPngExportButtons({mAnnualChart, mBacklogChart});
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
    if(currentModule === "moussanada"){
      await loadMoussanadaMonth(defaultMonth());
    } else {
      await loadMonth(defaultMonth());
    }
  }
  activateTab(defaultTabForRole());
}

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
// INIT
// ============================================================
(async function init(){
  const res = await fetch("/api/auth/me");
  const data = await res.json();
  if(data.setup_required){ showAuthOverlay(true); return; }
  if(!data.authenticated){ showAuthOverlay(false); return; }
  currentUser = data.user;
  hideAuthOverlay();
  await startApp();
})();
