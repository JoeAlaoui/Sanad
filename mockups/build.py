"""Génère les maquettes statiques du dashboard Tarkhiss dans plusieurs styles (python build.py)."""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))

BASE_CSS = """
*{box-sizing:border-box;margin:0}body{font-family:var(--font);background:var(--bg);color:var(--ink);font-size:14px;line-height:1.5}
.app{display:grid;grid-template-columns:248px 1fr;min-height:100vh}
.side{background:var(--side-bg);border-right:1px solid var(--side-line);padding:18px 14px;display:flex;flex-direction:column;gap:4px;position:sticky;top:0;height:100vh;overflow:auto;color:var(--side-ink)}
.brand{display:flex;align-items:center;gap:10px;padding:4px 8px 14px;font-weight:700;font-size:16px}
.logo{width:32px;height:32px;display:grid;place-items:center;background:var(--accent);color:var(--on-accent);font-weight:800;border-radius:var(--r-sm)}
.mod{display:flex;background:var(--mod-bg);border-radius:var(--r-sm);padding:3px;margin:0 4px 12px}.mod b{flex:1;text-align:center;padding:6px;border-radius:calc(var(--r-sm) - 2px);font-size:12px;font-weight:600;color:var(--muted)}.mod b.on{background:var(--mod-on);color:var(--ink);box-shadow:var(--sh-sm)}
.grp{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--side-muted);padding:12px 10px 4px;font-weight:600}
.nav{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:var(--r-sm);color:var(--side-ink);font-weight:500;font-size:13.5px}.nav:hover{background:var(--nav-hover)}.nav.on{background:var(--nav-on);color:var(--nav-on-ink);font-weight:600}
.nav em{margin-left:auto;font-style:normal;font-size:11px;padding:1px 7px;border-radius:99px;background:var(--warn-soft);color:var(--warn)}
.main{min-width:0}.top{display:flex;align-items:center;gap:12px;padding:12px 28px;background:var(--top-bg);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:5;backdrop-filter:blur(10px)}
.crumb{color:var(--muted);font-size:13px}.crumb b{color:var(--ink)}.sp{flex:1}
.pick{display:flex;align-items:center;gap:4px;border:1px solid var(--line);background:var(--card);border-radius:var(--r-sm);padding:3px;box-shadow:var(--sh-sm)}.pick span{padding:3px 12px;font-weight:600;font-size:13px}.pick i{font-style:normal;width:26px;height:26px;display:grid;place-items:center;border-radius:calc(var(--r-sm) - 2px);color:var(--muted)}
.search{border:1px solid var(--line);background:var(--card);border-radius:var(--r-sm);padding:7px 12px;color:var(--muted);width:240px;font-size:13px}
.pill{font-size:12px;padding:3px 10px;border-radius:99px;background:var(--accent-soft);color:var(--accent-ink);font-weight:600}
.avatar{width:32px;height:32px;border-radius:99px;background:var(--accent);color:var(--on-accent);display:grid;place-items:center;font-weight:700;font-size:12px}
.page{padding:26px 28px 48px;max-width:1500px}
.head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:20px;flex-wrap:wrap}.head h1{font-size:var(--h1);letter-spacing:-.02em;font-weight:700}.head p{color:var(--muted);margin-top:2px}
.acts{display:flex;gap:8px;flex-wrap:wrap}
.btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:var(--r-sm);border:1px solid var(--line);background:var(--card);font-weight:600;font-size:13px;color:var(--ink);box-shadow:var(--sh-sm)}.btn.pri{background:var(--accent);border-color:var(--accent);color:var(--on-accent)}
.alert{background:var(--warn-soft);color:var(--warn);border:1px solid var(--warn-line);border-radius:var(--r);padding:10px 14px;margin-bottom:16px;font-weight:500}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh);padding:18px}
.kpi .l{color:var(--muted);font-size:12.5px;font-weight:500}.kpi .v{font-size:var(--kv);font-weight:700;letter-spacing:-.02em;margin:4px 0 6px;font-variant-numeric:tabular-nums}
.kpi.red{--k:var(--bad)}.kpi.green{--k:var(--good)}.kpi.amber{--k:var(--warn)}.kpi[class*=" "]{border-left:3px solid var(--k)}
.tr{display:inline-flex;align-items:center;gap:4px;font-size:12px;font-weight:600;padding:1px 8px;border-radius:99px}.tr.up{background:var(--good-soft);color:var(--good)}.tr.dn{background:var(--bad-soft);color:var(--bad)}.tr.fl{background:var(--mod-bg);color:var(--muted)}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:14px 0}
.card h3{font-size:15px;font-weight:600;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center}.card h3 small{font-weight:500;color:var(--muted);font-size:12px}
.yoy{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.yoy div{border:1px solid var(--line);border-radius:var(--r-sm);padding:12px}.yoy small{color:var(--muted);display:block}.yoy b{font-size:20px}
table{width:100%;border-collapse:collapse}th{text-align:left;font-size:12px;color:var(--muted);font-weight:600;padding:8px 10px;border-bottom:1px solid var(--line);background:var(--th-bg)}td{padding:11px 10px;border-bottom:1px solid var(--line)}tr:last-child td{border-bottom:0}
.bar{display:inline-block;width:90px;height:6px;border-radius:9px;background:var(--mod-bg);vertical-align:middle;margin-right:8px;overflow:hidden}.bar i{display:block;height:100%;background:var(--bad)}.bar.g i{background:var(--good)}
.badge{display:inline-flex;align-items:center;gap:6px;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:600;background:var(--good-soft);color:var(--good)}.badge:before{content:"";width:6px;height:6px;border-radius:9px;background:currentColor}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}ul.l{padding-left:18px;color:var(--ink)}
.cv{position:relative;height:250px}
@media(max-width:1000px){.app{grid-template-columns:1fr}.side{display:none}.kpis,.yoy{grid-template-columns:1fr 1fr}.row2,.cols{grid-template-columns:1fr}}
"""

THEMES = {
 "heroui": dict(name="HeroUI", tag="Bleu #006FEE · cartes très arrondies · ombres douces · onglets « pilule »", css="""
:root{--font:Inter,system-ui,sans-serif;--bg:#f4f4f5;--card:#fff;--ink:#11181c;--muted:#71717a;--line:#e4e4e7;--accent:#006fee;--on-accent:#fff;--accent-soft:#cce3fd;--accent-ink:#004493;
--side-bg:#fff;--side-line:#e4e4e7;--side-ink:#3f3f46;--side-muted:#a1a1aa;--nav-hover:#f4f4f5;--nav-on:#cce3fd;--nav-on-ink:#004493;--mod-bg:#f4f4f5;--mod-on:#fff;--top-bg:rgba(244,244,245,.8);--th-bg:#fafafa;
--good:#17c964;--good-soft:#d4f6e3;--bad:#f31260;--bad-soft:#fdd0df;--warn:#c4841d;--warn-soft:#fef3dc;--warn-line:#fbe2b0;--r:20px;--r-sm:12px;--h1:26px;--kv:30px;
--sh:0 2px 12px rgba(0,0,0,.05);--sh-sm:0 1px 2px rgba(0,0,0,.05);--c1:#006fee;--c2:#f31260;--c3:#17c964;--c4:#f5a524}
.btn{border-radius:12px;border-color:transparent;background:#e4e4e7;box-shadow:none}.btn.pri{background:#006fee}.nav{border-radius:12px}.card{border:0}
.kpi.red,.kpi.green,.kpi.amber{border-left:0;background:linear-gradient(180deg,color-mix(in srgb,var(--k) 9%,#fff),#fff 70%)}"""),
 "untitled": dict(name="Untitled UI", tag="Violet #7F56D9 · bordures fines · badges à pastille · très sobre", css="""
:root{--font:Inter,system-ui,sans-serif;--bg:#fcfcfd;--card:#fff;--ink:#101828;--muted:#667085;--line:#eaecf0;--accent:#7f56d9;--on-accent:#fff;--accent-soft:#f4ebff;--accent-ink:#6941c6;
--side-bg:#fff;--side-line:#eaecf0;--side-ink:#344054;--side-muted:#98a2b3;--nav-hover:#f9fafb;--nav-on:#f9f5ff;--nav-on-ink:#6941c6;--mod-bg:#f2f4f7;--mod-on:#fff;--top-bg:#fff;--th-bg:#f9fafb;
--good:#067647;--good-soft:#ecfdf3;--bad:#b42318;--bad-soft:#fef3f2;--warn:#b54708;--warn-soft:#fffaeb;--warn-line:#fedf89;--r:12px;--r-sm:8px;--h1:24px;--kv:32px;
--sh:0 1px 3px rgba(16,24,40,.1),0 1px 2px rgba(16,24,40,.06);--sh-sm:0 1px 2px rgba(16,24,40,.05);--c1:#7f56d9;--c2:#f04438;--c3:#12b76a;--c4:#f79009}
.btn{font-weight:600}.kpi .l{font-weight:600;color:#475467}.tr{border-radius:16px}.kpi.red,.kpi.green,.kpi.amber{border-left:1px solid var(--line)}"""),
 "dark": dict(name="Linear / Geist (sombre)", tag="Noir profond · accent émeraude · densité élevée · mode « pro »", css="""
:root{--font:Inter,system-ui,sans-serif;--bg:#0a0a0b;--card:#111113;--ink:#ededef;--muted:#8b8b94;--line:#232326;--accent:#3ecf8e;--on-accent:#04130c;--accent-soft:#0f2a1e;--accent-ink:#3ecf8e;
--side-bg:#0d0d0f;--side-line:#232326;--side-ink:#b4b4bb;--side-muted:#66666e;--nav-hover:#17171a;--nav-on:#1c1c20;--nav-on-ink:#fff;--mod-bg:#17171a;--mod-on:#232328;--top-bg:rgba(10,10,11,.8);--th-bg:#0d0d0f;
--good:#3ecf8e;--good-soft:#0f2a1e;--bad:#ff6369;--bad-soft:#2d1315;--warn:#f5a524;--warn-soft:#2a2008;--warn-line:#4a3912;--r:10px;--r-sm:7px;--h1:22px;--kv:28px;
--sh:none;--sh-sm:none;--c1:#3ecf8e;--c2:#ff6369;--c3:#6e8bff;--c4:#f5a524}
.btn{background:#17171a;border-color:#2a2a2e}.btn.pri{background:#3ecf8e;border-color:#3ecf8e}.card{background:linear-gradient(180deg,#131316,#0f0f11)}"""),
 "fluent": dict(name="Fluent 2 / Microsoft 365", tag="Bleu #0F6CBD · angles 4 px · familier pour un environnement M365", css="""
:root{--font:'Segoe UI',system-ui,sans-serif;--bg:#f5f5f5;--card:#fff;--ink:#242424;--muted:#616161;--line:#e0e0e0;--accent:#0f6cbd;--on-accent:#fff;--accent-soft:#e6f1fb;--accent-ink:#0f548c;
--side-bg:#fafafa;--side-line:#e0e0e0;--side-ink:#242424;--side-muted:#707070;--nav-hover:#ededed;--nav-on:#e6f1fb;--nav-on-ink:#0f548c;--mod-bg:#ebebeb;--mod-on:#fff;--top-bg:#fff;--th-bg:#fafafa;
--good:#107c10;--good-soft:#dff6dd;--bad:#c50f1f;--bad-soft:#fdf3f4;--warn:#835b00;--warn-soft:#fff9f5;--warn-line:#f9e2ae;--r:4px;--r-sm:4px;--h1:24px;--kv:28px;
--sh:0 1.6px 3.6px rgba(0,0,0,.13),0 .3px .9px rgba(0,0,0,.1);--sh-sm:none;--c1:#0f6cbd;--c2:#c50f1f;--c3:#107c10;--c4:#c19c00}
.nav.on{border-left:3px solid var(--accent);border-radius:2px}.btn{font-weight:600}.top{box-shadow:0 1px 0 var(--line)}"""),
 "pastel": dict(name="Pastel (votre kit « Pastel »)", tag="Violet #8B6CFF · fond lavande · rayons généreux · doux", css="""
:root{--font:'Plus Jakarta Sans',Inter,system-ui,sans-serif;--bg:#f6f5fb;--card:#fff;--ink:#201c33;--muted:#6b6580;--line:#eeecf7;--accent:#8b6cff;--on-accent:#fff;--accent-soft:#e8e0fe;--accent-ink:#5a3fd1;
--side-bg:#fff;--side-line:#eeecf7;--side-ink:#4a4560;--side-muted:#a79ccb;--nav-hover:#f6f5fb;--nav-on:#e8e0fe;--nav-on-ink:#5a3fd1;--mod-bg:#f1eefc;--mod-on:#fff;--top-bg:rgba(246,245,251,.85);--th-bg:#faf9fe;
--good:#2bb673;--good-soft:#dcf6ea;--bad:#f0627e;--bad-soft:#fde4ea;--warn:#c98a17;--warn-soft:#fff4dc;--warn-line:#f5e0ad;--r:22px;--r-sm:14px;--h1:26px;--kv:30px;
--sh:0 8px 30px rgba(139,108,255,.08);--sh-sm:0 2px 8px rgba(139,108,255,.08);--c1:#8b6cff;--c2:#f0627e;--c3:#2bb673;--c4:#ffb84d}
.card{border:1px solid #fff}.btn{border-color:#eeecf7}.kpi.red,.kpi.green,.kpi.amber{border-left:0}.kpi .v{color:#2c2650}"""),
 "ammps": dict(name="AMMPS Modern 2026 (votre kit)", tag="Vert profond + lime #B9F227 · boutons pilule · identité institutionnelle", css="""
:root{--font:Inter,system-ui,sans-serif;--bg:#f5f7f4;--card:#fff;--ink:#0f172a;--muted:#64748b;--line:#dfe6e1;--accent:#b9f227;--on-accent:#052e24;--accent-soft:#ecfdf5;--accent-ink:#08744b;
--side-bg:#052e24;--side-line:#052e24;--side-ink:#d1fae5;--side-muted:#6ee7b7;--nav-hover:rgba(255,255,255,.07);--nav-on:#b9f227;--nav-on-ink:#052e24;--mod-bg:rgba(255,255,255,.08);--mod-on:#fff;--top-bg:rgba(245,247,244,.9);--th-bg:#f5f7f4;
--good:#08744b;--good-soft:#d1fae5;--bad:#df2328;--bad-soft:#fee2e2;--warn:#9a6a07;--warn-soft:#fef3c7;--warn-line:#fde68a;--r:24px;--r-sm:14px;--h1:28px;--kv:32px;
--sh:0 6px 20px rgba(15,23,42,.06);--sh-sm:none;--c1:#064e3b;--c2:#df2328;--c3:#b9f227;--c4:#0d9488}
.side .mod b{color:#a7f3d0}.side .mod b.on{color:#052e24}.btn{border-radius:99px}.btn.pri{background:#052e24;border-color:#052e24;color:#fff}.nav{border-radius:99px}
.logo{border-radius:14px;background:#b9f227}.card{border-radius:26px}.kpi.red,.kpi.green,.kpi.amber{border-left:0}.head h1{font-weight:900;letter-spacing:-.04em}.avatar{background:#052e24;color:#b9f227}"""),
}

BODY = """<div class="app">
<aside class="side"><div class="brand"><div class="logo">S</div>SANAD</div>
<div class="mod"><b class="on">Tarkhiss</b><b>Moussanada</b></div>
<div class="grp">Reporting Hotline &amp; Emails</div>
<a class="nav on">▦ Dashboard</a><a class="nav">☎ Appels <em>10/22</em></a><a class="nav">✉ Emails <em>10/22</em></a><a class="nav">◇ Analyse mensuelle</a><a class="nav">📈 Vue annuelle</a><a class="nav">✦ Prompts Copilot</a>
<div class="grp">Reporting Métier</div><a class="nav">▦ Dashboard</a><a class="nav">📥 Import données</a><a class="nav">◇ Analyse &amp; Copilot</a>
<div class="grp">Administration</div><a class="nav">⚙ Administration</a><a class="nav">👤 Utilisateurs</a><a class="nav">💬 Notes</a><a class="nav">📚 Base de connaissances</a>
<div class="grp">Qualité &amp; conformité</div><a class="nav">🩺 Qualité des données</a><a class="nav">📑 SOP &amp; preuves d'audit</a><a class="nav">🗄 Sauvegardes</a></aside>
<div class="main"><div class="top"><span class="crumb">Tarkhiss › Reporting Hotline &amp; Emails › <b>Dashboard</b></span><span class="sp"></span>
<div class="pick"><i>‹</i><span>octobre 2026</span><i>›</i></div><div class="search">🔍 Recherche globale…</div><span class="pill">Synchro à l'instant</span><span class="pill">Administrateur</span><div class="avatar">YA</div></div>
<div class="page"><div class="head"><div><h1>Dashboard — Tarkhiss — octobre 2026</h1><p>Vue interactive, codes couleur automatiques.</p></div>
<div class="acts"><span class="btn">⬇ Excel</span><span class="btn">⬇ PDF</span><span class="btn">🖼 One-pager</span><span class="btn pri">✉ Envoyer par email</span></div></div>
<div class="alert">⚠️ <b>Seuils dépassés :</b> Variation de volume anormale : baisse de 55% (seuil : 40%)</div>
<div class="kpis">
<div class="card kpi"><div class="l">📧 Emails reçus</div><div class="v">457</div><span class="tr up">▲ +12% vs préc. (408)</span></div>
<div class="card kpi"><div class="l">📤 Emails envoyés</div><div class="v">364</div><span class="tr up">▲ +8% vs préc. (337)</span></div>
<div class="card kpi"><div class="l">📞 Total appels</div><div class="v">363</div><span class="tr dn">▼ −5% vs préc. (382)</span></div>
<div class="card kpi"><div class="l">⏱️ Temps comm.</div><div class="v">11:42:08</div><span class="tr fl">→ stable</span></div></div>
<div class="kpis">
<div class="card kpi red"><div class="l">🔴 Problèmes techniques</div><div class="v">13</div><span class="tr dn">▲ +30% vs préc. (10)</span></div>
<div class="card kpi green"><div class="l">🟢 Demandes d'info</div><div class="v">29</div><span class="tr fl">→ stable</span></div>
<div class="card kpi"><div class="l">Ratio Bugs / Demandes</div><div class="v">0,45</div><span class="tr fl">—</span></div>
<div class="card kpi amber"><div class="l">⚠️ Semaine critique</div><div class="v">S2</div><span class="tr dn">13 tickets</span></div></div>
<div class="card"><h3>📆 Comparatif octobre 2026 vs octobre 2025 (N-1) <small>admin / superviseur</small></h3>
<div class="yoy"><div><small>Emails reçus</small><b>457</b> <small style="display:inline">(402)</small><br><span class="tr up">▲ +14%</span></div><div><small>Appels</small><b>363</b> <small style="display:inline">(390)</small><br><span class="tr dn">▼ −7%</span></div><div><small>Problèmes tech.</small><b>13</b> <small style="display:inline">(9)</small><br><span class="tr dn">▲ +44%</span></div><div><small>Demandes info</small><b>29</b> <small style="display:inline">(31)</small><br><span class="tr fl">→ −6%</span></div></div></div>
<div class="row2"><div class="card"><h3>Évolution hebdomadaire — Bugs vs Demandes <small>⬇ PNG</small></h3><div class="cv"><canvas id="c1"></canvas></div></div>
<div class="card"><h3>Répartition des problèmes techniques <small>⬇ PNG</small></h3><div class="cv"><canvas id="c2"></canvas></div></div></div>
<div class="row2"><div class="card"><h3>Charge quotidienne (appels + emails reçus) <small>⬇ PNG</small></h3><div class="cv"><canvas id="c3"></canvas></div></div>
<div class="card"><h3>Emails reçus / envoyés par semaine <small>⬇ PNG</small></h3><div class="cv"><canvas id="c4"></canvas></div></div></div>
<div class="card"><h3>🔴 Problèmes techniques signalés (13)</h3><table><thead><tr><th>Type</th><th>Nb</th><th>%</th><th>Description</th><th>Statut</th></tr></thead><tbody>
<tr><td>Connexion / authentification</td><td><b>6</b></td><td><span class="bar"><i style="width:46%"></i></span>46,2%</td><td>Échecs de login portail</td><td><span class="badge">Suivi Karaz</span></td></tr>
<tr><td>Upload de pièces jointes</td><td><b>4</b></td><td><span class="bar"><i style="width:31%"></i></span>30,8%</td><td>Fichiers &gt; 5 Mo rejetés</td><td><span class="badge">En cours</span></td></tr>
<tr><td>Génération de certificat</td><td><b>3</b></td><td><span class="bar"><i style="width:23%"></i></span>23,1%</td><td>PDF non généré</td><td><span class="badge">Corrigé</span></td></tr></tbody></table></div>
<div class="cols"><div class="card"><h3>🔑 Constats clés</h3><ul class="l"><li>Pic de charge en semaine 2 (13 tickets).</li><li>Les échecs de connexion dominent (46 %).</li></ul></div>
<div class="card"><h3>✅ Recommandations</h3><ul class="l"><li>Escalader le correctif d'authentification auprès de Karaz.</li><li>Publier un guide « pièces jointes » pour les opérateurs.</li></ul></div></div>
</div></div></div>
<script src="chart.umd.js"></script>
<script>
const css=getComputedStyle(document.documentElement),v=n=>css.getPropertyValue(n).trim();
Chart.defaults.font.family=v('--font');Chart.defaults.color=v('--muted');Chart.defaults.borderColor=v('--line');
const base={responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8}}}};
new Chart(c1,{type:'bar',data:{labels:['S1','S2','S3','S4'],datasets:[{label:'Bugs',data:[4,9,3,2],backgroundColor:v('--c2'),borderRadius:6},{label:'Demandes',data:[7,11,6,5],backgroundColor:v('--c3'),borderRadius:6}]},options:{...base,scales:{x:{grid:{display:false}},y:{beginAtZero:true}}}});
new Chart(c2,{type:'doughnut',data:{labels:['Connexion','Upload','Certificat'],datasets:[{data:[6,4,3],backgroundColor:[v('--c1'),v('--c2'),v('--c4')],borderWidth:0}]},options:{...base,cutout:'68%'}});
const d=[...Array(22)].map((_,i)=>i+1),load=[30,42,38,25,0,0,33,47,40,36,29,0,0,22,35,41,38,27,0,0,31,28];
new Chart(c3,{type:'line',data:{labels:d,datasets:[{label:'Charge',data:load,borderColor:v('--c1'),backgroundColor:v('--c1')+'22',fill:true,tension:.35,pointRadius:0}]},options:{...base,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true}}}});
new Chart(c4,{type:'bar',data:{labels:['Sem. 1','Sem. 2','Sem. 3','Sem. 4'],datasets:[{label:'Reçus',data:[111,93,128,125],backgroundColor:v('--c1'),stack:'e',borderRadius:4},{label:'Envoyés',data:[80,70,110,104],backgroundColor:v('--c4'),stack:'e',borderRadius:4}]},options:{...base,scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,beginAtZero:true}}}});
</script>"""

for key, t in THEMES.items():
    html = f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>SANAD — maquette {t["name"]}</title><style>{BASE_CSS}{t["css"]}</style></head><body>{BODY}</body></html>'
    open(os.path.join(HERE, f"{key}.html"), "w", encoding="utf-8").write(html)

cards = "".join(f'<a class="c" href="{k}.html" target="_blank"><div class="f"><iframe src="{k}.html" loading="lazy" tabindex="-1"></iframe></div><b>{t["name"]}</b><span>{t["tag"]}</span></a>' for k, t in THEMES.items())
index = f'''<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>SANAD — comparatif d'interfaces</title><style>
body{{font-family:Inter,system-ui,sans-serif;background:#f4f4f5;margin:0;padding:28px;color:#18181b}}h1{{margin:0 0 4px}}p{{color:#52525b;margin:0 0 22px}}
.g{{display:grid;grid-template-columns:repeat(auto-fill,minmax(420px,1fr));gap:20px}}.c{{background:#fff;border-radius:16px;padding:14px;text-decoration:none;color:inherit;box-shadow:0 2px 10px rgba(0,0,0,.06);display:flex;flex-direction:column;gap:4px}}
.f{{height:270px;overflow:hidden;border-radius:10px;border:1px solid #e4e4e7;position:relative}}.f iframe{{width:1500px;height:950px;border:0;transform:scale(.265);transform-origin:0 0;pointer-events:none}}.c b{{margin-top:8px}}.c span{{color:#71717a;font-size:13px}}</style></head>
<body><h1>Comparatif d'interfaces — dashboard Tarkhiss</h1><p>Mêmes données, 6 habillages. Cliquez une carte pour l'ouvrir en grand (maquettes statiques, sans lien avec le code SANAD).</p><div class="g">{cards}</div></body></html>'''
open(os.path.join(HERE, "index.html"), "w", encoding="utf-8").write(index)
print("ok", list(THEMES))
