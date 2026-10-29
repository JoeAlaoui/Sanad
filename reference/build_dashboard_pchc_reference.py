import openpyxl, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.chart import BarChart, PieChart, DoughnutChart, Reference, Series
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.properties import PageSetupProperties

SRC, OUT = '/tmp/dash/in3.xlsx', '/tmp/dash/out3.xlsx'
wb = openpyxl.load_workbook(SRC)

NAVY, TEAL, GREY, LGREY, LINE = '1F3864', '0E7C86', '5B6B7F', 'F2F5F9', 'D0D7E2'
GREEN, AMBER, RED = 'C6EFCE', 'FFE699', 'FFCCCC'
CUR = dict(title='1F4E79', light='D6DCE4', sec='2F5496', secfont='FFFFFF', hdr='2F5496')
THEMES = {'KPI Identification': dict(title='2E75B6', light='D6E3F8', sec='4472C4', secfont='FFFFFF', hdr='8FAADC'),
 'KPI Déclaration': dict(title='538135', light='E2EFDA', sec='70AD47', secfont='FFFFFF', hdr='A9D08E'),
 'KPI Certificat Enreg.': dict(title='C65911', light='FCE4D6', sec='ED7D31', secfont='FFFFFF', hdr='F4B183'),
 'KPI Certificat LV': dict(title='7030A0', light='E2D1F2', sec='9966CC', secfont='FFFFFF', hdr='B4A7D6'),
 'KPI Autor. Import.': dict(title='1F4E79', light='D6DCE5', sec='2F5496', secfont='FFFFFF', hdr='8FAADC'),
 'Dashboard': dict(title='1F4E79', light='D6DCE4', sec='BDD7EE', secfont='1F4E79', hdr='2F5496')}
UPDATE = datetime.datetime(2026, 10, 28)
LAST = 50000  # plage de lecture des feuilles brutes (lignes)

def F(sz=10, b=False, c='000000', i=False): return Font(name='Calibri', size=sz, bold=b, color=c, italic=i)
def fill(c): return PatternFill('solid', start_color=c, end_color=c)
thin = Side(style='thin', color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_range(ws, rng, font=None, fl=None, al=None, border=None):
    for row in ws[rng]:
        for c in row:
            if font: c.font = font
            if fl: c.fill = fl
            if al: c.alignment = al
            if border: c.border = border

def put(ws, ref, v, font=None, fl=None, al=None, border=None, nf=None):
    c = ws[ref]; c.value = v
    if font: c.font = font
    if fl: c.fill = fl
    if al: c.alignment = al
    if border: c.border = border
    if nf: c.number_format = nf
    return c

def merge(ws, rng, v, font, fl=None, al=None, border=None, nf=None):
    a, b = rng.split(':')
    if a != b: ws.merge_cells(rng)
    style_range(ws, rng, font, fl, al, border)
    ws[a].value = v
    if nf: ws[a].number_format = nf

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
RIGHT = Alignment(horizontal='right', vertical='center')

def banner(ws, title, subtitle, last_col='L'):
    ws.sheet_view.showGridLines = False
    merge(ws, f'A1:{last_col}1', title, F(18, True, 'FFFFFF'), fill(CUR['title']), Alignment(horizontal='left', vertical='center', indent=1))
    ws.row_dimensions[1].height = 38
    merge(ws, 'A2:E2', subtitle, F(10, False, '000000', True), fill(CUR['light']), Alignment(horizontal='left', vertical='center', indent=1))
    style_range(ws, f'F2:{last_col}2', fl=fill(CUR['light']))
    put(ws, 'F2', 'Mise à jour :', F(10, True), fill(CUR['light']), RIGHT)
    merge(ws, 'G2:H2', UPDATE, F(10, True), fill(CUR['light']), Alignment(horizontal='left', vertical='center', indent=1), nf='dd/mm/yyyy')
    ws.row_dimensions[2].height = 20

def section(ws, rng, text):
    merge(ws, rng, text, F(12, True, CUR['secfont']), fill(CUR['sec']), Alignment(horizontal='left', vertical='center', indent=1))
    ws.row_dimensions[ws[rng.split(':')[0]].row].height = 22
    for row in ws[rng]:
        for c in row: c.border = Border(bottom=Side(style='medium', color=CUR['title']))

def header(ws, rng, labels=None):
    for row in ws[rng]:
        for c in row:
            c.font = F(10, True, '000000' if CUR['hdr']=='FFC000' else 'FFFFFF'); c.fill = fill(CUR['hdr']); c.alignment = CENTER; c.border = BOX
    if labels:
        for ref, t in labels.items(): ws[ref].value = t

CAT_FILL = {'Délivré': GREEN, 'En cours': 'E4EEF9', 'Non conforme': AMBER, 'Refusé': RED, 'Autre': 'ECEFF3'}
CAT_PT = {'Délivré': ['2E8B57'], 'En cours': ['1F6FB2', '4FA3D9', '8CC4EC', '0E4C7A', '5B7DB1', 'A9C4E0', '3B5998'],
          'Non conforme': ['F4B400'], 'Refusé': ['D64541', 'A93226'], 'Autre': ['9AA5B1']}

def rng(raw, col, start): return f"'{raw}'!${col}${start}:${col}${LAST}"

MODS = {
 'KPI Déclaration': dict(
    title="📋 KPI DÉCLARATION D'ACTIVITÉ", raw='déclarationActivité', idc='B', stc='H', start=6,
    statuses=[('En Vérification','En cours'),('En Evaluation','En cours'),('Evaluation Compléments','En cours'),
      ('Attente validation évaluation','En cours'),('Séance de travail','En cours'),('Validation acte','En cours'),
      ('Attente Signature','En cours'),('Non Conforme','Non conforme'),('Irrecevable','Refusé'),('Rejeté','Refusé'),('Délivré','Délivré')]),
 'KPI Certificat Enreg.': dict(
    title="📋 KPI CERTIFICAT D'ENREGISTREMENT (CE)", raw='certificatEnregistrement', idc='B', stc='J', start=6,
    statuses=[('En vérification','En cours'),('En Evaluation','En cours'),('Evaluation Compléments','En cours'),
      ('Attente validation évaluation','En cours'),('Edition acte','En cours',['Editon acte']),('Attente paiement','En cours'),
      ('Séance de travail','En cours'),('Validation acte','En cours'),('Attente Signature','En cours'),('Non conforme','Non conforme'),
      ('Irrecevable','Refusé'),('Rejeté','Refusé'),('Délivré','Délivré')]),
 'KPI Certificat LV': dict(
    title='📋 KPI CERTIFICAT LIBRE VENTE (CLV)', raw='certificatLibreVente', idc='B', stc='H', start=6,
    statuses=[('En vérification','En cours'),('En Evaluation','En cours'),('Attente validation évaluation','En cours'),
      ('Edition acte','En cours'),('Validation acte','En cours'),('Attente signature','En cours'),('Non conforme','Non conforme'),
      ('Irrecevable','Refusé'),('Rejeté','Refusé'),('Délivré','Délivré')]),
 'KPI Autor. Import.': dict(
    title="📋 KPI AUTORISATION D'IMPORTATION MATIÈRES PREMIÈRES", raw='autorisationImportationMatièreP', idc='B', stc='F', start=6,
    statuses=[('En vérification','En cours'),('En Evaluation','En cours'),('Evaluation Compléments','En cours'),
      ('Attente validation évaluation','En cours'),('Séance de travail','En cours',['Sénace de travail']),('Edition acte','En cours'),('Validation acte','En cours'),
      ('Signature acte','En cours'),('Non conforme','Non conforme'),('Irrecevable','Refusé'),('Rejeté','Refusé'),('Délivré','Délivré')]),
}
ROWMAP, INFO = {}, {}

def new_sheet(name, tab):
    idx = wb.sheetnames.index(name)
    wb.remove(wb[name])
    ws = wb.create_sheet(name, idx)
    ws.sheet_properties.tabColor = tab
    for col, w in zip('ABCDEFGHIJKL', [40, 12, 10, 17, 3, 12, 12, 12, 12, 12, 12, 12]):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return ws

def pt_colors(cats):
    cnt, out = {}, []
    for c in cats:
        i = cnt.get(c, 0); cnt[c] = i + 1
        out.append(CAT_PT[c][i % len(CAT_PT[c])])
    return out

def make_pie(ws, kind, title, cats_ref, data_ref, colors, anchor, w=15.5, h=8.2):
    ch = DoughnutChart() if kind == 'doughnut' else PieChart()
    ch.title = title; ch.style = 10
    ch.add_data(data_ref, titles_from_data=False); ch.set_categories(cats_ref)
    s = ch.series[0]
    for i, col in enumerate(colors):
        p = DataPoint(idx=i); p.graphicalProperties.solidFill = col
        p.graphicalProperties.line.solidFill = 'FFFFFF'; s.dPt.append(p)
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.showPercent = False
    s.dLbls.showCatName = False; s.dLbls.showSerName = False; s.dLbls.showLegendKey = False
    s.dLbls.numFmt = '#,##0;;'
    ch.legend.position = 'r'
    ch.width, ch.height = w, h
    ws.add_chart(ch, anchor)

def make_bar(ws, title, cats_ref, data_ref, colors, anchor, horizontal=True, w=15.5, h=9):
    ch = BarChart(); ch.type = 'bar' if horizontal else 'col'; ch.style = 10; ch.title = title
    ch.add_data(data_ref, titles_from_data=False); ch.set_categories(cats_ref)
    s = ch.series[0]
    for i, col in enumerate(colors):
        p = DataPoint(idx=i); p.graphicalProperties.solidFill = col; s.dPt.append(p)
    s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.showCatName = False
    s.dLbls.showSerName = False; s.dLbls.showPercent = False; s.dLbls.showLegendKey = False
    ch.legend = None; ch.gapWidth = 40
    ch.x_axis.delete = False; ch.y_axis.delete = False
    if horizontal: ch.x_axis.scaling.orientation = 'maxMin'
    ch.y_axis.numFmt = '#,##0'
    ch.width, ch.height = w, h
    ws.add_chart(ch, anchor)

def build_status_sheet(name, cfg):
    CUR.update(THEMES[name]); ws = new_sheet(name, CUR['title'])
    banner(ws, cfg['title'], 'Données Tarkhiss - Produits Cosmétiques et d’Hygiène Corporelle')
    section(ws, 'A4:D4', '📊 INDICATEURS CLÉS')
    ids = rng(cfg['raw'], cfg['idc'], cfg['start'])
    sts = [(x[0], x[1]) for x in cfg['statuses']]; n = len(sts); als = {x[0]: x[2] for x in cfg['statuses'] if len(x) > 2}
    t0 = 15; tl = t0 + n - 1; rA = tl + 1; rT = tl + 2
    cat = f'$D${t0}:$D${rA}'; num = f'$B${t0}:$B${rA}'
    ind = [
     ('Total dossiers', f'=COUNTA({ids})', '#,##0'),
     ('Délivrés', f'=SUMIF({cat},"Délivré",{num})', '#,##0'),
     ('En cours de traitement', f'=SUMIF({cat},"En cours",{num})+SUMIF({cat},"Autre",{num})', '#,##0'),
     ('Non conformes (modifiables)', f'=SUMIF({cat},"Non conforme",{num})', '#,##0'),
     ('Refusés (Rejeté + Irrecevable)', f'=SUMIF({cat},"Refusé",{num})', '#,##0'),
     ('Taux de délivrance', '=IFERROR(B6/B5,0)', '0.0%'),
     ('Taux de refus', '=IFERROR(B9/B5,0)', '0.0%'),
    ]
    fills = [LGREY, GREEN, 'E4EEF9', AMBER, RED, GREEN, RED]
    for k, (lab, f, nf) in enumerate(ind):
        r = 5 + k
        put(ws, f'A{r}', lab, F(10, True, GREY), fill(LGREY), LEFT, BOX)
        put(ws, f'B{r}', f, F(12, True, NAVY), fill(fills[k]), RIGHT, BOX, nf)
        ws.row_dimensions[r].height = 20
    section(ws, 'A13:D13', '📈 RÉPARTITION PAR STATUT (ordre de traitement)')
    header(ws, 'A14:D14', {'A14': 'Statut', 'B14': 'Nombre', 'C14': '%', 'D14': 'Catégorie'})
    sr = rng(cfg['raw'], cfg['stc'], cfg['start'])
    for k, (s, c) in enumerate(sts):
        r = t0 + k; ROWMAP[(name, s)] = r
        put(ws, f'A{r}', s, F(10), fill(CAT_FILL[c]), LEFT, BOX)
        put(ws, f'B{r}', f'=COUNTIF({sr},A{r})' + ''.join(f'+COUNTIF({sr},"{a}")' for a in als.get(s, [])), F(10, True), None, RIGHT, BOX, '#,##0')
        put(ws, f'C{r}', f'=IFERROR(B{r}/$B${rT},0)', F(10), None, RIGHT, BOX, '0.0%')
        put(ws, f'D{r}', c, F(9, False, GREY), fill(CAT_FILL[c]), CENTER, BOX)
    put(ws, f'A{rA}', 'Autres statuts / statut vide', F(10, False, GREY, True), fill(CAT_FILL['Autre']), LEFT, BOX)
    put(ws, f'B{rA}', f'=$B$5-SUM(B{t0}:B{tl})', F(10, True), None, RIGHT, BOX, '#,##0')
    put(ws, f'C{rA}', f'=IFERROR(B{rA}/$B${rT},0)', F(10), None, RIGHT, BOX, '0.0%')
    put(ws, f'D{rA}', 'Autre', F(9, False, GREY), fill(CAT_FILL['Autre']), CENTER, BOX)
    ROWMAP[(name, 'Autres')] = rA
    put(ws, f'A{rT}', 'TOTAL', F(10, True), fill('D9E2F3'), LEFT, BOX)
    put(ws, f'B{rT}', f'=SUM(B{t0}:B{rA})', F(10, True), fill('D9E2F3'), RIGHT, BOX, '#,##0')
    put(ws, f'C{rT}', f'=IFERROR(SUM(C{t0}:C{rA}),0)', F(10, True), fill('D9E2F3'), RIGHT, BOX, '0.0%')
    put(ws, f'D{rT}', f'=IF(B{rT}=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill('D9E2F3'), CENTER, BOX)
    ws.conditional_formatting.add(f'C{t0}:C{rA}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='6FA8DC', showValue=True))
    ws.conditional_formatting.add(f'B{t0}:B{rA}', FormulaRule(formula=[f'B{t0}=0'], font=Font(name='Calibri', color='A6A6A6')))
    put(ws, f'A{rT+2}', f'=IF($B$5=0,"⚠ La feuille « {cfg["raw"]} » ne contient aucune donnée : indicateurs à 0 / N/D.","")', F(10, True, 'B03A2E'))
    r = rT + 4
    put(ws, f'A{r}', '📌 LÉGENDE DES CATÉGORIES', F(10, True, NAVY))
    for k, c in enumerate(['Délivré', 'En cours', 'Non conforme', 'Refusé', 'Autre']):
        put(ws, f'A{r+1+k}', c, F(10), fill(CAT_FILL[c]), LEFT, BOX)
    if als: put(ws, f'A{r+8}', 'Coquilles de saisie regroupées : ' + ' ; '.join(f'« {a} » → {k}' for k, v in als.items() for a in v) + '.', F(8, False, GREY, True))
    put(ws, f'A{r+7}', f"Source : feuille « {cfg['raw']} » (col. {cfg['stc']} = statut, col. {cfg['idc']} = identifiant, lignes {cfg['start']}–{LAST}).", F(8, False, GREY, True))
    INFO[name] = dict(t0=t0, tl=tl, rA=rA, rT=rT, n=n, sts=sts)
    colors = pt_colors([c for _, c in sts])
    make_pie(ws, 'doughnut', 'Répartition par statut', Reference(ws, min_col=1, min_row=t0, max_row=tl),
             Reference(ws, min_col=2, min_row=t0, max_row=tl), colors, 'F4')
    make_bar(ws, 'Dossiers par statut', Reference(ws, min_col=1, min_row=t0, max_row=tl),
             Reference(ws, min_col=2, min_row=t0, max_row=tl), colors, f'F{t0+4}')
    ws.freeze_panes = 'A4'

for nme, cfg in MODS.items(): build_status_sheet(nme, cfg)

# ------------------------------------------------ KPI Identification
name = 'KPI Identification'
CUR.update(THEMES[name]); ws = new_sheet(name, CUR['title'])
banner(ws, '📋 KPI IDENTIFICATION OPÉRATEUR', 'Données Tarkhiss - Produits Cosmétiques et d’Hygiène Corporelle')
section(ws, 'A4:D4', '📊 INDICATEURS CLÉS')
raw = 'identificationOpérateur'
ids = rng(raw, 'B', 6); sts = rng(raw, 'F', 6); typ = rng(raw, 'D', 6)
for k, (lab, f, nf, fl) in enumerate([
    ('Total opérateurs', f'=COUNTA({ids})', '#,##0', LGREY),
    ('Opérateurs identifiés', f'=COUNTIF({sts},"Identifié")', '#,##0', GREEN),
    ('En cours / autres statuts', '=B5-B6', '#,##0', 'E4EEF9'),
    ('Taux de réussite (identifiés)', '=IFERROR(B6/B5,0)', '0.0%', GREEN)]):
    r = 5 + k
    put(ws, f'A{r}', lab, F(10, True, GREY), fill(LGREY), LEFT, BOX)
    put(ws, f'B{r}', f, F(12, True, NAVY), fill(fl), RIGHT, BOX, nf); ws.row_dimensions[r].height = 20
section(ws, 'A10:D10', '📈 RÉPARTITION PAR STATUT')
header(ws, 'A11:D11', {'A11': 'Statut', 'B11': 'Nombre', 'C11': '%', 'D11': 'Catégorie'})
for r, (lab, f, c, cl) in zip([12, 13], [('Identifié', '=B6', 'Délivré', 'Validé'), ('Autres statuts / statut vide', '=B7', 'Autre', 'Autre')]):
    put(ws, f'A{r}', lab, F(10), fill(CAT_FILL[c]), LEFT, BOX)
    put(ws, f'B{r}', f, F(10, True), None, RIGHT, BOX, '#,##0')
    put(ws, f'C{r}', f'=IFERROR(B{r}/$B$14,0)', F(10), None, RIGHT, BOX, '0.0%')
    put(ws, f'D{r}', cl, F(9, False, GREY), fill(CAT_FILL[c]), CENTER, BOX)
put(ws, 'A14', 'TOTAL', F(10, True), fill('D9E2F3'), LEFT, BOX)
put(ws, 'B14', '=SUM(B12:B13)', F(10, True), fill('D9E2F3'), RIGHT, BOX, '#,##0')
put(ws, 'C14', '=IFERROR(SUM(C12:C13),0)', F(10, True), fill('D9E2F3'), RIGHT, BOX, '0.0%')
put(ws, 'D14', '=IF(B14=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill('D9E2F3'), CENTER, BOX)
section(ws, 'A16:D16', "📈 RÉPARTITION PAR TYPE D'OPÉRATEUR")
header(ws, 'A17:D17', {'A17': "Type d'opérateur", 'B17': 'Nombre', 'C17': '%', 'D17': 'Part'})
types = [('Etablissement fabrication / import / export', 'Etablissement de fabrication*'),
         ('Etablissement pharmaceutique industriel', 'Etablissement pharmaceutique*'),
         ('Coopérative', 'Coopérative*'),
         ('Etablissement compléments alimentaires', 'Etablissement exerçant les activités portant sur les compléments alimentaires*')]
for k, (lab, pat) in enumerate(types):
    r = 18 + k
    put(ws, f'A{r}', lab, F(10), None, LEFT, BOX)
    put(ws, f'B{r}', f'=COUNTIF({typ},"{pat}")', F(10, True), None, RIGHT, BOX, '#,##0')
    put(ws, f'C{r}', f'=IFERROR(B{r}/$B$23,0)', F(10), None, RIGHT, BOX, '0.0%')
    put(ws, f'D{r}', f'=C{r}', F(9, False, GREY), None, CENTER, BOX, '0.0%')
put(ws, 'A22', 'Autres types / vide', F(10, False, GREY, True), fill('ECEFF3'), LEFT, BOX)
put(ws, 'B22', '=$B$5-SUM(B18:B21)', F(10, True), None, RIGHT, BOX, '#,##0')
put(ws, 'C22', '=IFERROR(B22/$B$23,0)', F(10), None, RIGHT, BOX, '0.0%')
put(ws, 'D22', '=C22', F(9, False, GREY), None, CENTER, BOX, '0.0%')
put(ws, 'A23', 'TOTAL', F(10, True), fill('D9E2F3'), LEFT, BOX)
put(ws, 'B23', '=SUM(B18:B22)', F(10, True), fill('D9E2F3'), RIGHT, BOX, '#,##0')
put(ws, 'C23', '=IFERROR(SUM(C18:C22),0)', F(10, True), fill('D9E2F3'), RIGHT, BOX, '0.0%')
put(ws, 'D23', '=IF(B23=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill('D9E2F3'), CENTER, BOX)
for rr in ('C12:C13', 'C18:C22'):
    ws.conditional_formatting.add(rr, DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='6FA8DC', showValue=True))
put(ws, 'A25', f"Source : feuille « {raw} » (col. B = ID, D = type, F = statut, lignes 6–{LAST}). Types mutuellement exclusifs (préfixes).", F(8, False, GREY, True))
pal = ['1F6FB2', '2E8B57', 'F4B400', '8E6BBF']
make_pie(ws, 'pie', 'Opérateurs par type', Reference(ws, min_col=1, min_row=18, max_row=21), Reference(ws, min_col=2, min_row=18, max_row=21), pal, 'F4')
make_bar(ws, 'Opérateurs par type', Reference(ws, min_col=1, min_row=18, max_row=21), Reference(ws, min_col=2, min_row=18, max_row=21), pal, 'F21', h=7.5)
ws.freeze_panes = 'A4'

# ------------------------------------------------ Dashboard
CUR.update(THEMES['Dashboard']); ws = new_sheet('Dashboard', CUR['title'])
for col, w in zip('ABCDEFGHIJKL', [36, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]):
    ws.column_dimensions[col].width = w
banner(ws, '📊 TABLEAU DE BORD - REVUE DE DIRECTION', "Plateforme Tarkhiss - Produits Cosmétiques et d'Hygiène Corporelle")
ws.sheet_view.zoomScale = 90
section(ws, 'A4:L4', '🎯 SYNTHÈSE EXÉCUTIVE')
header(ws, 'A5:I5', {'A5': 'Module', 'B5': 'Total dossiers', 'C5': 'Délivrés / Validés', 'D5': 'Taux réussite',
                      'E5': 'En cours*', 'F5': 'Bloqués / Refusés', 'G5': 'Seuil 🟢 ≥', 'H5': 'Seuil 🟡 ≥', 'I5': 'Voyant'})
ws.row_dimensions[5].height = 30
mods = [('1. Identification Opérateur', 'KPI Identification', 0.9, 0.5),
        ("2. Déclaration d'Activité", 'KPI Déclaration', 0.9, 0.5),
        ('3. Certificat Enregistrement (CE)', 'KPI Certificat Enreg.', 0.1, 0.05),
        ('4. Certificat Libre Vente (CLV)', 'KPI Certificat LV', 0.5, 0.3),
        ('5. Autorisation Import. MP (AIMP)', 'KPI Autor. Import.', 0.2, 0.1)]
for k, (lab, sh, g, o) in enumerate(mods):
    r = 6 + k; q = f"'{sh}'!"
    put(ws, f'A{r}', lab, F(10, True), fill(LGREY), LEFT, BOX)
    put(ws, f'B{r}', f'={q}B5', F(10), None, RIGHT, BOX, '#,##0')
    put(ws, f'C{r}', f'={q}B6', F(10), None, RIGHT, BOX, '#,##0')
    put(ws, f'D{r}', f'=IF(B{r}>0,C{r}/B{r},0)', F(10, True), None, RIGHT, BOX, '0.0%')
    put(ws, f'E{r}', f'={q}B7', F(10), None, RIGHT, BOX, '#,##0')
    if sh == 'KPI Identification': put(ws, f'F{r}', '–', F(10, False, GREY), None, RIGHT, BOX)
    else: put(ws, f'F{r}', f'={q}B8+{q}B9', F(10), None, RIGHT, BOX, '#,##0')
    put(ws, f'G{r}', g, F(10, False, '0000FF'), fill('FFFBE6'), RIGHT, BOX, '0%')
    put(ws, f'H{r}', o, F(10, False, '0000FF'), fill('FFFBE6'), RIGHT, BOX, '0%')
    put(ws, f'I{r}', f'=IF(B{r}=0,"⚪ N/D",IF(D{r}>=G{r},"🟢",IF(D{r}>=H{r},"🟡","🔴")))', F(12), None, CENTER, BOX)
    ws.row_dimensions[r].height = 21
for col in 'ABCDEFGHI':
    put(ws, f'{col}11', None, F(10, True), fill('D9E2F3'), RIGHT if col != 'A' else LEFT, BOX)
ws['A11'].value = 'TOTAL'
for col in 'BCEF': ws[f'{col}11'].value = f'=SUM({col}6:{col}10)'; ws[f'{col}11'].number_format = '#,##0'
ws['D11'].value = '=IF(B11>0,C11/B11,0)'; ws['D11'].number_format = '0.0%'
ws.conditional_formatting.add('D6:D10', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='63BE7B', showValue=True))
put(ws, 'A12', '=IF(B8=0,"⚠ Feuille « certificatEnregistrement » vide : CE en N/D (aucune donnée). ","")&"* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté. Seuils en bleu = paramètres modifiables."',
    F(8, False, GREY, True))
CUR.update(sec='7030A0', secfont='FFFFFF'); section(ws, 'A14:L14', '📋 RÉSUMÉ DÉCISIONNEL - CHIFFRES CLÉS')
cards = [('A', 'B', 'OPÉRATEURS IDENTIFIÉS', '=C6', GREEN), ('C', 'D', 'CE DÉLIVRÉS', '=C8', GREEN),
         ('E', 'F', 'CLV DÉLIVRÉS', '=C9', GREEN), ('G', 'H', 'AIMP DÉLIVRÉS', '=C10', GREEN),
         ('I', 'J', 'DOSSIERS EN COURS', '=E11', 'E4EEF9'), ('K', 'L', 'BLOQUÉS / REFUSÉS', '=F11', RED)]
for a, b, lab, f, col in cards:
    merge(ws, f'{a}15:{b}15', lab, F(8, True, GREY), fill(LGREY), CENTER, BOX)
    merge(ws, f'{a}16:{b}16', f, F(28, True, '375623' if col == GREEN else '1F4E79'), fill(col), CENTER, BOX, nf='#,##0')
ws.row_dimensions[15].height = 22; ws.row_dimensions[16].height = 40
CUR.update(sec='C00000', secfont='FFFFFF', hdr='FFC000'); section(ws, 'A18:L18', "⚠️ POINTS D'ALERTE - ACTIONS REQUISES")
header(ws, 'A19:G19', {'A19': 'Indicateur', 'B19': 'Module', 'C19': 'Valeur actuelle', 'D19': 'Seuil attention', 'E19': 'Seuil critique', 'F19': 'Statut', 'G19': 'Priorité'})
ws.row_dimensions[19].height = 30
def rowof(sh, s): return ROWMAP[(sh, s)]
alerts = [
 ('Taux de délivrance - Déclaration', "Déclaration", '=D7', 'taux', 0.5, 0.3, 'B7'),
 ('Taux de délivrance - CE', 'CE', '=D8', 'taux', 0.2, 0.1, 'B8'),
 ('Taux de délivrance - CLV', 'CLV', '=D9', 'taux', 0.5, 0.3, 'B9'),
 ('Taux de délivrance - AIMP', 'AIMP', '=D10', 'taux', 0.2, 0.1, 'B10'),
 ('Déclarations Irrecevables', "Déclaration", f"='KPI Déclaration'!B{rowof('KPI Déclaration','Irrecevable')}", 'nb', 5, 10, 'B7'),
 ('Déclarations Non Conformes', "Déclaration", f"='KPI Déclaration'!B{rowof('KPI Déclaration','Non Conforme')}", 'nb', 50, 100, 'B7'),
 ('Dossiers Irrecevables CE', 'CE', f"='KPI Certificat Enreg.'!B{rowof('KPI Certificat Enreg.','Irrecevable')}", 'nb', 30, 50, 'B8'),
 ('Dossiers en Attente Paiement CE', 'CE', f"='KPI Certificat Enreg.'!B{rowof('KPI Certificat Enreg.','Attente paiement')}", 'nb', 50, 100, 'B8'),
 ('Dossiers Non Conformes AIMP', 'AIMP', f"='KPI Autor. Import.'!B{rowof('KPI Autor. Import.','Non conforme')}", 'nb', 10, 20, 'B10'),
]
for k, (lab, mod, f, typ_, a, c, tot) in enumerate(alerts):
    r = 20 + k
    put(ws, f'A{r}', lab, F(10, True), fill(LGREY), LEFT, BOX)
    put(ws, f'B{r}', mod, F(9), None, CENTER, BOX)
    put(ws, f'C{r}', f, F(10, True), None, RIGHT, BOX, '0.0%' if typ_ == 'taux' else '#,##0')
    put(ws, f'D{r}', a, F(10, False, '0000FF'), fill('FFFBE6'), RIGHT, BOX, '0%' if typ_ == 'taux' else '#,##0')
    put(ws, f'E{r}', c, F(10, False, '0000FF'), fill('FFFBE6'), RIGHT, BOX, '0%' if typ_ == 'taux' else '#,##0')
    op = '<' if typ_ == 'taux' else '>'
    cond = f'IF(C{r}{op}E{r},"🔴 CRITIQUE",IF(C{r}{op}D{r},"🟡 ATTENTION","🟢 OK"))'
    put(ws, f'F{r}', f'=IF({tot}=0,"⚪ N/D",{cond})', F(10, True), None, CENTER, BOX)
    put(ws, f'G{r}', f'=IF(ISNUMBER(SEARCH("CRITIQUE",F{r})),"HAUTE",IF(ISNUMBER(SEARCH("ATTENTION",F{r})),"MOYENNE",IF(ISNUMBER(SEARCH("OK",F{r})),"BASSE","—")))', F(10, True), None, CENTER, BOX)
    ws.row_dimensions[r].height = 19
aL = 20 + len(alerts) - 1
for txt, colr in (('CRITIQUE', RED), ('ATTENTION', AMBER), ('OK', GREEN)):
    ws.conditional_formatting.add(f'F20:G{aL}', FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",$F20))'], fill=fill(colr)))
put(ws, f'A{aL+1}', 'Seuils : « taux » → alerte si la valeur est INFÉRIEURE au seuil ; « nombre » → SUPÉRIEURE. Seuils en bleu = paramètres (valeurs d’origine du fichier ou hypothèses à valider).', F(8, False, GREY, True))

cr = aL + 3
CUR.update(sec='BDD7EE', secfont='1F4E79', hdr='2F5496')
section(ws, f'A{cr}:L{cr}', '📊 GRAPHIQUES')
c1 = Reference(ws, min_col=1, min_row=6, max_row=10)
ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100; ch.style = 10
ch.title = 'Dossiers par module : délivrés / en cours / bloqués'
for col, nm, colr in ((3, 'Délivrés', '2E8B57'), (5, 'En cours', '4FA3D9'), (6, 'Bloqués / Refusés', 'D64541')):
    s = Series(Reference(ws, min_col=col, min_row=6, max_row=10), title=nm)
    s.graphicalProperties.solidFill = colr; ch.series.append(s)
ch.set_categories(c1); ch.legend.position = 'b'; ch.x_axis.delete = False; ch.y_axis.delete = False
ch.y_axis.numFmt = '#,##0'; ch.width, ch.height = 17, 8.5
ws.add_chart(ch, f'A{cr+2}')
tx = BarChart(); tx.type = 'col'; tx.style = 10; tx.title = 'Taux de réussite par module'
tx.add_data(Reference(ws, min_col=4, min_row=6, max_row=10), titles_from_data=False); tx.set_categories(c1)
s = tx.series[0]
for i, colr in enumerate(['1F6FB2', '2E8B57', '8E6BBF', 'F4B400', 'D64541']):
    p = DataPoint(idx=i); p.graphicalProperties.solidFill = colr; s.dPt.append(p)
s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = '0%'
s.dLbls.showCatName = False; s.dLbls.showSerName = False; s.dLbls.showPercent = False; s.dLbls.showLegendKey = False
tx.legend = None; tx.x_axis.delete = False; tx.y_axis.delete = False; tx.y_axis.numFmt = '0%'
tx.y_axis.scaling.min = 0; tx.y_axis.scaling.max = 1; tx.width, tx.height = 17.5, 8.5
ws.add_chart(tx, f'F{cr+2}')
def kref(sh, col):
    i = INFO[sh]; return Reference(wb[sh], min_col=col, min_row=i['t0'], max_row=i['tl'])
def kcolors(sh): return pt_colors([c for _, c in INFO[sh]['sts']])
r2 = cr + 2 + 18
make_pie(ws, 'doughnut', "Déclarations d'activité par statut", kref('KPI Déclaration', 1), kref('KPI Déclaration', 2), kcolors('KPI Déclaration'), f'A{r2}', w=17, h=8.5)
make_pie(ws, 'doughnut', 'Certificats Libre Vente par statut', kref('KPI Certificat LV', 1), kref('KPI Certificat LV', 2), kcolors('KPI Certificat LV'), f'F{r2}', w=17.5, h=8.5)
r3 = r2 + 18
make_bar(ws, 'Autorisations d’importation MP par statut', kref('KPI Autor. Import.', 1), kref('KPI Autor. Import.', 2), kcolors('KPI Autor. Import.'), f'A{r3}', w=17, h=9)
ki = wb['KPI Identification']
make_pie(ws, 'pie', 'Opérateurs par type', Reference(ki, min_col=1, min_row=18, max_row=21), Reference(ki, min_col=2, min_row=18, max_row=21), pal, f'F{r3}', w=17.5, h=9)

dr = r3 + 20
section(ws, f'A{dr}:L{dr}', '📈 DÉTAIL PAR STATUT (liens vers les feuilles KPI)')
def detail(r0, a, b, numcol, pctcol, title, sh, rows_):
    merge(ws, f'{a}{r0}:{pctcol}{r0}', title, F(10, True, 'FFFFFF'), fill('4472C4'), CENTER, BOX)
    merge(ws, f'{a}{r0+1}:{b}{r0+1}', 'Statut', F(9, True, NAVY), fill('D6DCE4'), CENTER, BOX)
    put(ws, f'{numcol}{r0+1}', 'Nombre', F(9, True, NAVY), fill('D6DCE4'), CENTER, BOX)
    put(ws, f'{pctcol}{r0+1}', '%', F(9, True, NAVY), fill('D6DCE4'), CENTER, BOX)
    for k, kr in enumerate(rows_):
        r = r0 + 2 + k; q = f"'{sh}'!"
        merge(ws, f'{a}{r}:{b}{r}', f'={q}A{kr}', F(9), None, LEFT, BOX)
        put(ws, f'{numcol}{r}', f'={q}B{kr}', F(9, True), None, RIGHT, BOX, '#,##0')
        put(ws, f'{pctcol}{r}', f'={q}C{kr}', F(9), None, RIGHT, BOX, '0.0%')
    ws.conditional_formatting.add(f'{pctcol}{r0+2}:{pctcol}{r0+1+len(rows_)}', DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='6FA8DC', showValue=True))
    return r0 + 2 + len(rows_)
def rows_for(sh):
    i = INFO[sh]; return list(range(i['t0'], i['rA'] + 1))
d0 = dr + 2
e1 = detail(d0, 'A', 'A', 'B', 'C', 'DÉCLARATION D’ACTIVITÉ', 'KPI Déclaration', rows_for('KPI Déclaration'))
e2 = detail(d0, 'E', 'F', 'G', 'H', 'AIMP', 'KPI Autor. Import.', rows_for('KPI Autor. Import.'))
e3 = detail(d0, 'I', 'J', 'K', 'L', 'CLV', 'KPI Certificat LV', rows_for('KPI Certificat LV'))
d1 = max(e1, e2, e3) + 1
e4 = detail(d1, 'A', 'A', 'B', 'C', "CERTIFICAT D'ENREGISTREMENT (CE)", 'KPI Certificat Enreg.', rows_for('KPI Certificat Enreg.'))
e5 = detail(d1, 'E', 'F', 'G', 'H', "TYPES D'OPÉRATEURS", 'KPI Identification', list(range(18, 23)))
put(ws, f'A{max(e4,e5)+1}', 'Sources : feuilles KPI (calculées depuis les feuilles brutes). Dernière actualisation : 28/10/2026.', F(8, False, GREY, True))
ws.freeze_panes = 'A4'
ws.page_setup.orientation = 'portrait'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0

for n in ['identificationOpérateur', 'déclarationActivité', 'certificatEnregistrement', 'certificatLibreVente', 'autorisationImportationMatièreP']:
    wb[n].sheet_state = 'visible'; wb[n].sheet_properties.tabColor = '8A97A8'
wb.active = 0
for s in wb.worksheets: s.sheet_view.tabSelected = (s.title == 'Dashboard')
wb.save(OUT)
print('saved', wb.sheetnames)
