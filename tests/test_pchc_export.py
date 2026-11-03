# -*- coding: utf-8 -*-
"""Tests Reporting Métier PCHC : python tests/test_pchc_export.py   (ou : python -m pytest tests/test_pchc_export.py)

1. pchc_compute (classeur d'exemple) = valeurs recalculées du classeur de référence.
2. Excel généré = référence, cellule par cellule (valeurs, formules, styles, fusions, largeurs,
   graphiques, mises en forme conditionnelles) ; écart toléré 0.
3. Serveur réel (copie isolée de l'application, data/ vierge) + requests : droits par rôle, import,
   contrôle qualité bloquant, exports Excel / PPTX / PDF, .eml (3 pièces jointes, types MIME, ouverture,
   corps HTML = aperçu), Excel allégé, désactivation du volet, compatibilité du Dashboard « période ».
Fichiers : SANAD_PCHC_SAMPLE (classeur brut) et SANAD_PCHC_REFERENCE (classeur de référence), sinon
recherche automatique dans reference/.
"""
import email
import glob
import io
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import warnings
from datetime import datetime
from email import policy

warnings.simplefilter("ignore")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl  # noqa: E402
import requests  # noqa: E402

import pchc_compare  # noqa: E402
import pchc_excel  # noqa: E402
import pchc_raw  # noqa: E402

UPDATE = datetime(2026, 10, 28)
ADMIN = {"username": "admin", "password": "Admin#12345678", "name": "Admin"}


def _first(env, patterns):
    if os.environ.get(env):
        return os.environ[env]
    for p in patterns:
        hits = sorted(glob.glob(os.path.join(ROOT, p)))
        if hits:
            return hits[-1]
    raise FileNotFoundError(f"Fichier introuvable ({env} / {patterns})")


SAMPLE = _first("SANAD_PCHC_SAMPLE", ["reference/classeur_brut_exemple.xlsx", "reference/*dashboardTarkhisss.xlsx.xlsx"])
REFERENCE = _first("SANAD_PCHC_REFERENCE", ["reference/Rapport Tarkhiss_PCHC_Metier_YTD.xlsx"])
# Parties du paquet régénérées à chaque export (données brutes, chaînes, date de modification) ;
# toutes les autres (6 feuilles de restitution, 16 graphiques, styles, thème…) doivent être identiques à l'octet.
REGENERATED = {"xl/sharedStrings.xml", "docProps/core.xml"} | {f"xl/worksheets/sheet{i}.xml" for i in range(7, 12)}


def _parsed():
    with open(SAMPLE, "rb") as f:
        data = f.read()
    return pchc_raw.parse_workbook(data, os.path.basename(SAMPLE))


# ----------------------------------------------------------------------------- 1. calcul
def test_compute_equals_reference_recalculated_values():
    parsed, rep = _parsed()
    assert not rep["blocking"], rep["blocking"]
    c = pchc_raw.pchc_compute(parsed)
    ref = openpyxl.load_workbook(REFERENCE, data_only=True)
    d = ref["Dashboard"]
    for i, s in enumerate(c["synthesis"]):
        r = 6 + i
        assert d[f"B{r}"].value == s["total"] and d[f"C{r}"].value == s["delivered"] and d[f"E{r}"].value == s["encours"]
        assert d[f"D{r}"].value == s["taux"], (r, d[f"D{r}"].value, s["taux"])
        assert d[f"I{r}"].value == s["voyant"]
        assert d[f"F{r}"].value == ("–" if s["bloques"] is None else s["bloques"])
    t = c["totals"]
    assert (d["B11"].value, d["C11"].value, d["E11"].value, d["F11"].value) == (t["total"], t["delivered"], t["encours"], t["bloques"])
    for i, a in enumerate(c["alerts"]):
        r = 20 + i
        assert (d[f"F{r}"].value, d[f"G{r}"].value) == (a["status"], a["priority"]), r
    for key, mod in c["modules"].items():
        ws = ref[pchc_raw.MOD[key]["kpi"]]
        assert ws["B5"].value == mod["total"] and ws["B6"].value == mod["delivered"]
        if key != "identification":
            for k, row in enumerate(mod["statuses"]):
                assert ws[f"A{15 + k}"].value == row["status"] and ws[f"B{15 + k}"].value == row["count"], (key, row)


# ----------------------------------------------------------------------------- 2. Excel
def test_excel_identical_to_reference():
    parsed, _ = _parsed()
    data = pchc_excel.build_xlsx(parsed, pchc_raw.pchc_compute(parsed), UPDATE)
    diffs = pchc_compare.compare(data, REFERENCE)
    assert not diffs, "\n".join(diffs[:20])
    import zipfile
    zr, zg = zipfile.ZipFile(REFERENCE), zipfile.ZipFile(io.BytesIO(data))
    assert sorted(zr.namelist()) == sorted(zg.namelist())
    def view_neutral(b):
        """Seul écart voulu : le classeur s'ouvre sur le Dashboard (onglet actif / sélectionné)."""
        return re.sub(rb' (tabSelected="1"|firstSheet="\d+"|activeTab="\d+")', b"", b)
    different = [n for n in zr.namelist() if n not in REGENERATED and view_neutral(zr.read(n)) != view_neutral(zg.read(n))]
    assert not different, f"parties différentes de la référence : {different}"
    wbx = zg.read("xl/workbook.xml").decode("utf-8")
    assert "activeTab=" not in wbx and 'tabSelected="1"' in zg.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert all('tabSelected="1"' not in zg.read(f"xl/worksheets/sheet{i}.xml").decode("utf-8") for i in range(2, 7))
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert sum(len(ws._charts) for ws in wb.worksheets) == 16 and len(wb["Dashboard"]._charts) == 6
    assert [ws.sheet_state for ws in wb.worksheets][6:] == ["hidden"] * 5
    res = pchc_excel.recalc_check(data)  # LibreOffice si installé : 0 erreur, valeurs = cache
    if res["available"]:
        assert res["ok"], res
    else:
        print("  (LibreOffice absent : recalcul externe non exécuté —", res["detail"], ")")


def test_excel_light_has_no_raw_sheets_and_same_values():
    parsed, _ = _parsed()
    comp = pchc_raw.pchc_compute(parsed)
    light = pchc_excel.build_xlsx(parsed, comp, UPDATE, light=True)
    wb = openpyxl.load_workbook(io.BytesIO(light), data_only=True)
    assert wb.sheetnames == ["Dashboard", "KPI Identification", "KPI Déclaration", "KPI Certificat Enreg.", "KPI Certificat LV", "KPI Autor. Import."]
    full = openpyxl.load_workbook(REFERENCE, data_only=True)
    for name in wb.sheetnames:
        for row in full[name].iter_rows():
            for c in row:
                if c.column == 1 and isinstance(c.value, str) and c.value.startswith("Source"):
                    continue   # mention « version allégée » dans la note de source
                assert wb[name][c.coordinate].value == c.value, (name, c.coordinate)
    assert not any("!" in str(c.value) and "COUNT" in str(c.value) for ws in openpyxl.load_workbook(io.BytesIO(light)).worksheets for r in ws.iter_rows() for c in r)


def test_reference_workbook_can_be_imported_itself():
    """Le classeur de référence complet (avec Dashboard/KPI et feuilles brutes masquées) s'importe tel quel."""
    with open(REFERENCE, "rb") as f:
        parsed, rep = pchc_raw.parse_workbook(f.read(), os.path.basename(REFERENCE))
    assert not rep["blocking"] and not rep["missing"], rep
    assert {k: v["rows"] for k, v in rep["sheets"].items()} == {"identification": 1266, "declaration": 663, "ce": 14373, "clv": 153, "aimp": 2194}
    assert set(rep["ignored"]) == {"Dashboard", "KPI Identification", "KPI Déclaration", "KPI Certificat Enreg.", "KPI Certificat LV", "KPI Autor. Import."}
    sample, _ = _parsed()
    assert pchc_raw.pchc_compute(parsed)["totals"] == pchc_raw.pchc_compute(sample)["totals"]


def test_import_report_and_mismatch_detection():
    _, rep = _parsed()
    assert rep["typos"]["ce"] == {"« Editon acte » → Edition acte": 1}
    assert rep["typos"]["aimp"] == {"« Sénace de travail » → Séance de travail": 4}
    assert not rep["missing"] and not rep["unknown_statuses"]
    _, bad = pchc_raw.parse_workbook(mismatched_workbook(), "mauvais.xlsx")
    assert any("identificationOpérateur" in b and "Déclaration" in b for b in bad["blocking"]), bad["blocking"]
    assert "certificatEnregistrement" in bad["missing"]


def mismatched_workbook():
    """identificationOpérateur contenant des colonnes de Déclaration ; CE absent ; statut inconnu en CLV."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    dec = ["  ", "Réf. Dossier", "Etablissement", "Type demande", "Type activités", "Responsable évaluation", "Date dépot", "Statut"]
    for name, hdr, rows in (
            ("identificationOpérateur", dec, [["  ", "IS-DEC-1/2026", "X", "Déclaration initiale", "Import", None, "2026-01-01", "Délivré"]]),
            ("DECLARATIONACTIVITE", dec, [["  ", "IS-DEC-2/2026", "Y", "Déclaration initiale", "Import", None, "2026-01-02", "délivré"]]),
            ("certificatLibreVente", ["  ", "Réf. dossier", "Etablissement", "Nom produit", "N° CE", "Date dépôt", "Evaluateur", "Statut"],
             [["  ", "PCHC-CLV-1/2026", "Z", "P", "N1", "01/01/2026", None, "Statut bizarre"]]),
            ("autorisationImportationMatièreP", ["  ", "Réf.Dossier", "Importateur", "Date dépôt", "Evaluateur", "Statut"],
             [["  ", "PCHC-AIMP-1/2026", "W", "01/01/2026", None, "Sénace de travail"]])):
        ws = wb.create_sheet(name)
        ws.append([])
        ws.append([])
        ws.append([])
        ws.append([])
        ws.append(hdr)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------- 3. serveur réel
def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


class Server:
    """Copie isolée de l'application (code seul, data/ vierge) lancée sur un port libre.
    seed = {volet: dossier source} : copie de fichiers de données (lecture seule de la source)."""

    def __init__(self, seed=None):
        self.seed = seed or {}

    def __enter__(self):
        self.dir = tempfile.mkdtemp(prefix="sanad_test_")
        for module, src in self.seed.items():
            os.makedirs(os.path.join(self.dir, "data", module), exist_ok=True)
            for fn in os.listdir(src):
                if re.match(r"^(calls|emails|analysis|meta_calls|meta_emails)_\d{4}-\d{2}\.json$", fn):
                    shutil.copy2(os.path.join(src, fn), os.path.join(self.dir, "data", module, fn))
        for fn in os.listdir(ROOT):
            if fn.endswith(".py"):
                shutil.copy2(os.path.join(ROOT, fn), self.dir)
        shutil.copytree(os.path.join(ROOT, "templates"), os.path.join(self.dir, "templates"))
        shutil.copytree(os.path.join(ROOT, "pchc_template"), os.path.join(self.dir, "pchc_template"))
        shutil.copytree(os.path.join(ROOT, "static"), os.path.join(self.dir, "static"), ignore=shutil.ignore_patterns("uploads"))
        self.port = _free_port()
        env = dict(os.environ, SANAD_NO_AUTO_SNAPSHOT="1", SANAD_NO_OUTLOOK="1", SANAD_PORT=str(self.port))
        boot = (f"import sys; sys.path.insert(0, {self.dir!r}); import app; "
                f"app.NETWORK_STATE['port']={self.port}; app.app.run(host='127.0.0.1', port={self.port}, use_reloader=False)")
        self.log = open(os.path.join(self.dir, "server.log"), "w")
        self.proc = subprocess.Popen([sys.executable, "-c", boot], cwd=self.dir, env=env, stdout=self.log, stderr=subprocess.STDOUT)
        self.url = f"http://127.0.0.1:{self.port}"
        for _ in range(100):
            try:
                requests.get(self.url + "/api/modules", timeout=1)
                return self
            except requests.ConnectionError:
                time.sleep(0.2)
        raise RuntimeError("Serveur non démarré : " + open(os.path.join(self.dir, "server.log")).read()[-2000:])

    def session(self, user=None):
        s = requests.Session()
        if user:
            r = s.post(self.url + "/api/auth/login", json=user)
            assert r.status_code == 200, r.text
        return s

    def __exit__(self, *exc):
        self.proc.terminate()
        self.proc.wait(10)
        self.log.close()
        shutil.rmtree(self.dir, ignore_errors=True)


def _pdf_pages(data):
    """Nombre de pages (pypdfium2 ou pypdf si installés ; sinon objets non compressés) — None si indéterminable."""
    try:
        import pypdfium2 as pdfium
        return len(pdfium.PdfDocument(data))
    except ImportError:
        pass
    try:
        import pypdf
        return len(pypdf.PdfReader(io.BytesIO(data)).pages)
    except ImportError:
        pass
    n = len(re.findall(rb"/Type\s*/Page(?!s)", data))
    return n or None


def _pdf_text(data):
    exe = shutil.which("pdftotext")
    if exe:
        p = subprocess.run([exe, "-", "-"], input=data, capture_output=True)
        return p.stdout.decode("utf-8", "replace")
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(data)
        return "\n".join(pg.get_textpage().get_text_range() for pg in doc)
    except ImportError:
        return None


def test_server_end_to_end():
    with Server() as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            r = admin.post(U + "/api/users", json={"username": name, "password": "Passw0rd#2026", "role": role, "name": name, "modules": ["tarkhiss", "pchc"]})
            assert r.status_code == 200, r.text
        sup = srv.session({"username": "sup", "password": "Passw0rd#2026"})
        hot = srv.session({"username": "hot", "password": "Passw0rd#2026"})

        # Avant import
        assert admin.get(U + "/api/pchc/summary").json()["has_data"] is False
        r = admin.get(U + "/api/pchc/export/xlsx")
        assert r.status_code == 409 and "Aucun import" in r.json()["error"]
        assert admin.get(U + "/api/pchc/eml").status_code == 409

        # Droits : hotliner interdit, superviseur sans import
        assert hot.get(U + "/api/pchc/summary").status_code == 403
        assert hot.get(U + "/api/pchc/export/pdf").status_code == 403
        with open(SAMPLE, "rb") as f:
            sample = f.read()
        assert sup.post(U + "/api/pchc/import", files={"file": ("x.xlsx", sample)}).status_code == 403

        # Contrôle qualité bloquant, puis import forcé
        r = admin.post(U + "/api/pchc/import", files={"file": ("mauvais.xlsx", mismatched_workbook())})
        assert r.status_code == 422 and r.json()["blocking"] and r.json()["can_force"]
        r = admin.post(U + "/api/pchc/import", files={"file": ("mauvais.xlsx", mismatched_workbook())}, data={"force": "1"})
        assert r.status_code == 200 and r.json()["ok"]

        # Import du classeur d'exemple
        t = time.time()
        r = admin.post(U + "/api/pchc/import", files={"file": (os.path.basename(SAMPLE), sample)})
        d = r.json()
        assert r.status_code == 200 and d["ok"], r.text
        print(f"  import {time.time() - t:.1f} s — lignes {d['result']}")
        assert d["result"] == {"identification": 1266, "declaration": 663, "ce": 14373, "clv": 153, "aimp": 2194}
        assert d["snapshot"] and d["snapshot"].endswith("_pchc-import.zip")
        assert os.path.isfile(os.path.join(srv.dir, "data", "_snapshots", d["snapshot"]))
        for k in ("identification", "declaration", "ce", "clv", "aimp", "meta", "history"):
            assert os.path.isfile(os.path.join(srv.dir, "data", "pchc", "raw", f"{k}.json")), k
        audit = open(os.path.join(srv.dir, "data", "audit_log.jsonl"), encoding="utf-8").read()
        assert '"pchc_import"' in audit and '"pchc_import_refused"' in audit

        s = sup.get(U + "/api/pchc/summary").json()
        assert s["has_data"] and s["summary"]["totals"]["total"] == 18649

        # Exports (superviseur autorisé)
        r = sup.get(U + "/api/pchc/export/xlsx?date=2026-10-28")
        assert r.status_code == 200 and r.headers["Content-Type"].startswith(pchc_excel.XLSX_MIME)
        assert "Reporting_PCHC_2026-10.xlsx" in r.headers["Content-Disposition"]
        diffs = pchc_compare.compare(r.content, REFERENCE)
        assert not diffs, "\n".join(diffs[:20])
        xlsx = r.content
        # Bouton « Excel consolidé » du Dashboard (ancienne route) → même livrable (gabarit)
        r = sup.get(U + "/api/pchc/export-xlsx?start=2026-01-01&end=2026-10-28&date=2026-10-28")
        assert r.status_code == 200 and r.content == xlsx, "le Dashboard doit exporter le gabarit"
        r = sup.get(U + "/api/pchc/export/pptx?date=2026-10-28")
        assert r.status_code == 200
        from pptx import Presentation
        prs = Presentation(io.BytesIO(r.content))
        assert len(prs.slides) == 9 and abs(prs.slide_width / prs.slide_height - 16 / 9) < 0.01
        pptx_bytes = r.content
        for slide in prs.slides:
            for sh in slide.shapes:
                assert sh.left >= 0 and sh.top >= 0 and sh.left + sh.width <= prs.slide_width + 10 and sh.top + sh.height <= prs.slide_height + 10, sh.name
        r = admin.get(U + "/api/pchc/export/pdf?date=2026-10-28")
        assert r.status_code == 200 and r.content[:5] == b"%PDF-"
        pdf = r.content
        # Boutons PPTX / PDF du Dashboard (anciennes routes POST) → mêmes livrables officiels
        body = {"start": "2026-01-01", "end": "2026-10-28", "charts": {}}
        r = admin.post(U + "/api/pchc/export-pptx?date=2026-10-28", json=body)
        assert r.status_code == 200 and r.content == pptx_bytes and "Reporting_PCHC_2026-10.pptx" in r.headers["Content-Disposition"]
        r = admin.post(U + "/api/pchc/export-pdf?date=2026-10-28", json=body)
        assert r.status_code == 200 and r.content == pdf and "Reporting_PCHC_2026-10.pdf" in r.headers["Content-Disposition"]
        pages = _pdf_pages(pdf)
        assert pages in (9, None), pages
        if pages is None:
            print("  (pypdfium2 / pypdf absents : nombre de pages du PDF non vérifié)")
        text = _pdf_text(pdf)
        if text is not None:
            assert "18 649" in text.replace(" ", " ").replace("\xa0", " ") and "Page 2 / 9" in text
        else:
            print("  (pdftotext / pypdfium2 absents : texte du PDF non vérifié)")

        # .eml : aperçu puis génération
        pv = sup.get(U + "/api/pchc/eml-preview?date=2026-10-28").json()
        assert pv["ok"] and len(pv["attachments"]) == 3 and not pv["too_big"]
        r = sup.get(U + "/api/pchc/eml?date=2026-10-28&xlsx=1&pptx=1&pdf=1")
        assert r.status_code == 200, r.text
        msg = email.message_from_bytes(r.content, policy=policy.default)
        assert msg.get_content_type() == "multipart/mixed" and msg["X-Unsent"] == "1"
        parts = list(msg.iter_parts())
        assert parts[0].get_content_type() == "multipart/alternative"
        html = parts[0].get_body(("html",)).get_content()
        no_img = lambda h: re.sub(r'src="(data|cid):[^"]+"', 'src=""', h).strip()   # data: à l'aperçu, pièces CID dans le .eml
        assert no_img(html) == no_img(pv["html"]), "corps HTML ≠ aperçu"
        assert 'src="cid:' in html and "data:image" not in html, "images non converties en CID"
        atts = {p.get_filename(): (p.get_content_type(), p.get_content()) for p in msg.iter_attachments()}
        assert set(atts) == {"Reporting_PCHC_2026-10.xlsx", "Reporting_PCHC_2026-10.pptx", "Reporting_PCHC_2026-10.pdf"}
        assert atts["Reporting_PCHC_2026-10.xlsx"][0] == pchc_excel.XLSX_MIME
        assert atts["Reporting_PCHC_2026-10.pptx"][0] == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        assert atts["Reporting_PCHC_2026-10.pdf"][0] == "application/pdf"
        assert atts["Reporting_PCHC_2026-10.xlsx"][1] == xlsx and atts["Reporting_PCHC_2026-10.pptx"][1] == pptx_bytes
        openpyxl.load_workbook(io.BytesIO(atts["Reporting_PCHC_2026-10.xlsx"][1]))
        Presentation(io.BytesIO(atts["Reporting_PCHC_2026-10.pptx"][1]))
        assert atts["Reporting_PCHC_2026-10.pdf"][1][:5] == b"%PDF-"
        print(f"  .eml {len(r.content) / 1e6:.2f} Mo")

        # Excel allégé seul
        r = admin.get(U + "/api/pchc/eml?date=2026-10-28&xlsx=1&pptx=0&pdf=0&light=1")
        msg = email.message_from_bytes(r.content, policy=policy.default)
        (name, part), = [(p.get_filename(), p) for p in msg.iter_attachments()]
        assert name == "Reporting_PCHC_2026-10_allege.xlsx" and len(openpyxl.load_workbook(io.BytesIO(part.get_content())).sheetnames) == 6
        assert admin.get(U + "/api/pchc/eml?xlsx=0&pptx=0&pdf=0").status_code == 400

        # Nom de fichier accentué encodé RFC 2231
        import pchc_deliverables
        raw = pchc_deliverables.build_eml("t", ["a@b.c"], [], "<p>x</p>", [(b"x", "application/pdf", "Synthèse é.pdf")])
        assert b"filename*=utf-8''Synth%C3%A8se%20%C3%A9.pdf" in raw
        assert email.message_from_bytes(raw, policy=policy.default).get_body(("html",)) is not None

        # Historique, journal d'envoi, audit
        hist = admin.get(U + "/api/pchc/summary").json()["history"]
        assert {h["type"] for h in hist} >= {"import", "export", "eml"}
        assert any(e.get("file") == "Reporting_PCHC_2026-10.eml" for e in admin.get(U + "/api/pchc/send-log").json())
        audit = open(os.path.join(srv.dir, "data", "audit_log.jsonl"), encoding="utf-8").read()
        assert audit.count('"pchc_export"') >= 3 and '"pchc_eml"' in audit

        # Compatibilité : Dashboard « période » existant alimenté par le même import
        legacy = admin.get(U + "/api/pchc/data?start=2000-01-01&end=2100-01-01").json()
        assert legacy["categories"]["ce"]["total"] > 0

        # Désactivation du volet (Administration)
        assert admin.post(U + "/api/global-settings", json={"pchc_enabled": False}).status_code == 200
        assert admin.get(U + "/api/modules").json()["pchc"]["ready"] is False
        assert admin.get(U + "/api/pchc/summary").status_code == 403
        assert admin.post(U + "/api/global-settings", json={"pchc_enabled": True}).status_code == 200
        assert admin.get(U + "/api/pchc/summary").status_code == 200


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for fn in tests:
        t = time.time()
        try:
            fn()
            print(f"OK    {fn.__name__} ({time.time() - t:.1f} s)")
        except Exception as e:  # noqa: BLE001
            failed += 1
            import traceback
            traceback.print_exc()
            print(f"ÉCHEC {fn.__name__} : {e}")
    print(f"{len(tests) - failed}/{len(tests)} tests OK")
    sys.exit(1 if failed else 0)
