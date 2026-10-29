# -*- coding: utf-8 -*-
"""Tests du rapport de période Tarkhiss (Vue annuelle) : python tests/test_tarkhiss_annual.py

Serveur réel (copie isolée de l'application) alimenté par une COPIE des fichiers Tarkhiss
(SANAD_TARKHISS_DATA, défaut : data/tarkhiss — jamais modifiés). Vérifie : droits par rôle, exports
PDF / PPTX / Excel (garde, résumé, mois), aperçu et .eml (pièces jointes au choix, corps = aperçu),
absence de toute mention « SANAD » dans les fichiers exportés/envoyés, non-régression mensuelle.
"""
import email
import io
import os
import re
import sys
import time
import zipfile
from email import policy

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_pchc_export import ADMIN, Server, _pdf_pages, _pdf_text  # noqa: E402

DATA = os.environ.get("SANAD_TARKHISS_DATA") or os.path.join(ROOT, "data", "tarkhiss")
PWD = "Passw0rd#2026"


def _months_available():
    return sorted({m.group(1) for fn in os.listdir(DATA) for m in [re.match(r"^(?:calls|emails|analysis)_(\d{4}-\d{2})\.json$", fn)] if m})


def _no_sanad_zip(data, label):
    z = zipfile.ZipFile(io.BytesIO(data))
    for n in z.namelist():
        assert b"sanad" not in z.read(n).lower(), f"mention SANAD dans {label}:{n}"


def test_annual_report_end_to_end():
    months = _months_available()
    assert months, f"aucune donnée Tarkhiss dans {DATA}"
    start, end = months[0][:4] + "-01", months[-1][:4] + "-12"
    with Server(seed={"tarkhiss": DATA}) as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        admin.post(U + "/api/tarkhiss/settings", json={"contacts": [{"email": "direction@ammps.gov.ma", "name": "Direction", "role": "to"}]})
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            assert admin.post(U + "/api/users", json={"username": name, "password": PWD, "role": role, "name": name, "modules": ["tarkhiss"]}).status_code == 200
        sup = srv.session({"username": "sup", "password": PWD})
        hot = srv.session({"username": "hot", "password": PWD})

        # Données de la Vue annuelle (écran) = base du rapport
        view = admin.get(U + f"/api/tarkhiss/annual/{start[:4]}?start={start}&end={end}").json()
        n = len(view["months"])
        assert n >= 1 and view["totals"]["calls"] == sum(m["calls"] for m in view["months"])
        print(f"  {n} mois avec données, période {view['period_label']}")

        # Droits : hotliner interdit ; superviseur exporte mais n'envoie pas
        assert hot.get(U + f"/api/tarkhiss/annual-report/pdf?start={start}&end={end}").status_code == 403
        assert hot.get(U + f"/api/tarkhiss/annual/{start[:4]}").status_code == 200   # vue écran inchangée
        assert sup.post(U + "/api/tarkhiss/annual-send-email", json={"start": start, "end": end}).status_code == 403
        assert admin.get(U + "/api/tarkhiss/annual-report/pdf?start=2026-13&end=2026-01").status_code == 400

        # Descriptifs par bloc + prompts (chiffres de la période), admin seul en écriture
        nd = admin.get(U + f"/api/tarkhiss/annual-notes?start={start}&end={end}").json()
        keys = [b["key"] for b in nd["blocks"]]
        assert keys == ["indicateurs", "charge", "top3", "detail", "synthese"]
        assert str(view["totals"]["calls"]) in nd["prompts"]["indicateurs"] and view["months"][0]["label"] in nd["prompts"]["detail"]
        assert hot.get(U + f"/api/tarkhiss/annual-notes?start={start}&end={end}").status_code == 403
        notes = {k: f"Descriptif {k} rédigé pour le test." for k in keys}
        assert sup.post(U + "/api/tarkhiss/annual-notes", json={"start": start, "end": end, "notes": notes}).status_code == 403
        assert admin.post(U + "/api/tarkhiss/annual-notes", json={"start": start, "end": end, "notes": notes}).json()["ok"]
        assert sup.get(U + f"/api/tarkhiss/annual-notes?start={start}&end={end}").json()["notes"] == notes

        # PDF : garde + résumé + mois (1-2 pages par mois)
        t = time.time()
        r = sup.get(U + f"/api/tarkhiss/annual-report/pdf?start={start}&end={end}")
        assert r.status_code == 200 and r.content[:5] == b"%PDF-", r.text[:300]
        assert f"Rapport_Tarkhiss_{start}_{end}.pdf" in r.headers["Content-Disposition"]
        pages = _pdf_pages(r.content)
        print(f"  PDF {len(r.content) / 1e6:.2f} Mo, {pages} pages ({time.time() - t:.1f} s)")
        if pages is not None:
            assert 2 + n <= pages <= 2 + 2 * n + 2, pages
        text = _pdf_text(r.content)
        if text is not None:
            assert "sanad" not in text.lower()
            assert "Page 2 /" in text and view["period_label"].split(" — ")[0] in text
            flat = " ".join(text.split())
            assert all(notes[k] in flat for k in keys), "descriptifs absents du PDF"
        pdf = r.content

        # PPTX : garde + 2 diapositives de résumé + 3 « Analyse » (5 descriptifs, 2 par diapo) + 2 par mois
        r = sup.get(U + f"/api/tarkhiss/annual-report/pptx?start={start}&end={end}")
        assert r.status_code == 200
        from pptx import Presentation
        prs = Presentation(io.BytesIO(r.content))
        assert len(prs.slides) == 3 + 3 + 2 * n, len(prs.slides)
        ptxt = " ".join(p.text for sl in prs.slides for sh in sl.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs)
        assert all(notes[k] in ptxt for k in keys), "descriptifs absents du PPTX"
        for slide in prs.slides:
            for sh in slide.shapes:
                assert sh.left >= 0 and sh.top >= 0 and sh.left + sh.width <= prs.slide_width + 10 and sh.top + sh.height <= prs.slide_height + 10, sh.name
        _no_sanad_zip(r.content, "pptx")
        pptx = r.content

        # Excel : synthèse + une feuille par mois
        r = sup.get(U + f"/api/tarkhiss/annual-report/xlsx?start={start}&end={end}")
        assert r.status_code == 200
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(r.content))
        assert wb.sheetnames[0] == "Synthèse" and len(wb.sheetnames) == 1 + n, wb.sheetnames
        assert wb.sheetnames[1:] == [m["label"] for m in view["months"]]
        cells = [c.value for row in wb["Synthèse"].iter_rows() for c in row if isinstance(c.value, str)]
        assert all(notes[k] in cells for k in keys), "descriptifs absents de l'Excel"
        _no_sanad_zip(r.content, "xlsx")
        xlsx = r.content

        # Aperçu de l'email = éléments de la Vue annuelle
        pv = admin.get(U + f"/api/tarkhiss/annual-email-preview?start={start}&end={end}&greeting=Bonjour%20Direction%2C").json()
        html = pv["html"]
        for needle in ("Tendance mensuelle", "Charge moyenne par jour de la semaine", "Détail par mois", "Bonjour Direction,", view["months"][0]["label"]):
            assert needle in html, needle
        assert "sanad" not in html.lower() and "sanad" not in pv["subject"].lower()
        assert all(notes[k] in html for k in keys), "descriptifs absents du corps de l'email"

        # Envoi : pièces jointes au choix (PDF + Excel ici)
        r = admin.post(U + "/api/tarkhiss/annual-send-email", json={"start": start, "end": end, "greeting": "Bonjour Direction,",
                                                                    "attachments": {"pdf": True, "pptx": False, "xlsx": True}})
        d = r.json()
        assert r.status_code == 200 and d["ok"] and d["attachments"] == [f"Rapport_Tarkhiss_{start}_{end}.pdf", f"Rapport_Tarkhiss_{start}_{end}.xlsx"], d
        raw = admin.get(U + d["download"]).content
        assert b"sanad" not in raw.lower()
        msg = email.message_from_bytes(raw, policy=policy.default)
        assert msg.get_content_type() == "multipart/mixed" and msg["To"] == "direction@ammps.gov.ma"
        assert msg["X-Unsent"] == "1", "Outlook doit ouvrir le message en brouillon"
        assert d["opened"] is False and "SANAD_NO_OUTLOOK" in (d["error"] or "")   # ouverture Outlook neutralisée en test
        assert msg.get_body(("html",)).get_content().strip() == html.strip(), "corps ≠ aperçu"
        # Même générateur que les téléchargements (seul l'horodatage « généré le » diffère)
        atts = {p.get_filename(): p.get_content() for p in msg.iter_attachments()}
        assert _pdf_pages(atts[f"Rapport_Tarkhiss_{start}_{end}.pdf"]) == _pdf_pages(pdf)
        assert openpyxl.load_workbook(io.BytesIO(atts[f"Rapport_Tarkhiss_{start}_{end}.xlsx"])).sheetnames == wb.sheetnames
        r = admin.post(U + "/api/tarkhiss/annual-send-email", json={"start": start, "end": end, "attachments": {"pptx": True}})
        msg = email.message_from_bytes(admin.get(U + r.json()["download"]).content, policy=policy.default)
        (att,) = list(msg.iter_attachments())
        assert att.get_filename() == f"Rapport_Tarkhiss_{start}_{end}.pptx"
        assert len(Presentation(io.BytesIO(att.get_content())).slides) == len(Presentation(io.BytesIO(pptx)).slides)
        log = admin.get(U + "/api/tarkhiss/send-log").json()
        assert log[0]["ym"] == f"{start}_{end}" and log[0]["attachments"] == [f"Rapport_Tarkhiss_{start}_{end}.pptx"]
        audit = open(os.path.join(srv.dir, "data", "audit_log.jsonl"), encoding="utf-8").read()
        assert audit.count('"annual_export"') >= 3 and audit.count('"annual_send_email"') == 2

        # Non-régression : exports et e-mail mensuels, sans mention SANAD
        ym = view["months"][-1]["ym"]
        r = admin.post(U + f"/api/tarkhiss/export-pdf/{ym}", json={})
        assert r.status_code == 200 and r.content[:5] == b"%PDF-"
        t = _pdf_text(r.content)
        assert t is None or "sanad" not in t.lower()
        r = admin.get(U + f"/api/tarkhiss/export-xlsx/{ym}")
        assert r.status_code == 200
        _no_sanad_zip(r.content, "xlsx mensuel")
        r = admin.post(U + f"/api/tarkhiss/preview-email/{ym}", json={})
        mh = r.json()["html"]
        assert r.status_code == 200 and "sanad" not in mh.lower()
        # Analyse mensuelle : plus d'« action corrective » ni de mots-clés dans les rapports
        assert "Action :" not in mh and "Mots-clés" not in mh
        r = admin.get(U + f"/api/tarkhiss/export-xlsx/{ym}")
        assert "Action corrective" not in [c.value for c in openpyxl.load_workbook(io.BytesIO(r.content))["Dashboard"][9]]
        r = admin.post(U + f"/api/tarkhiss/send-email/{ym}", json={"to": ["direction@ammps.gov.ma"]})
        raw = admin.get(U + f"/api/tarkhiss/send-log/download/{r.json()['file'].split('/')[-1]}").content
        assert email.message_from_bytes(raw, policy=policy.default)["X-Unsent"] == "1"


if __name__ == "__main__":
    t = time.time()
    try:
        test_annual_report_end_to_end()
        print(f"OK    test_annual_report_end_to_end ({time.time() - t:.1f} s)")
    except Exception:
        import traceback
        traceback.print_exc()
        print("ÉCHEC test_annual_report_end_to_end")
        sys.exit(1)
