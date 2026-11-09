# -*- coding: utf-8 -*-
"""Tests du rapport de période Moussanada (Vue annuelle) : python tests/test_moussanada_annual.py

Serveur réel (copie isolée) alimenté par une COPIE des données Moussanada (SANAD_MOUSSANADA_DATA,
défaut data/moussanada, jamais modifiées). Vérifie : données (total, en cours, en instance, top
catégories, top demandeurs, détail par mois), droits par rôle, PDF / PPTX / Excel, aperçu et .eml,
vocabulaire « en instance », aucune mention « SANAD ».
"""
import email
import io
import os
import sys
import zipfile
from email import policy

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_pchc_export import ADMIN, Server, _pdf_pages, _pdf_text  # noqa: E402

DATA = os.environ.get("SANAD_MOUSSANADA_DATA") or os.path.join(ROOT, "data", "moussanada")
PWD = "Passw0rd#2026"


def _sanad(raw):
    """Mention de la plateforme (le nom du volet « Moussanada » contient « sanad » : écarté)."""
    s = raw.lower() if isinstance(raw, bytes) else raw.lower().encode()
    return b"sanad" in s.replace(b"moussanada", b"")


def _no_sanad_zip(data, label):
    z = zipfile.ZipFile(io.BytesIO(data))
    for n in z.namelist():
        assert not _sanad(z.read(n)), f"mention SANAD dans {label}:{n}"


def test_annual_report_end_to_end():
    start, end = "2026-01", "2026-12"
    with Server(seed={"moussanada": DATA}) as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        admin.post(U + "/api/moussanada/settings", json={"contacts": [{"email": "direction@ammps.gov.ma", "name": "Direction", "role": "to"}]})
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            assert admin.post(U + "/api/users", json={"username": name, "password": PWD, "role": role, "name": name, "modules": ["moussanada"]}).status_code == 200
        sup = srv.session({"username": "sup", "password": PWD})
        hot = srv.session({"username": "hot", "password": PWD})

        view = admin.get(U + f"/api/moussanada/annual/2026?start={start}&end={end}").json()
        ms, t = view["months"], view["totals"]
        n = len(ms)
        assert n >= 1
        assert t["total"] == sum(m["total"] for m in ms) == sum(m["ouverts"] for m in ms)
        assert t["en_instance"] == sum(m["en_retard"] for m in ms)
        assert all(m["en_cours"] >= 0 for m in ms) and t["en_cours"] == sum(m["en_cours"] for m in ms)
        assert 1 <= len(view["top_categories"]) <= 10 and 1 <= len(view["top_demandeurs"]) <= 10
        tc = [c["ouverts"] for c in view["top_categories"]]
        assert tc == sorted(tc, reverse=True)
        assert 1 <= len(view["top_services"]) <= 10
        assert view["period_label"] == f"{ms[0]['label']} — {ms[-1]['label']}", view["period_label"]   # 1er et dernier mois avec données
        print(f"  {n} mois, total {t['total']}, en cours {t['en_cours']}, en instance {t['en_instance']}")

        # Droits : hotliner interdit ; superviseur exporte mais n'envoie pas
        assert hot.get(U + f"/api/moussanada/annual-report/pdf?start={start}&end={end}").status_code == 403
        assert sup.post(U + "/api/moussanada/annual-send-email", json={"start": start, "end": end}).status_code == 403
        assert admin.get(U + "/api/moussanada/annual-report/pdf?start=2026-13&end=2026-01").status_code == 400
        assert admin.get(U + f"/api/moussanada/annual-report/doc?start={start}&end={end}").status_code == 404

        top_cat = view["top_categories"][0]["label"].split(">")[-1].strip()
        top_dem = view["top_demandeurs"][0]["label"]

        r = sup.get(U + f"/api/moussanada/annual-report/pdf?start={start}&end={end}")
        assert r.status_code == 200 and r.content[:5] == b"%PDF-", r.text[:300]
        text = _pdf_text(r.content)
        if text is not None:
            flat = " ".join(text.split())
            assert not _sanad(flat) and "en retard" not in flat.lower() and "échéance dépassée" not in flat
            for needle in ("Total tickets", "En cours", "En instance", "Détail par mois", "Top des demandeurs", "Top des services demandeurs", ms[0]["label"], ms[-1]["label"], "Catégories les plus demandées", top_dem):
                assert needle in flat, needle
        print(f"  PDF {len(r.content) / 1e6:.2f} Mo, {_pdf_pages(r.content)} pages")

        r = sup.get(U + f"/api/moussanada/annual-report/pptx?start={start}&end={end}")
        assert r.status_code == 200
        from pptx import Presentation
        prs = Presentation(io.BytesIO(r.content))
        assert len(prs.slides) == 6, len(prs.slides)
        ptxt = " ".join(p.text for sl in prs.slides for sh in sl.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs)
        ptxt += " ".join(c.text for sl in prs.slides for sh in sl.shapes if sh.has_table for row in sh.table.rows for c in row.cells)
        for needle in ("En instance", "Total tickets", "Top des demandeurs", "Détail par mois", top_dem):   # cartes KPI en majuscules
            assert needle.lower() in ptxt.lower(), needle
        for slide in prs.slides:
            for sh in slide.shapes:
                assert sh.left >= 0 and sh.top >= 0 and sh.left + sh.width <= prs.slide_width + 10 and sh.top + sh.height <= prs.slide_height + 10, sh.name
        _no_sanad_zip(r.content, "pptx")
        pptx_slides = len(prs.slides)

        r = sup.get(U + f"/api/moussanada/annual-report/xlsx?start={start}&end={end}")
        assert r.status_code == 200
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(r.content))
        assert wb.sheetnames == ["Synthèse", "Catégories", "Demandeurs", "Services"], wb.sheetnames
        cells = [c.value for row in wb["Synthèse"].iter_rows() for c in row]
        assert "En instance" in cells and "Total tickets" in cells and ms[0]["label"] in cells
        assert wb["Demandeurs"]["A4"].value == top_dem
        _no_sanad_zip(r.content, "xlsx")

        pv = admin.get(U + f"/api/moussanada/annual-email-preview?start={start}&end={end}&greeting=Bonjour%20Direction%2C").json()
        html = pv["html"]
        for needle in ("Total tickets", "En cours", "En instance", "Tendance mensuelle", "Catégories les plus demandées", "Top des demandeurs",
                       "Détail par mois", "Bonjour Direction,", ms[0]["label"], top_dem, top_cat):
            assert needle in html, needle
        assert "échéance dépassée" not in html and ms[-1]["label"] in html and "Top des services demandeurs" in html
        assert "en retard" not in html.lower() and not _sanad(html) and not _sanad(pv["subject"])
        assert html.count("data:image/png") >= 4, "graphiques (tendance, instance, catégories, demandeurs) attendus"

        r = admin.post(U + "/api/moussanada/annual-send-email", json={"start": start, "end": end, "greeting": "Bonjour Direction,",
                                                                      "attachments": {"pdf": True, "pptx": True, "xlsx": True}})
        d = r.json()
        base = f"Rapport_Moussanada_{start}_{end}"
        assert r.status_code == 200 and d["ok"] and d["attachments"] == [base + ".pdf", base + ".pptx", base + ".xlsx"], d
        raw = admin.get(U + d["download"]).content
        msg = email.message_from_bytes(raw, policy=policy.default)
        # corps décodé (le brut coupe « Moussanada » en quoted-printable) + chaque pièce jointe
        assert not _sanad(msg.get_body(("html",)).get_content()) and not _sanad(msg["Subject"])
        for part in msg.iter_attachments():
            if part.get_filename().endswith((".pptx", ".xlsx")):
                _no_sanad_zip(part.get_content(), part.get_filename())
        assert msg["To"] == "direction@ammps.gov.ma" and msg["X-Unsent"] == "1"
        atts = {p.get_filename(): p.get_content() for p in msg.iter_attachments()}
        assert len(Presentation(io.BytesIO(atts[base + ".pptx"])).slides) == pptx_slides
        assert openpyxl.load_workbook(io.BytesIO(atts[base + ".xlsx"])).sheetnames == wb.sheetnames
        assert atts[base + ".pdf"][:5] == b"%PDF-"
        log = admin.get(U + "/api/moussanada/send-log").json()
        assert log[0]["ym"] == f"{start}_{end}" and len(log[0]["attachments"]) == 3
        audit = srv.audit_text()
        assert audit.count('"annual_export"') >= 3 and audit.count('"annual_send_email"') == 1

        # Fichier annuel importé : alternative aux données mois par mois
        assert admin.get(U + "/api/moussanada/annual-import").json() == {"available": False}
        assert admin.get(U + f"/api/moussanada/annual/2026?start={start}&end={end}&source=import").status_code == 409
        wbi = openpyxl.Workbook()
        w1 = wbi.active
        w1.title = "rawData - Tickets"
        w1.append(["Mois", "Ouverts", "Résolus", "En retard", "Clos"])
        w1.append(["2026-01", 100, 90, 7, 80])
        w1.append(["2026-02", 120, 110, 9, 100])
        w2 = wbi.create_sheet("rawData - Ticket - Catégories")
        w2.append(["Catégorie", "Ouverts", "Résolus", "En retard", "Fermés"])
        w2.append(["Catégorie IMPORTEE A", 150, 140, 3, 130])
        w2.append(["Catégorie IMPORTEE B", 70, 60, 1, 50])
        w3 = wbi.create_sheet("rawData - Ticket - Demandeur")
        w3.append(["Demandeur", "Ouverts", "Résolus", "En retard", "Fermés"])
        w3.append(["DEMANDEUR IMPORTE", 99, 90, 2, 80])
        w4 = wbi.create_sheet("rawData - Ticket - Services")
        w4.append(["Service", "Ouverts", "Résolus", "En retard", "Fermés"])
        w4.append(["SERVICE IMPORTE", 88, 80, 1, 70])
        buf = io.BytesIO()
        wbi.save(buf)
        files = lambda: {"file": ("annuel_2026.xlsx", buf.getvalue())}
        assert sup.post(U + "/api/moussanada/annual-import", files=files()).status_code == 403
        assert admin.post(U + "/api/moussanada/annual-import", files={"file": ("x.csv", b"a,b")}).status_code == 400
        r = admin.post(U + "/api/moussanada/annual-import", files=files())
        assert r.status_code == 200 and r.json()["info"]["months"] == ["2026-01", "2026-02"], r.text
        assert sup.get(U + "/api/moussanada/annual-import").json()["filename"] == "annuel_2026.xlsx"
        imp = admin.get(U + f"/api/moussanada/annual/2026?start={start}&end={end}&source=import").json()
        assert imp["source"] == "import" and imp["totals"]["total"] == 220 and imp["totals"]["en_instance"] == 16
        assert [c["label"] for c in imp["top_categories"]] == ["Catégorie IMPORTEE A", "Catégorie IMPORTEE B"]
        assert imp["top_demandeurs"][0]["label"] == "DEMANDEUR IMPORTE" and imp["top_services"][0]["label"] == "SERVICE IMPORTE"
        assert imp["period_label"] == "Janvier 2026 — Février 2026", imp["period_label"]
        # le choix « mois par mois » reste inchangé (défaut)
        assert admin.get(U + f"/api/moussanada/annual/2026?start={start}&end={end}").json()["totals"] == t
        # livrables selon la source
        q = f"start={start}&end={end}&source=import"
        px = Presentation(io.BytesIO(sup.get(U + f"/api/moussanada/annual-report/pptx?{q}").content))
        ptxt = " ".join(c.text for sl in px.slides for sh in sl.shapes if sh.has_table for row in sh.table.rows for c in row.cells)
        assert "DEMANDEUR IMPORTE" in ptxt and "SERVICE IMPORTE" in ptxt and top_dem not in ptxt
        wx = openpyxl.load_workbook(io.BytesIO(sup.get(U + f"/api/moussanada/annual-report/xlsx?{q}").content))
        assert wx["Demandeurs"]["A4"].value == "DEMANDEUR IMPORTE" and "annuel_2026.xlsx" in wx["Synthèse"]["A3"].value
        text = _pdf_text(sup.get(U + f"/api/moussanada/annual-report/pdf?{q}").content)
        assert text is None or "DEMANDEUR IMPORTE" in " ".join(text.split())
        h = admin.get(U + f"/api/moussanada/annual-email-preview?{q}").json()["html"]
        assert "DEMANDEUR IMPORTE" in h and "annuel_2026.xlsx" in h and top_dem not in h
        d = admin.post(U + "/api/moussanada/annual-send-email", json={"start": start, "end": end, "source": "import", "attachments": {"xlsx": True}}).json()
        msg = email.message_from_bytes(admin.get(U + d["download"]).content, policy=policy.default)
        (att,) = list(msg.iter_attachments())
        assert openpyxl.load_workbook(io.BytesIO(att.get_content()))["Demandeurs"]["A4"].value == "DEMANDEUR IMPORTE"
        assert admin.post(U + "/api/moussanada/annual-import/clear").json()["ok"]
        assert admin.get(U + "/api/moussanada/annual-import").json() == {"available": False}
        assert '"annual_import"' in srv.audit_text()

        # Non-régression : e-mail mensuel Moussanada
        ym = ms[-1]["ym"]
        r = admin.post(U + f"/api/moussanada/preview-email/{ym}", json={})
        assert r.status_code == 200 and not _sanad(r.json()["html"])


if __name__ == "__main__":
    test_annual_report_end_to_end()
    print("OK — rapport de période Moussanada")
