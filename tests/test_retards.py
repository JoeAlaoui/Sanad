# -*- coding: utf-8 -*-
"""Tickets en retard Moussanada : calcul sur le CSV de référence (dates JJ/MM/AAAA), exclusion de techniciens
dans le TOP de l'email, Excel joint (tous les retards) rempli depuis le gabarit, droits par rôle.
Lancer : python tests/test_retards.py
"""
import email
import io
import os
import re
import sys
from datetime import datetime
from email import policy

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook  # noqa: E402

import retards  # noqa: E402
import retards_xlsx as X  # noqa: E402
from glpi_import import parse_glpi_raw_tickets  # noqa: E402
from test_pchc_export import ADMIN, Server  # noqa: E402
from test_tarkhiss_annual import PWD  # noqa: E402

CSV = os.path.join(ROOT, "reference", "glpi (1).csv")
NOW = datetime(2026, 11, 2, 18, 0)
EXCLUDED = "Alaoui Youssef"   # porte les plus gros retards : doit disparaître du TOP quand il est exclu


class _Upload:
    def __init__(self, path):
        self.f = open(path, "rb")

    def read(self):
        return self.f.read()

    filename = "glpi.csv"


def _tickets():
    return parse_glpi_raw_tickets(_Upload(CSV))


def test_calcul():
    tickets = _tickets()
    assert len(tickets) == 55
    first = tickets[0]
    assert first["id"] == "1784" and first["date_ouverture"] == "2026-11-02 08:44" and first["ttr"] == "2026-11-02 14:44", first
    assert "&#039;" not in "".join(t["titre"] for t in tickets) and not any(t["titre"].startswith('"') for t in tickets)

    v = retards.view(tickets, None, NOW)
    assert v["open_total"] == 55 and v["total"] == 44 and v["no_due"] == 11, (v["open_total"], v["total"], v["no_due"])
    hours = [t["retard_h"] for t in v["late"]]
    assert hours == sorted(hours, reverse=True) and v["late"][0]["id"] == "1369"
    assert len(v["top"]) == 10 and v["hidden"] == 0

    cfg = {"late_tickets": {"top_n": 5, "excluded_technicians": [EXCLUDED.upper()]}}   # insensible à la casse
    v2 = retards.view(tickets, cfg, NOW)
    assert v2["total"] == 44 and len(v2["top"]) == 5 and v2["hidden"] == 6
    assert all(t["technicien"] != EXCLUDED for t in v2["top"]) and v2["late"] == v["late"]
    # un ticket résolu ou clos n'est jamais en retard
    done = [dict(t, statut="Résolu") for t in tickets]
    assert retards.view(done, None, NOW)["total"] == 0
    print(f"  calcul : {v['total']} retards sur {v['open_total']} ouverts, {v['no_due']} sans échéance")


def _check_workbook(data, total):
    wb = load_workbook(io.BytesIO(data))
    assert wb.sheetnames == [X.SHEET_T, X.SHEET_S]
    ws = wb[X.SHEET_T]
    rows = [r for r in ws.iter_rows(min_row=X.FIRST_ROW, min_col=2, max_col=11, values_only=True) if r[0]]
    assert len(rows) == total, (len(rows), total)
    assert [c.value for c in ws[X.HEADER_ROW]][1:11] == X.T_HEAD
    assert ws[X.KPI_VALUE_CELLS[0]].value == total and isinstance(rows[0][7], datetime) and rows[0][9] > 0
    assert len(ws._images) == 1 and ws.conditional_formatting, "logo ou mises en forme conditionnelles perdus"
    assert not re.search(r"\b(sanad|glpi)\b", " ".join(str(c) for r in rows for c in r), re.I)
    ss = wb[X.SHEET_S]
    tot = next(r for r in ss.iter_rows(min_row=X.S_FIRST_ROW, min_col=2, max_col=7, values_only=True) if r[0] == X.FOOTER_MODEL)
    assert tot[1] == total
    return rows


def test_excel_gabarit():
    v = retards.view(_tickets(), None, NOW)
    rows = _check_workbook(X.fill(v), 44)
    assert rows[0][0] == "1369" and rows[0][4] == "Alaoui Youssef"
    # technicien écarté : absent de l'Excel (tickets, KPI, synthèse)
    cfg = {"late_tickets": {"top_n": 5, "excluded_technicians": [EXCLUDED]}}
    rows = _check_workbook(X.fill(retards.view(_tickets(), cfg, NOW)), 38)
    assert not any(r[4] == EXCLUDED for r in rows) and rows[0][0] == "1111"
    # peu de tickets, puis aucun : le gabarit se remplit sans reste des lignes d'exemple
    few = retards.view([t for t in _tickets() if t["id"] in ("1369", "1352")], None, NOW)
    assert len(_check_workbook(X.fill(few), 2)) == 2
    empty = load_workbook(io.BytesIO(X.fill(retards.view([], None, NOW))))[X.SHEET_T]
    assert empty.cell(X.FIRST_ROW, 3).value == "Aucun ticket en retard." and empty.cell(X.FIRST_ROW + 1, 2).value is None
    print("  Excel gabarit : 44 / 2 / 0 tickets OK")


def _top_names(html):
    """Techniciens (3e colonne) du tableau TOP de la section « Tickets en retard » de l'email."""
    html = re.sub(r'src="data:[^"]+"', 'src=""', html)
    sec = html[html.index("Tickets en retard"):]
    sec = sec[:sec.index("Cordialement")]
    return sec, re.findall(r'<td align="left"[^>]*>([^<]*)</td>', sec)


def test_serveur():
    with Server() as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            assert admin.post(U + "/api/users", json={"username": name, "password": PWD, "role": role, "name": name, "modules": ["tarkhiss", "moussanada"]}).status_code == 200
        sup, hot = srv.session({"username": "sup", "password": PWD}), srv.session({"username": "hot", "password": PWD})

        # avant tout import : pas de section, pas de pièce jointe
        d = admin.post(U + "/api/moussanada/preview-email/2026-10", json={}).json()
        assert "Tickets en retard" not in d["html"] and d["attachments"] == []

        with open(CSV, "rb") as f:
            r = admin.post(U + "/api/moussanada/import-tickets", files={"file": ("glpi.csv", f)})
        assert r.status_code == 200 and r.json()["result"]["rows"] == 55, r.text

        # l'horloge réelle compte : les retards évoluent avec l'heure (autant d'attendus que de tickets du jour)
        cfg = {"late_tickets": {"top_n": 5, "excluded_technicians": [EXCLUDED]}}
        ref = retards.view(_tickets(), None, datetime.now())
        expect, hidden = ref["total"], retards.view(_tickets(), cfg, datetime.now())["hidden"]
        assert expect >= 30 and hidden >= 5, (expect, hidden)
        api = admin.get(U + "/api/moussanada/late-tickets").json()
        assert api["total"] == expect and len(api["top"]) == 10 and len(api["all"]) == expect
        assert EXCLUDED in {t["name"] for t in api["technicians"]}

        # email : section, TOP, pièce jointe
        d = admin.post(U + "/api/moussanada/preview-email/2026-10", json={}).json()
        sec, names = _top_names(d["html"])
        assert "TOP 10" in sec and EXCLUDED in names and "fichier Excel joint" in sec
        assert d["attachments"] and d["attachments"][0]["name"].startswith("Tickets_en_retard_Moussanada_")
        assert not re.search(r"\b(SANAD|GLPI)\b", re.sub(r'src="data:[^"]+"', "", d["html"]), re.I)

        # exclusion : disparaît du TOP, mais pas du total ni de l'Excel
        assert admin.post(U + "/api/moussanada/settings", json={"late_tickets": {"top_n": 5, "excluded_technicians": [EXCLUDED]}}).status_code == 200
        d = admin.post(U + "/api/moussanada/preview-email/2026-10", json={}).json()
        sec, names = _top_names(d["html"])
        assert "TOP 5" in sec and EXCLUDED not in sec and len([n for n in names if n]) >= 5
        assert f">{expect - hidden}</div>" in sec and f"complète des {expect - hidden} tickets" in sec, "KPI de l'email : techniciens écartés exclus"
        api = admin.get(U + "/api/moussanada/late-tickets").json()
        assert api["hidden"] == hidden and api["total"] == expect and len(api["top"]) == 5
        assert any(t["excluded"] for t in api["all"]) and not any(t["technicien"] == EXCLUDED for t in api["top"])

        # .eml : Excel joint sans les techniciens écartés (ni leurs tickets, ni leur ligne de synthèse)
        r = admin.post(U + "/api/moussanada/send-email/2026-10", json={"to": ["a@ammps.test"]})
        assert r.status_code == 200 and r.json()["attachments"], r.text
        eml = email.message_from_bytes(open(r.json()["file"], "rb").read(), policy=policy.default)
        atts = [p for p in eml.iter_attachments() if p.get_filename().endswith(".xlsx")]
        assert len(atts) == 1
        data = atts[0].get_payload(decode=True)
        rows = _check_workbook(data, expect - hidden)
        assert not any(r[4] == EXCLUDED for r in rows)
        assert EXCLUDED not in {r[0] for r in load_workbook(io.BytesIO(data))[X.SHEET_S].iter_rows(min_row=X.S_FIRST_ROW, min_col=2, max_col=2, values_only=True)}

        # bouton « Excel » du dashboard = même gabarit, mêmes exclusions
        r = admin.get(U + "/api/moussanada/late-tickets.xlsx")
        assert r.status_code == 200 and len(_check_workbook(r.content, expect - hidden)) == expect - hidden

        # sections désactivées
        es = {"late_xlsx": False}
        assert admin.post(U + "/api/moussanada/settings", json={"email_sections": es}).status_code == 200
        d = admin.post(U + "/api/moussanada/preview-email/2026-10", json={}).json()
        assert d["attachments"] == [] and "fichier Excel joint" not in d["html"]
        assert admin.post(U + "/api/moussanada/settings", json={"email_sections": {"late": False}}).status_code == 200
        d = admin.post(U + "/api/moussanada/preview-email/2026-10", json={}).json()
        assert "Tickets en retard" not in d["html"] and d["attachments"] == []

        # droits : superviseur lecture seule, hotliner exclu
        assert sup.get(U + "/api/moussanada/late-tickets").status_code == 200
        assert sup.get(U + "/api/moussanada/late-tickets.xlsx").status_code == 200
        assert sup.post(U + "/api/moussanada/settings", json={"late_tickets": {"top_n": 20}}).status_code == 403
        assert hot.get(U + "/api/moussanada/late-tickets").status_code == 403
        srv.log.flush()
        log = open(os.path.join(srv.dir, "server.log"), encoding="utf-8", errors="replace").read()
        assert "Traceback" not in log and "Email charte" not in log, log[-1500:]
    print("  serveur : API, email, .eml, droits OK")


if __name__ == "__main__":
    test_calcul()
    test_excel_gabarit()
    test_serveur()
    print("OK")
