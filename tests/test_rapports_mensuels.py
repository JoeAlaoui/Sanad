# -*- coding: utf-8 -*-
"""Rapports mensuels selon la charte AMMPS (Tarkhiss, Moussanada) sur un serveur isolé :
routes PDF / PPTX / aperçu email, droits par rôle, aucune mention SANAD ni GLPI, mois vide.
Lancer : python tests/test_rapports_mensuels.py
"""
import io
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_pchc_export import ADMIN, Server, _pdf_pages  # noqa: E402
from test_tarkhiss_annual import PWD  # noqa: E402

CASES = (("tarkhiss", "2026-10"), ("moussanada", "2026-09"), ("tarkhiss", "2020-01"), ("moussanada", "2020-01"))  # 2020-01 = mois vide


def test_monthly_reports():
    from pptx import Presentation
    seed = {m: os.path.join(ROOT, "data", m) for m in ("tarkhiss", "moussanada")}
    with Server(seed=seed) as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            assert admin.post(U + "/api/users", json={"username": name, "password": PWD, "role": role, "name": name, "modules": ["tarkhiss", "moussanada"]}).status_code == 200
        sup = srv.session({"username": "sup", "password": PWD})
        hot = srv.session({"username": "hot", "password": PWD})

        for module, ym in CASES:
            r = sup.get(U + f"/api/{module}/export-pptx/{ym}")
            assert r.status_code == 200 and r.content[:2] == b"PK", r.text[:300]
            deck = Presentation(io.BytesIO(r.content))
            assert len(deck.slides) >= 7
            z = zipfile.ZipFile(io.BytesIO(r.content))
            for n in z.namelist():
                if n.endswith(".xml"):   # « Moussanada » contient les lettres « sanad » : recherche du mot entier
                    assert not re.search(rb"\b(sanad|glpi)\b", z.read(n), re.I), f"mention interne dans {module}:{n}"
            assert hot.get(U + f"/api/{module}/export-pptx/{ym}").status_code == 403

            r = admin.post(U + f"/api/{module}/export-pdf/{ym}", json={})
            assert r.status_code == 200 and r.content[:5] == b"%PDF-", r.text[:300]
            print(f"  {module} {ym} : PPTX {len(deck.slides)} diapos, PDF {_pdf_pages(r.content)} pages")

            d = admin.post(U + f"/api/{module}/preview-email/{ym}", json={"greeting": "Bonjour Direction,"}).json()
            html = re.sub(r'src="data:[^"]+"', 'src=""', d["html"])
            assert not re.search(r"\b(SANAD|GLPI)\b", html, re.I)
            assert html.index("Bonjour Direction,") < html.index('width="720"') < html.index("Cordialement"), "salutation/signature hors gabarit"
            assert "Rue Lamfadal" not in html, "pied de page agence retiré des emails"

        assert admin.get(U + "/api/pchc/export-pptx/2026-10").status_code == 404
        srv.log.flush()
        log = open(os.path.join(srv.dir, "server.log"), encoding="utf-8", errors="replace").read()
        assert "PDF charte" not in log and "Email charte" not in log, "repli sur l'ancien rapport : " + log[-1500:]
    print("OK")


if __name__ == "__main__":
    test_monthly_reports()
