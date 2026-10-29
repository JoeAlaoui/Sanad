# -*- coding: utf-8 -*-
"""Fabrique le gabarit Excel du Reporting Métier PCHC à partir du classeur de référence validé.

    python tools/build_pchc_template.py "reference/Rapport Tarkhiss_PCHC_Metier_YTD.xlsx"

Le gabarit (pchc_template/dashboard_gabarit.xlsx) est la référence elle-même, à l'identique
(graphiques, styles, couleurs, mises en forme conditionnelles, mise en page), sauf :
- feuilles brutes vidées (<sheetData/>) : elles sont réécrites à chaque génération ;
- table des chaînes partagées réduite aux chaînes des feuilles Dashboard/KPI.
- ouverture sur le Dashboard (onglet actif et sélectionné), au lieu du dernier onglet consulté dans
  la référence ; c'est le seul écart volontaire avec elle.
Les styles des feuilles brutes (par colonne, ligne 6) sont relevés dans dashboard_gabarit.json.
À relancer uniquement si la référence change.
"""
import io
import json
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "pchc_template")
RAW = ["identificationOpérateur", "déclarationActivité", "certificatEnregistrement", "certificatLibreVente",
       "autorisationImportationMatièreP"]


def main(src):
    z = zipfile.ZipFile(src)
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = {}
    for tag in re.findall(r"<Relationship [^>]*>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', tag))
        rid_target[a["Id"]] = a["Target"]
    sheets = []
    for tag in re.findall(r"<sheet [^>]*>", wbx):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        sheets.append((a["name"], "xl/" + rid_target[a["r:id"]]))
    names = [s[0] for s in sheets]
    assert names[:6] == ["Dashboard", "KPI Identification", "KPI Déclaration", "KPI Certificat Enreg.", "KPI Certificat LV",
                         "KPI Autor. Import."] and names[6:] == RAW, names
    meta = {"source": os.path.basename(src), "sheets": {n: p for n, p in sheets}, "raw": {}}

    # Chaînes partagées utilisées par les 6 feuilles de restitution : elles doivent être en tête
    used = set()
    for n, p in sheets[:6]:
        used |= {int(v) for v in re.findall(r't="s"[^>]*><v>(\d+)</v>', z.read(p).decode("utf-8"))}
    keep = max(used) + 1
    assert used == set(range(keep)), "chaînes partagées non contiguës : renumérotation nécessaire"
    sst = z.read("xl/sharedStrings.xml").decode("utf-8")
    items = re.findall(r"<si>.*?</si>", sst, re.S)
    head = sst[:sst.index("<sst")]
    sst_tag = re.search(r"<sst[^>]*>", sst).group(0)
    sst_tag = re.sub(r' count="\d+"', ' count="0"', re.sub(r' uniqueCount="\d+"', f' uniqueCount="{keep}"', sst_tag))
    new_sst = head + sst_tag + "".join(items[:keep]) + "</sst>"

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            name = info.filename
            data = z.read(name)
            if name == "xl/sharedStrings.xml":
                data = new_sst.encode("utf-8")
            elif name == "xl/workbook.xml":
                x = re.sub(r' (firstSheet|activeTab)="\d+"', "", data.decode("utf-8"))
                data = x.encode("utf-8")
            elif name in [p for _, p in sheets[:6]]:
                x = re.sub(r' tabSelected="1"', "", data.decode("utf-8"))
                if name == sheets[0][1]:
                    x = x.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
                data = x.encode("utf-8")
            elif name in [p for _, p in sheets[6:]]:
                x = data.decode("utf-8")
                raw_name = next(n for n, p in sheets if p == name)
                row6 = re.search(r'<row r="6"[^>]*>(.*?)</row>', x, re.S)
                styles = {}
                if row6:
                    for ref, attrs in re.findall(r'<c r="([A-Z]+)6"([^>]*)', row6.group(1)):
                        m = re.search(r' s="(\d+)"', attrs)
                        styles[ref] = m.group(1) if m else None
                row_tag = re.search(r'<row r="6"([^>]*)>', x)
                meta["raw"][raw_name] = {"styles": styles,
                                         "row_attrs": re.sub(r' (r|spans)="[^"]*"', "", row_tag.group(1)) if row_tag else ""}
                x = re.sub(r"<sheetData>.*</sheetData>", "<sheetData/>", x, flags=re.S)
                data = x.encode("utf-8")
            zo.writestr(info, data)
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, "dashboard_gabarit.xlsx"), "wb") as f:
        f.write(out.getvalue())
    with open(os.path.join(OUT_DIR, "dashboard_gabarit.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print(f"Gabarit écrit : {len(out.getvalue()) / 1024:.0f} Ko, {keep} chaînes conservées")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "reference", "Rapport Tarkhiss_PCHC_Metier_YTD.xlsx"))
