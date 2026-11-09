"""SANAD — fonctions d'exploitation : qualité des données, sauvegarde/restauration, SOP & preuves d'audit.

Branché depuis app.py via `ops.register(app, sys.modules[__name__])` (H = module app, pour réutiliser
ses helpers : stockage JSON, journal d'audit, génération PDF…). Toutes les routes /api/admin/* sont
réservées à l'administrateur (voir auth_gate) ; /api/quality/* est lisible par admin/superviseur.
"""
import csv
import hashlib
import io
import json
import os
import re
import threading
import time
import zipfile
from datetime import date, datetime, timedelta
from html import escape as esc

import store
from flask import jsonify, request, send_file

SNAP_RE = re.compile(r"^snapshot_\d{8}_\d{6}_[a-z-]+\.zip$")
EXCLUDED_DIRS = {"_snapshots", "_backups", "exports"}
YM_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")


def register(app, H):
    DATA_DIR = H.DATA_DIR
    SNAP_DIR = os.path.join(DATA_DIR, "_snapshots")
    os.makedirs(SNAP_DIR, exist_ok=True)

    # ------------------------------------------------------------------ helpers
    def gs():
        return H.load_global_settings()

    def holidays():
        raw = gs().get("quality_holidays") or ""
        return {x.strip() for x in re.split(r"[,;\s]+", raw) if re.match(r"^\d{4}-\d{2}-\d{2}$", x.strip())}

    def business_days(ym):
        y, m = map(int, ym.split("-"))
        today = date.today()
        out, hol = [], holidays()
        d = date(y, m, 1)
        while d.month == m:
            iso = d.isoformat()
            if d.weekday() < 5 and iso not in hol and d <= today:
                out.append(iso)
            d += timedelta(days=1)
        return out

    def month_list(first, last):
        return H.month_range(first, last)

    def now_ym():
        return datetime.now().strftime("%Y-%m")

    def tarkhiss_months():
        d = H.module_dir("tarkhiss")
        out = set()
        for fn in store.names(d):
            mm = re.match(r"^(calls|emails|analysis)_(\d{4}-\d{2})\.json$", fn)
            if mm:
                out.add(mm.group(2))
        return sorted(out)

    def moussanada_months():
        out = {ym for ym, row in H.load_timeseries().items() if any((row.get(k) or 0) for k in ("ouverts", "resolus", "clos", "en_retard"))}
        d = H.module_dir("moussanada")
        for fn in store.names(d):
            mm = re.match(r"^(sources|notes)_(\d{4}-\d{2})\.json$", fn)
            if mm:
                out.add(mm.group(2))
        return sorted(out)

    def was_sent(module, ym):
        return any(e.get("ym") == ym for e in H.load_send_log(module))

    def check(key, label, ok, detail="", weight=1):
        return {"key": key, "label": label, "ok": bool(ok), "detail": detail, "weight": weight}

    def score(checks):
        tot = sum(c["weight"] for c in checks)
        return round(100 * sum(c["weight"] for c in checks if c["ok"]) / tot) if tot else None

    def tarkhiss_quality(ym):
        days = business_days(ym)
        calls = H._load("tarkhiss", "calls", ym, {}) or {}
        emails = H._load("tarkhiss", "emails", ym, {}) or {}
        an = H._load("tarkhiss", "analysis", ym, None) or {}
        miss_c = [d for d in days if (calls.get(d) or {}).get("calls") in (None, "")]
        miss_e = [d for d in days if (emails.get(d) or {}).get("received") in (None, "")]
        closed = ym < now_ym()
        has_an = bool(an.get("problems") or an.get("demandes"))
        checks = [
            check("calls", "Appels saisis (jours ouvrés)", not miss_c, f"{len(days) - len(miss_c)}/{len(days)} jours", 3),
            check("emails", "Emails saisis (jours ouvrés)", not miss_e, f"{len(days) - len(miss_e)}/{len(days)} jours", 3),
            check("analysis", "Analyse mensuelle renseignée", has_an, "problèmes/demandes présents" if has_an else "aucune catégorie", 2),
            check("synth", "Constats et recommandations", bool(an.get("constats")) and bool(an.get("recommandations")), "", 1),
        ]
        if closed:
            checks.append(check("sent", "Rapport généré/envoyé", was_sent("tarkhiss", ym), "voir historique d'envoi", 2))
        return {"ym": ym, "closed": closed, "score": score(checks), "checks": checks,
                "missing_calls": miss_c, "missing_emails": miss_e}

    def moussanada_quality(ym):
        ts = H.load_timeseries().get(ym) or {}
        src = H.load_sources(ym)
        notes = H.load_notes(ym)
        sections = notes.get("sections", {})
        filled = sum(1 for v in sections.values() if (v or "").strip())
        closed = ym < now_ym()
        has_ts = any((ts.get(k) or 0) for k in ("ouverts", "resolus", "clos", "en_retard"))
        checks = [
            check("tickets", "Tickets importés (volumes du mois)", has_ts, "" if has_ts else "aucun volume", 3),
            check("categories", "Catégories importées", bool(src.get("categories")), "", 2),
            check("services", "Services importés", bool(src.get("services")), "", 2),
            check("techniciens", "Techniciens importés", bool(src.get("techniciens")), "", 2),
            check("notes", "Analyse rédigée (7 sections)", filled >= len(sections) and filled > 0, f"{filled}/{len(sections)} sections", 2),
        ]
        if closed:
            checks.append(check("sent", "Rapport généré/envoyé", was_sent("moussanada", ym), "voir historique d'envoi", 2))
        return {"ym": ym, "closed": closed, "score": score(checks), "checks": checks}

    def last_sync(module):
        latest = store.latest_mtime(H.module_dir(module))
        return datetime.fromtimestamp(latest).isoformat(timespec="seconds") if latest else None

    def overview(n=12):
        cur = now_ym()
        res = {"generated": datetime.now().isoformat(timespec="seconds"), "current": cur, "modules": {}, "alerts": []}
        for module, months_fn, qfn, label in (("tarkhiss", tarkhiss_months, tarkhiss_quality, "Tarkhiss"),
                                              ("moussanada", moussanada_months, moussanada_quality, "Moussanada")):
            ms = months_fn()
            rows = []
            if ms:
                span = month_list(ms[0], max(ms[-1], cur))[-n:]
                rows = [qfn(ym) for ym in span]
            ls = last_sync(module)
            res["modules"][module] = {"label": label, "months": rows, "last_sync": ls}
            # alertes : mois en cours + précédent, et fraîcheur
            for r in rows[-2:]:
                for c in r["checks"]:
                    if not c["ok"] and (r["closed"] or c["key"] in ("calls", "emails", "tickets")):
                        res["alerts"].append({"module": module, "ym": r["ym"], "label": c["label"], "detail": c["detail"],
                                              "level": "danger" if r["closed"] else "warning"})
            if ls and (datetime.now() - datetime.fromisoformat(ls)).days >= 7:
                res["alerts"].append({"module": module, "ym": cur, "label": "Aucune donnée écrite depuis 7 jours ou plus", "detail": ls[:10], "level": "warning"})
            if not ls:
                res["alerts"].append({"module": module, "ym": cur, "label": "Aucune donnée enregistrée", "detail": "", "level": "info"})
        hist = H.load_import_history()
        res["last_import"] = hist[0] if hist else None
        snaps = list_snapshots()
        res["last_snapshot"] = snaps[0] if snaps else None
        if not snaps or (datetime.now() - datetime.fromisoformat(snaps[0]["created"])).days >= 2:
            res["alerts"].append({"module": "systeme", "ym": cur, "label": "Aucune sauvegarde récente (plus de 48 h)", "detail": "", "level": "warning"})
        return res

    @app.route("/api/quality/overview")
    def quality_overview():
        return jsonify(overview(min(int(request.args.get("months", 12)), 36)))

    # ------------------------------------------------------------------ snapshots
    def list_snapshots():
        out = []
        for fn in os.listdir(SNAP_DIR):
            if not SNAP_RE.match(fn):
                continue
            p = os.path.join(SNAP_DIR, fn)
            reason = fn[:-4].split("_", 3)[3]
            created = datetime.strptime(fn[9:24], "%Y%m%d_%H%M%S").isoformat(timespec="seconds")
            out.append({"name": fn, "size": os.path.getsize(p), "created": created, "reason": reason})
        return sorted(out, key=lambda x: x["created"], reverse=True)

    def data_files():
        for root, dirs, files in os.walk(DATA_DIR):
            dirs[:] = [d for d in dirs if d not in EXCLUDED_DIRS]
            for fn in files:
                if fn.endswith(".tmp") or fn.startswith("sanad.db"):   # la base est ajoutée ci-dessous (copie cohérente)
                    continue
                full = os.path.join(root, fn)
                yield full, "data/" + os.path.relpath(full, DATA_DIR).replace(os.sep, "/")
        if store.backend_name() == "sqlite":
            tmp = os.path.join(SNAP_DIR, f"sanad-{os.getpid()}-{threading.get_ident()}.db.tmp")
            try:
                store.snapshot_to(tmp)
                yield tmp, "data/sanad.db"
            finally:
                if os.path.exists(tmp):
                    os.remove(tmp)
        if os.path.isdir(H.UPLOAD_DIR):
            for fn in os.listdir(H.UPLOAD_DIR):
                full = os.path.join(H.UPLOAD_DIR, fn)
                if os.path.isfile(full):
                    yield full, "uploads/" + fn

    def create_snapshot(reason="manual", user=None):
        reason = re.sub(r"[^a-z-]", "", reason.lower()) or "manual"
        name = f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{reason}.zip"
        manifest = {"created": datetime.now().isoformat(timespec="seconds"), "reason": reason, "by": user or "system", "files": {}}
        path = os.path.join(SNAP_DIR, name)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for full, arc in data_files():
                with open(full, "rb") as f:
                    content = f.read()
                manifest["files"][arc] = hashlib.sha256(content).hexdigest()
                z.writestr(arc, content)
            z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=1))
        apply_retention()
        return name, len(manifest["files"])

    # Exposé à app.py (historisation de chaque import Reporting Métier PCHC)
    app.extensions["sanad_create_snapshot"] = create_snapshot

    def apply_retention():
        keep = int(gs().get("snapshot_retention") or 14)
        autos = [s for s in list_snapshots() if s["reason"] == "auto"]
        for s in autos[keep:]:
            try:
                os.remove(os.path.join(SNAP_DIR, s["name"]))
            except OSError:
                pass

    def safe_snap(name):
        return name if SNAP_RE.match(name or "") and os.path.isfile(os.path.join(SNAP_DIR, name)) else None

    @app.route("/api/admin/snapshots", methods=["GET"])
    def snapshots_list():
        g = gs()
        return jsonify({"items": list_snapshots(), "auto": g.get("snapshot_auto", True) is not False,
                        "retention": int(g.get("snapshot_retention") or 14)})

    @app.route("/api/admin/snapshots", methods=["POST"])
    def snapshots_create():
        u = H.current_user()
        name, n = create_snapshot("manual", u["username"] if u else None)
        H.log_audit("snapshot_create", {"name": name, "files": n})
        return jsonify({"ok": True, "name": name, "files": n})

    @app.route("/api/admin/snapshots/<name>/download")
    def snapshots_download(name):
        if not safe_snap(name):
            return jsonify({"error": "introuvable"}), 404
        H.log_audit("snapshot_download", {"name": name})
        return send_file(os.path.join(SNAP_DIR, name), as_attachment=True, download_name=name)

    @app.route("/api/admin/snapshots/<name>", methods=["DELETE"])
    def snapshots_delete(name):
        if not safe_snap(name):
            return jsonify({"error": "introuvable"}), 404
        os.remove(os.path.join(SNAP_DIR, name))
        H.log_audit("snapshot_delete", {"name": name})
        return jsonify({"ok": True})

    def restore_zip(zpath, user):
        with zipfile.ZipFile(zpath) as z:
            names = [n for n in z.namelist() if n != "manifest.json"]
            if not names or "manifest.json" not in z.namelist():
                raise ValueError("Archive invalide (manifest absent)")
            manifest = json.loads(z.read("manifest.json").decode("utf-8"))
            for n in names:
                if n.startswith("/") or ".." in n.split("/") or not (n.startswith("data/") or n.startswith("uploads/")):
                    raise ValueError(f"Entrée non autorisée : {n}")
            # Une archive d'un autre mode de stockage serait restaurée sans effet (ou à moitié) : refus explicite
            has_db = "data/sanad.db" in names
            has_docs = any(n.startswith("data/") and n.endswith(".json") for n in names)
            if store.backend_name() == "sqlite" and not has_db and has_docs:
                raise ValueError("Archive du mode JSON : cette instance utilise SQLite. Migrez-la d'abord avec tools/migrate_json_to_sqlite.py.")
            if store.backend_name() == "json" and has_db:
                raise ValueError("Archive du mode SQLite : cette instance utilise les fichiers JSON (SANAD_STORE=sqlite pour la restaurer).")
            contents = {}
            for n in names:  # vérification d'intégrité AVANT toute modification
                c = z.read(n)
                if manifest.get("files", {}).get(n) != hashlib.sha256(c).hexdigest():
                    raise ValueError(f"Empreinte SHA-256 invalide : {n}")
                contents[n] = c
        safety, _ = create_snapshot("pre-restore", user)  # filet de sécurité
        wanted = set(contents)
        for full, arc in list(data_files()):
            if arc not in wanted:
                os.remove(full)
        for arc, c in contents.items():
            base = DATA_DIR if arc.startswith("data/") else H.UPLOAD_DIR
            target = os.path.normpath(os.path.join(base, arc.split("/", 1)[1]))
            if not target.startswith(os.path.normpath(base)):
                raise ValueError("Chemin invalide")
            os.makedirs(os.path.dirname(target), exist_ok=True)
            if arc == "data/sanad.db":
                target = store.db_path() or target   # la base peut vivre hors de data/ (SANAD_DB_PATH)
                store.close()
                for ext in ("-wal", "-shm"):
                    if os.path.exists(target + ext):
                        os.remove(target + ext)
            with open(target, "wb") as f:
                f.write(c)
        if "data/sanad.db" in contents:
            store.reopen()
        return safety, len(contents)

    @app.route("/api/admin/snapshots/<name>/restore", methods=["POST"])
    def snapshots_restore(name):
        if not safe_snap(name):
            return jsonify({"ok": False, "error": "introuvable"}), 404
        u = H.current_user()
        try:
            safety, n = restore_zip(os.path.join(SNAP_DIR, name), u["username"] if u else None)
        except (ValueError, zipfile.BadZipFile) as e:
            return jsonify({"ok": False, "error": str(e)}), 400
        H.log_audit("snapshot_restore", {"name": name, "files": n, "safety_snapshot": safety})
        return jsonify({"ok": True, "files": n, "safety_snapshot": safety})

    @app.route("/api/admin/snapshots/upload", methods=["POST"])
    def snapshots_upload():
        f = request.files.get("file")
        if not f:
            return jsonify({"ok": False, "error": "Aucun fichier"}), 400
        tmp = os.path.join(SNAP_DIR, "_upload.tmp.zip")
        f.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as z:
                if "manifest.json" not in z.namelist():
                    raise ValueError("Archive invalide (manifest absent)")
            name = f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}_import.zip"
            os.replace(tmp, os.path.join(SNAP_DIR, name))
        except (ValueError, zipfile.BadZipFile) as e:
            os.remove(tmp)
            return jsonify({"ok": False, "error": str(e)}), 400
        H.log_audit("snapshot_upload", {"name": name})
        return jsonify({"ok": True, "name": name})

    def auto_loop():
        while True:
            try:
                g = gs()
                if g.get("snapshot_auto", True) is not False:
                    snaps = list_snapshots()
                    if not snaps or (datetime.now() - datetime.fromisoformat(snaps[0]["created"])).total_seconds() > 24 * 3600:
                        create_snapshot("auto")
            except Exception:
                pass
            time.sleep(3600)

    if os.environ.get("SANAD_NO_AUTO_SNAPSHOT") != "1":
        threading.Thread(target=auto_loop, daemon=True).start()

    # ------------------------------------------------------------------ SOP & preuves d'audit
    def sop_meta():
        g = gs()
        return {
            "ref": g.get("sop_ref") or "SOP-DSID-HELPDESK-001", "version": g.get("sop_version") or "1.0",
            "owner": g.get("sop_owner") or "PMO IT / Qualité IT", "approver": g.get("sop_approver") or "Responsable DSID",
            "date": g.get("sop_date") or datetime.now().strftime("%d/%m/%Y"),
            "agency": g.get("agency_name") or "AMMPS — DSID",
        }

    TH = "background:#0B4965;color:#fff;text-align:left;padding:6px 8px;font-size:11px;"
    TD = "border-bottom:1px solid #DAE0E7;padding:5px 8px;font-size:11px;vertical-align:top;"

    def table(headers, rows):
        h = "".join(f'<th style="{TH}">{esc(x)}</th>' for x in headers)
        b = "".join("<tr>" + "".join(f'<td style="{TD}">{c}</td>' for c in r) + "</tr>" for r in rows)
        return f'<table style="width:100%;border-collapse:collapse;margin:6px 0 14px;"><tr>{h}</tr>{b}</table>'

    def h2(t):
        return f'<h2 style="color:#0B4965;font-size:15px;border-bottom:2px solid #0B4965;padding-bottom:3px;margin:20px 0 8px;">{esc(t)}</h2>'

    def p(t):
        return f'<p style="font-size:12px;line-height:1.5;margin:4px 0;">{t}</p>'

    def sop_sections():
        """Structure SOP normalisée (objet → historique) — source unique pour HTML/PDF/DOCX."""
        return [
            ("1. Objet", ["Définir la procédure standardisée d'exploitation de la plateforme de reporting Helpdesk SI (Tarkhiss, Moussanada, PCHC) : saisie, import, contrôle, validation, diffusion et conservation des preuves."]),
            ("2. Périmètre", ["Modules : Tarkhiss (appels + emails), Moussanada (tickets GLPI), Reporting Métier PCHC.",
                              "Utilisateurs : administrateur, superviseur, hotliner. Fréquence : mensuelle (clôture au plus tard le jour de rappel configuré)."]),
            ("3. Références", ["Processus ITIL : gestion des incidents, des demandes, amélioration continue.", "Procédure de gestion des incidents interne (PRC_GESTION_INCIDENTS).", "Règlement RGPD — minimisation des données de contact."]),
            ("4. Rôles et responsabilités", [("table", ["Rôle", "Responsabilités"], [
                ["Administrateur", "Paramétrage, comptes, sauvegardes/restaurations, RGPD, journal d'audit, intégration GLPI."],
                ["Superviseur", "Consultation des tableaux de bord et rapports, contrôle qualité, validation du rapport avant diffusion."],
                ["Hotliner", "Consultation en lecture seule des dashboards Tarkhiss, notes et base de connaissances."],
                ["PMO IT / Qualité IT", "Propriétaire de la procédure, revue périodique, suivi des écarts et des KPI qualité."]])]),
            ("5. Prérequis", ["Compte nominatif actif et rôle adéquat ; exports sources disponibles (journal d'appels, export Outlook, exports GLPI) ; sauvegarde datée de moins de 48 h."]),
            ("6. Procédure mensuelle", [("ol", [
                "Contrôler le tableau de bord Qualité des données : aucun écart bloquant sur le mois précédent.",
                "Tarkhiss : importer le journal d'appels et l'export Outlook (ou saisir les jours manquants) ; vérifier la complétude des jours ouvrés.",
                "Moussanada : importer les exports GLPI (ou import direct) : volumes, catégories, services, techniciens.",
                "Renseigner l'analyse mensuelle (catégories, constats, recommandations) ; relire les libellés de KPI.",
                "Générer l'aperçu du rapport e-mail ; contrôler en-tête, chiffres N / N-1 et graphiques.",
                "Faire valider le rapport par le superviseur, puis générer/envoyer l'e-mail ; l'envoi est tracé dans l'historique.",
                "Exporter le dossier de preuves d'audit du mois et l'archiver (PDF + ZIP d'intégrité).",
                "Créer une sauvegarde manuelle de clôture mensuelle."])]),
            ("7. Contrôles qualité", [("table", ["Contrôle", "Critère d'acceptation", "Fréquence"], [
                ["Complétude Tarkhiss", "100 % des jours ouvrés avec appels et emails saisis", "Mensuelle"],
                ["Complétude Moussanada", "Volumes, catégories, services et techniciens importés", "Mensuelle"],
                ["Analyse rédigée", "Catégories + constats + recommandations renseignés", "Mensuelle"],
                ["Diffusion tracée", "Entrée dans l'historique d'envoi pour chaque mois clos", "Mensuelle"],
                ["Sauvegarde", "Dernière sauvegarde < 48 h ; test de restauration trimestriel", "Quotidienne / trimestrielle"],
                ["Accès", "Revue des comptes actifs et des rôles", "Trimestrielle"]])]),
            ("8. Indicateurs (KPI)", ["Taux de complétude des données (cible ≥ 95 %), délai de diffusion du rapport après clôture (cible ≤ 5 jours ouvrés), taux d'écarts qualité clôturés, nombre de sauvegardes valides."]),
            ("9. Traçabilité et conservation", ["Journal d'audit append-only (connexions, créations/modifications, imports, exports, restaurations). Historique des envois et des imports. Empreintes SHA-256 des fichiers de données dans le dossier de preuves. Conservation : selon la politique de l'agence (minimum 12 mois recommandé)."]),
            ("10. Sauvegarde et restauration", ["Sauvegarde automatique quotidienne (rétention configurable) + sauvegarde manuelle. Restauration : vérification SHA-256 de l'archive, sauvegarde de sécurité automatique avant restauration, action tracée. Test de restauration trimestriel sur poste de recette."]),
            ("11. Gestion des écarts et incidents", ["Tout écart (donnée manquante, import en échec, rapport non diffusé) est consigné, corrigé et rejoué ; les incidents applicatifs suivent la procédure de gestion des incidents interne."]),
            ("12. Historique des versions", [("table", ["Version", "Date", "Auteur", "Modification"], [])]),
        ]

    def sop_html(meta=None, with_cover=True):
        m = meta or sop_meta()
        out = []
        if with_cover:
            out.append(f'''<div style="background:#0B4965;color:#fff;padding:18px 22px;border-radius:6px;">
              <div style="font-size:20px;font-weight:bold;">Procédure opérationnelle standard — Exploitation du reporting Helpdesk SI</div>
              <div style="font-size:12px;color:#B9D3E0;margin-top:4px;">{esc(m["agency"])}</div></div>''')
            out.append(table(["Référence", "Version", "Date", "Propriétaire", "Approbateur"],
                             [[esc(m["ref"]), esc(m["version"]), esc(m["date"]), esc(m["owner"]), esc(m["approver"])]]))
        for title, blocks in sop_sections():
            out.append(h2(title))
            for b in blocks:
                if isinstance(b, str):
                    out.append(p(esc(b)))
                elif b[0] == "ol":
                    out.append('<ol style="font-size:12px;line-height:1.5;">' + "".join(f"<li>{esc(x)}</li>" for x in b[1]) + "</ol>")
                elif b[0] == "table":
                    rows = b[2] if title != "12. Historique des versions" else [[esc(m["version"]), esc(m["date"]), esc(m["owner"]), "Version initiale / mise à jour"]]
                    out.append(table(b[1], [[esc(c) if title != "12. Historique des versions" else c for c in r] for r in rows]))
        return "".join(out)

    def sop_docx(meta):
        from docx import Document
        from docx.shared import Pt
        d = Document()
        d.add_heading("Procédure opérationnelle standard — Exploitation du reporting Helpdesk SI", 0)
        t = d.add_table(rows=2, cols=5); t.style = "Light Grid Accent 1"
        for i, h in enumerate(["Référence", "Version", "Date", "Propriétaire", "Approbateur"]):
            t.rows[0].cells[i].text = h
        for i, v in enumerate([meta["ref"], meta["version"], meta["date"], meta["owner"], meta["approver"]]):
            t.rows[1].cells[i].text = v
        for title, blocks in sop_sections():
            d.add_heading(title, 1)
            for b in blocks:
                if isinstance(b, str):
                    d.add_paragraph(b)
                elif b[0] == "ol":
                    for x in b[1]:
                        d.add_paragraph(x, style="List Number")
                elif b[0] == "table":
                    rows = b[2] or [[meta["version"], meta["date"], meta["owner"], "Version initiale / mise à jour"]]
                    tb = d.add_table(rows=1, cols=len(b[1])); tb.style = "Light Grid Accent 1"
                    for i, h in enumerate(b[1]):
                        tb.rows[0].cells[i].text = h
                    for r in rows:
                        cells = tb.add_row().cells
                        for i, c in enumerate(r):
                            cells[i].text = str(c)
        d.add_heading("Validation", 1)
        v = d.add_table(rows=4, cols=4); v.style = "Light Grid Accent 1"
        for i, h in enumerate(["Rôle", "Nom", "Date", "Signature"]):
            v.rows[0].cells[i].text = h
        for i, r in enumerate(["Rédigé par", "Vérifié par", "Approuvé par"], start=1):
            v.rows[i].cells[0].text = r
        buf = io.BytesIO(); d.save(buf); buf.seek(0)
        return buf

    def month_audit(ym):
        return [e for e in store.read_jsonl(H.AUDIT_LOG_FILE) if (e.get("ts") or "")[:7] == ym]

    def month_files(ym):
        res = []
        for module in ("tarkhiss", "moussanada", "pchc"):
            d = H.module_dir(module)
            for fn in store.names(d):
                if ym in fn and fn.endswith(".json"):
                    full = os.path.join(d, fn)
                    content = json.dumps(store.read_json(full, None), ensure_ascii=False, indent=2).encode("utf-8")
                    res.append((f"{module}/{fn}", hashlib.sha256(content).hexdigest(), store.mtime(full)))
        return res

    def month_kpis(ym):
        rows = []
        calls = H._load("tarkhiss", "calls", ym, {}) or {}
        emails = H._load("tarkhiss", "emails", ym, {}) or {}
        rows += [["Tarkhiss", "Appels", sum((x.get("calls") or 0) for x in calls.values())],
                 ["Tarkhiss", "Emails reçus", sum((x.get("received") or 0) for x in emails.values())],
                 ["Tarkhiss", "Emails envoyés", sum((x.get("sent") or 0) for x in emails.values())]]
        ts = H.load_timeseries().get(ym) or {}
        for k, l in (("ouverts", "Tickets ouverts"), ("resolus", "Tickets résolus"), ("clos", "Tickets clos"), ("en_retard", "Tickets en retard")):
            rows.append(["Moussanada", l, ts.get(k, 0)])
        return rows

    def audit_pack_html(ym):
        m = sop_meta()
        user = H.current_user()
        q_t, q_m = tarkhiss_quality(ym), moussanada_quality(ym)
        audit = month_audit(ym)
        parts = [f'''<div style="background:#0B4965;color:#fff;padding:18px 22px;border-radius:6px;">
          <div style="font-size:20px;font-weight:bold;">Dossier de preuves qualité — {esc(H.month_label_fr(ym))}</div>
          <div style="font-size:12px;color:#B9D3E0;margin-top:4px;">{esc(m["agency"])} — Réf. DSID-QUAL-{ym} — généré le {datetime.now().strftime("%d/%m/%Y %H:%M")} par {esc(user["username"] if user else "système")}</div></div>''']
        parts.append(h2("1. Synthèse des contrôles qualité"))
        for label, q in (("Tarkhiss", q_t), ("Moussanada", q_m)):
            parts.append(p(f"<b>{label}</b> — score de conformité : <b>{q['score'] if q['score'] is not None else 'n/a'} %</b>"))
            parts.append(table(["Contrôle", "Résultat", "Détail"], [[esc(c["label"]), "✔ Conforme" if c["ok"] else "✘ Écart", esc(c["detail"])] for c in q["checks"]]))
        if q_t["missing_calls"] or q_t["missing_emails"]:
            parts.append(p("<b>Jours ouvrés non saisis (Tarkhiss)</b> — appels : " + (", ".join(q_t["missing_calls"]) or "aucun") + " ; emails : " + (", ".join(q_t["missing_emails"]) or "aucun")))
        parts.append(h2("2. Indicateurs du mois"))
        parts.append(table(["Volet", "Indicateur", "Valeur"], [[a, b, str(c)] for a, b, c in month_kpis(ym)]))
        parts.append(h2("3. Historique d'envoi des rapports"))
        sends = [(mod, e) for mod in ("tarkhiss", "moussanada") for e in H.load_send_log(mod) if e.get("ym") == ym]
        parts.append(table(["Volet", "Date", "Objet", "Destinataires"], [[mod, esc((e.get("date") or "").replace("T", " ")), esc(e.get("subject", "")), esc("; ".join(e.get("to") or []))] for mod, e in sends]) if sends else p("Aucun envoi enregistré pour ce mois."))
        parts.append(h2("4. Historique des imports"))
        imps = [e for e in H.load_import_history() if e.get("ym") == ym]
        parts.append(table(["Date", "Type", "Mode", "Source"], [[esc(e["date"].replace("T", " ")), esc(e["kind"]), esc(e["mode"]), esc(e["source"])] for e in imps]) if imps else p("Aucun import horodaté pour ce mois."))
        parts.append(h2("5. Journal d'audit du mois"))
        counts = {}
        for e in audit:
            counts[e["action"]] = counts.get(e["action"], 0) + 1
        parts.append(p(f"{len(audit)} événement(s) — " + (", ".join(f"{esc(k)} : {v}" for k, v in sorted(counts.items())) or "aucun")))
        parts.append(table(["Horodatage", "Utilisateur", "Action", "Volet"], [[esc(e["ts"][:19].replace("T", " ")), esc(e.get("user") or ""), esc(e["action"]), esc(e.get("module") or "")] for e in audit[-60:]]) if audit else "")
        if len(audit) > 60:
            parts.append(p("(60 derniers événements affichés — le journal complet est fourni au format CSV dans l'archive ZIP.)"))
        parts.append(h2("6. Sauvegardes disponibles"))
        snaps = list_snapshots()[:8]
        parts.append(table(["Archive", "Date", "Type", "Taille (Ko)"], [[esc(s["name"]), esc(s["created"].replace("T", " ")), esc(s["reason"]), str(round(s["size"] / 1024))] for s in snaps]) if snaps else p("Aucune sauvegarde."))
        parts.append(h2("7. Intégrité des données du mois (SHA-256)"))
        files = month_files(ym)
        parts.append(table(["Fichier", "Dernière modification", "SHA-256"], [[esc(f), datetime.fromtimestamp(t).strftime("%d/%m/%Y %H:%M"), f'<span style="font-family:monospace;font-size:9px;">{h}</span>'] for f, h, t in files]) if files else p("Aucun fichier de données pour ce mois."))
        body = "".join(parts)
        digest = hashlib.sha256(body.encode("utf-8")).hexdigest()
        parts.append(h2("8. Validation"))
        parts.append(table(["Rôle", "Nom", "Date", "Signature"], [["Rédigé par", "", "", ""], ["Vérifié par (Qualité IT)", "", "", ""], ["Approuvé par", "", "", ""]]))
        parts.append(p(f'Empreinte SHA-256 du contenu du dossier (sections 1 à 7) : <span style="font-family:monospace;font-size:9px;">{digest}</span>'))
        return "".join(parts), digest, audit

    def pdf_response(html, name, label):
        return send_file(H.render_pdf_bytes(html, header_label=label), as_attachment=True, download_name=name, mimetype="application/pdf")

    @app.route("/api/admin/sop-settings", methods=["GET"])
    def sop_settings_get():
        g = gs()
        return jsonify({**sop_meta(), "holidays": g.get("quality_holidays", ""), "date": g.get("sop_date") or ""})

    @app.route("/api/admin/sop")
    def sop_export():
        fmt = request.args.get("format", "html")
        m = sop_meta()
        H.log_audit("sop_export", {"format": fmt, "version": m["version"]})
        if fmt == "docx":
            return send_file(sop_docx(m), as_attachment=True, download_name=f"{m['ref']}_v{m['version']}.docx",
                             mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        html = f'<div style="font-family:Arial,sans-serif;">{sop_html(m)}</div>'
        if fmt == "pdf":
            return pdf_response(html, f"{m['ref']}_v{m['version']}.pdf", f"{m['ref']} v{m['version']}")
        return jsonify({"html": html})

    @app.route("/api/admin/audit-pack/<ym>")
    def audit_pack(ym):
        if not YM_RE.match(ym):
            return jsonify({"error": "mois invalide"}), 400
        fmt = request.args.get("format", "html")
        body, digest, audit = audit_pack_html(ym)
        html = f'<div style="font-family:Arial,sans-serif;">{body}</div>'
        if fmt == "html":
            return jsonify({"html": html, "sha256": digest})
        H.log_audit("audit_pack_export", {"ym": ym, "format": fmt, "sha256": digest})
        if fmt == "pdf":
            return pdf_response(html, f"dossier_preuves_{ym}.pdf", f"Dossier de preuves {ym}")
        if fmt == "zip":
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                pdf = H.render_pdf_bytes(html, header_label=f"Dossier de preuves {ym}").read()
                z.writestr(f"dossier_preuves_{ym}.pdf", pdf)
                m = sop_meta()
                z.writestr(f"{m['ref']}_v{m['version']}.docx", sop_docx(m).read())
                s = io.StringIO(); w = csv.writer(s, delimiter=";")
                w.writerow(["ts", "user", "action", "module", "details"])
                for e in audit:
                    w.writerow([e.get("ts"), e.get("user"), e.get("action"), e.get("module"), json.dumps(e.get("details") or {}, ensure_ascii=False)])
                z.writestr(f"journal_audit_{ym}.csv", "﻿" + s.getvalue())
                sends = io.StringIO(); w = csv.writer(sends, delimiter=";"); w.writerow(["volet", "date", "objet", "to", "cc"])
                for mod in ("tarkhiss", "moussanada"):
                    for e in H.load_send_log(mod):
                        if e.get("ym") == ym:
                            w.writerow([mod, e.get("date"), e.get("subject"), "; ".join(e.get("to") or []), "; ".join(e.get("cc") or [])])
                z.writestr(f"historique_envois_{ym}.csv", "﻿" + sends.getvalue())
                manifest = {"ym": ym, "generated": datetime.now().isoformat(timespec="seconds"), "dossier_sha256": digest,
                            "data_files": {f: h for f, h, _ in month_files(ym)}}
                z.writestr("manifest_integrite.json", json.dumps(manifest, ensure_ascii=False, indent=1))
            buf.seek(0)
            return send_file(buf, as_attachment=True, download_name=f"preuves_audit_{ym}.zip", mimetype="application/zip")
        return jsonify({"error": "format inconnu"}), 400
