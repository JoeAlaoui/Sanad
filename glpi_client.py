# -*- coding: utf-8 -*-
"""
glpi_client.py — Client minimal pour l'API REST "legacy" de GLPI (apirest.php).

Ce module est volontairement séparé de glpi_import.py (qui gère le parsing de fichiers) :
glpi_client.py ne fait que parler au réseau (initSession / search / killSession), sans aucune
connaissance du format de stockage SANAD. Le mapping GLPI -> SANAD se fait plus haut, dans les
"profils d'import" (glpi_profiles.py, lot suivant).

Pas d'accès serveur GLPI supposé : uniquement l'API REST, avec un compte disposant d'un
App-Token (jeton applicatif, généré dans Configuration > Générale > API de GLPI) et d'un
User-Token (jeton personnel, généré dans les préférences de l'utilisateur GLPI).
"""
import requests


class GLPIConnectionError(Exception):
    """Toute erreur de connexion, d'authentification ou de requête vers l'API GLPI."""
    pass


class GLPIClient:
    def __init__(self, base_url, app_token, user_token, entity_id=None, timeout=15):
        if not base_url:
            raise GLPIConnectionError("URL de l'API GLPI manquante (ex: https://glpi.ammps.ma/apirest.php)")
        self.base_url = base_url.rstrip("/")
        self.app_token = app_token
        self.user_token = user_token
        self.entity_id = entity_id
        self.timeout = timeout
        self.session_token = None

    # ------------------------------------------------------------------
    # Cycle de session
    # ------------------------------------------------------------------
    def init_session(self):
        """Ouvre une session GLPI (App-Token + User-Token) et conserve le Session-Token."""
        if not self.app_token or not self.user_token:
            raise GLPIConnectionError("App-Token et/ou User-Token manquant(s)")
        url = f"{self.base_url}/initSession"
        headers = {
            "App-Token": self.app_token,
            "Authorization": f"user_token {self.user_token}",
        }
        params = {}
        if self.entity_id:
            params["entities_id"] = self.entity_id
        try:
            r = requests.get(url, headers=headers, params=params, timeout=self.timeout)
        except requests.RequestException as e:
            raise GLPIConnectionError(f"Connexion impossible à {self.base_url} : {e}")
        if r.status_code != 200:
            raise GLPIConnectionError(self._format_glpi_error("initSession", r))
        try:
            data = r.json()
        except ValueError:
            raise GLPIConnectionError("Réponse initSession illisible (l'URL pointe-t-elle bien vers apirest.php ?)")
        self.session_token = data.get("session_token")
        if not self.session_token:
            raise GLPIConnectionError("initSession n'a pas renvoyé de session_token")
        return self.session_token

    def kill_session(self):
        """Ferme proprement la session GLPI. N'échoue jamais bruyamment (best-effort)."""
        if not self.session_token:
            return
        url = f"{self.base_url}/killSession"
        try:
            requests.get(url, headers=self._headers(), timeout=self.timeout)
        except requests.RequestException:
            pass
        self.session_token = None

    def __enter__(self):
        self.init_session()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.kill_session()
        return False

    # ------------------------------------------------------------------
    # Requêtes
    # ------------------------------------------------------------------
    def _headers(self):
        return {"App-Token": self.app_token, "Session-Token": self.session_token}

    def _format_glpi_error(self, action, response):
        try:
            body = response.json()
            detail = body[1] if isinstance(body, list) and len(body) > 1 else body
        except ValueError:
            detail = response.text[:300]
        return f"{action} a échoué (HTTP {response.status_code}) : {detail}"

    def search(self, itemtype, criteria=None, forcedisplay=None, range_=None, sort=None, order=None):
        """Appelle GET /search/<itemtype>. criteria : liste de dicts GLPI standards, ex.
        [{"field": 15, "searchtype": "morethan", "value": "2026-08-01"}]."""
        if not self.session_token:
            raise GLPIConnectionError("Session GLPI non initialisée (appeler init_session() d'abord)")
        url = f"{self.base_url}/search/{itemtype}"
        params = {}
        for i, c in enumerate(criteria or []):
            for k, v in c.items():
                params[f"criteria[{i}][{k}]"] = v
        for i, f in enumerate(forcedisplay or []):
            params[f"forcedisplay[{i}]"] = f
        if range_:
            params["range"] = range_
        if sort is not None:
            params["sort"] = sort
        if order:
            params["order"] = order
        try:
            r = requests.get(url, headers=self._headers(), params=params, timeout=self.timeout)
        except requests.RequestException as e:
            raise GLPIConnectionError(f"Erreur réseau lors de la recherche {itemtype} : {e}")
        if r.status_code not in (200, 206):
            raise GLPIConnectionError(self._format_glpi_error(f"search/{itemtype}", r))
        try:
            return r.json()
        except ValueError:
            raise GLPIConnectionError(f"Réponse search/{itemtype} illisible")

    def get_all(self, itemtype, criteria=None, forcedisplay=None, page_size=200, max_pages=200):
        """Pagine automatiquement search() jusqu'à épuisement des résultats (évite la limite
        par défaut de GLPI sur le nombre de lignes renvoyées en un seul appel)."""
        results = []
        start = 0
        for _ in range(max_pages):
            data = self.search(itemtype, criteria=criteria, forcedisplay=forcedisplay,
                                range_=f"{start}-{start + page_size - 1}")
            rows = data.get("data", [])
            results.extend(rows)
            total = data.get("totalcount", len(results))
            start += page_size
            if start >= total or not rows:
                break
        return results


def test_connection(base_url, app_token, user_token, entity_id=None):
    """Utilitaire pour le bouton "Tester la connexion" du panneau Admin. Ouvre puis ferme
    immédiatement une session. Retourne (ok: bool, message: str)."""
    client = GLPIClient(base_url, app_token, user_token, entity_id)
    try:
        client.init_session()
        client.kill_session()
        return True, "Connexion GLPI réussie."
    except GLPIConnectionError as e:
        return False, str(e)
