#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
interface_web.py  —  « Facturo » : de la photo au tableur
=================================================================
Interface web glisser-déposer pour un usage NON technique.
On dépose des photos ou des PDF de factures/reçus, on récupère un Excel.

Bilingue FR/EN. Aucune mention du moteur d'IA côté public.

  ▸ Pour renommer le produit : remplacez "Facturo" et la baseline dans
    les dictionnaires I18N ci-dessous (une seule fois).

Lancer en local :
    export ANTHROPIC_API_KEY="sk-ant-..."      # ou OPENAI_API_KEY
    pip install fastapi uvicorn python-multipart
    python interface_web.py            # ou : uvicorn interface_web:app --port 8000

Le port est lu depuis la variable d'environnement PORT (défaut 7860),
pour être compatible avec les hébergeurs gratuits (Hugging Face, Render...).
=================================================================
"""
from __future__ import annotations

import logging
import os
import tempfile
import threading
import time
import uuid
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, Request, Body
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from starlette.concurrency import run_in_threadpool

import facture_vers_excel as fve
import releves_bancaires as rb

app = FastAPI(title="Facturo")
_journal = logging.getLogger("facturo")

# jeton -> (chemin de l'Excel, heure de création). Les fichiers produits sont
# effacés après DUREE_TELECHARGEMENT secondes : on ne garde pas les données.
_TELECHARGEMENTS: dict[str, tuple[str, float]] = {}
# jeton -> pièces lues et options de l'envoi, pour reconstruire l'Excel quand
# le comptable corrige un compte (même durée de vie que le fichier produit).
_SESSIONS: dict[str, dict] = {}
_DOSSIER = tempfile.mkdtemp(prefix="facturo_")
DUREE_TELECHARGEMENT = int(os.getenv("FACTURO_DUREE_TELECHARGEMENT", "3600"))
# Limites d'envoi (protègent la mémoire et le disque du serveur).
TAILLE_MAX_FICHIER = int(os.getenv("FACTURO_TAILLE_MAX_MO", "15")) * 1024 * 1024
MAX_FICHIERS_REQUETE = int(os.getenv("FACTURO_MAX_FICHIERS", "20"))

_ENTETES_SECURITE = {
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "Content-Security-Policy": (
        "default-src 'self'; script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src https://fonts.gstatic.com; img-src 'self' data:; "
        "connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; "
        "form-action 'self'"),
}


@app.middleware("http")
async def _entetes_securite(request: Request, appel_suivant):
    reponse = await appel_suivant(request)
    for nom, valeur in _ENTETES_SECURITE.items():
        reponse.headers.setdefault(nom, valeur)
    return reponse


def _purger_telechargements() -> None:
    """Efface les fichiers produits plus vieux que DUREE_TELECHARGEMENT."""
    limite = time.time() - DUREE_TELECHARGEMENT
    for jeton, (chemin, cree) in list(_TELECHARGEMENTS.items()):
        if cree < limite:
            _TELECHARGEMENTS.pop(jeton, None)
            _SESSIONS.pop(jeton, None)
            for f in (chemin, chemin[:-5] + "_qbo.csv", chemin[:-5] + "_base.xlsx"):
                try:
                    os.remove(f)
                except OSError:
                    pass


async def _lire_limite(up: UploadFile) -> Optional[bytes]:
    """Contenu d'un fichier envoyé, ou None s'il dépasse TAILLE_MAX_FICHIER."""
    contenu = await up.read(TAILLE_MAX_FICHIER + 1)
    return None if len(contenu) > TAILLE_MAX_FICHIER else contenu


_MESSAGES = {
    "trop_gros": {"fr": "Fichier trop volumineux (max {mo} Mo)",
                  "en": "File too large (max {mo} MB)"},
    "service": {"fr": "Service de lecture momentanément indisponible, réessayez plus tard",
                "en": "Reading service temporarily unavailable, please try again later"},
    "illisible": {"fr": "Fichier illisible ou format non pris en charge",
                  "en": "Unreadable file or unsupported format"},
    "rien": {"fr": "Ajoutez au moins une facture ou un relevé bancaire",
             "en": "Add at least one invoice or a bank statement"},
    "releve": {"fr": "Relevé bancaire illisible (formats acceptés : CSV, OFX, QFX)",
               "en": "Unreadable bank statement (accepted formats: CSV, OFX, QFX)"},
    "trop_de_fichiers": {"fr": "Trop de fichiers dans un même envoi (max {n})",
                         "en": "Too many files in one upload (max {n})"},
}


def _message(cle: str, langue: str) -> str:
    return _MESSAGES[cle][langue].format(mo=TAILLE_MAX_FICHIER // (1024 * 1024),
                                         n=MAX_FICHIERS_REQUETE)


def _pret() -> bool:
    return bool(os.getenv("ANTHROPIC_API_KEY") or os.getenv("OPENAI_API_KEY"))


# Quota de démo : protège le crédit API d'un lien public. Compteurs en mémoire
# (fichiers traités par jour), remis à zéro chaque jour et au redémarrage.
QUOTA_JOUR_TOTAL = int(os.getenv("QUOTA_JOUR_TOTAL", "30"))
QUOTA_JOUR_VISITEUR = int(os.getenv("QUOTA_JOUR_VISITEUR", "10"))
# Fichiers lus en même temps (appels API en parallèle).
PARALLELES = int(os.getenv("FACTURO_PARALLELE", "4"))
# Adresse de contact pour la version sur mesure (Facturo Pro).
CONTACT = os.getenv("FACTURO_CONTACT", "degnonaurel@gmail.com")
_FUSEAU = timezone(timedelta(hours=-5))   # heure normale de l'Est (Ottawa)
_quota = {"jour": None, "total": 0, "visiteurs": {}}
_verrou_quota = threading.Lock()


def _visiteur(request: Request) -> str:
    # Render est servi derrière Cloudflare, qui écrit lui-même l'IP du client
    # dans CF-Connecting-IP / True-Client-IP (une valeur envoyée par le client
    # est écrasée) : on les lit d'abord. X-Forwarded-For, que le client peut
    # préremplir, ne sert qu'en dernier recours.
    for entete in ("cf-connecting-ip", "true-client-ip"):
        ip = (request.headers.get(entete) or "").strip()
        if ip:
            return ip
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    return request.client.host if request.client else "inconnu"


def _restant(visiteur: str) -> int:
    """Fichiers encore autorisés aujourd'hui pour ce visiteur."""
    jour = datetime.now(_FUSEAU).date()
    with _verrou_quota:
        if _quota["jour"] != jour:
            return min(QUOTA_JOUR_VISITEUR, QUOTA_JOUR_TOTAL)
        return max(0, min(QUOTA_JOUR_VISITEUR - _quota["visiteurs"].get(visiteur, 0),
                          QUOTA_JOUR_TOTAL - _quota["total"]))


def _reserver_quota(visiteur: str, n: int) -> bool:
    """Réserve n fichiers dans le quota du jour ; False si dépassement."""
    jour = datetime.now(_FUSEAU).date()
    with _verrou_quota:
        if _quota["jour"] != jour:
            _quota.update(jour=jour, total=0, visiteurs={})
        deja = _quota["visiteurs"].get(visiteur, 0)
        if (_quota["total"] + n > QUOTA_JOUR_TOTAL
                or deja + n > QUOTA_JOUR_VISITEUR):
            return False
        _quota["total"] += n
        _quota["visiteurs"][visiteur] = deja + n
        return True


@app.get("/api/statut")
def statut(request: Request):
    # On n'expose jamais le nom du moteur : seulement l'état de service.
    return {"pret": _pret(), "quota": QUOTA_JOUR_VISITEUR,
            "restant": _restant(_visiteur(request)), "contact": CONTACT}


def _plan_affiche(plan: list[dict]) -> list[dict]:
    """Comptes proposés dans les listes déroulantes : identifiant + libellé."""
    return [{"id": fve._id_compte(c), "libelle": fve.libelle_compte(plan, fve._id_compte(c)),
             "type": c.get("type", "")} for c in plan]


def _id_de_libelle(plan: list[dict], libelle: str) -> str:
    for c in plan:
        if libelle and fve.libelle_compte(plan, fve._id_compte(c)) == libelle:
            return fve._id_compte(c)
    return ""


async def _lire_plan(plan: Optional[UploadFile], tmp: str) -> list[dict]:
    """Plan comptable d'un fichier envoyé (Excel ou CSV) ; [] s'il est illisible."""
    ext = ".csv" if plan.filename.lower().endswith(".csv") else ".xlsx"
    plan_path = os.path.join(tmp, "plan" + ext)
    contenu = await _lire_limite(plan)
    if contenu is None:
        return []
    with open(plan_path, "wb") as f:
        f.write(contenu)
    try:
        return fve.lire_plan_comptable(plan_path)
    except Exception:
        _journal.warning("Plan comptable illisible", exc_info=True)
        return []


@app.post("/api/plan")
async def apercu_plan(plan: UploadFile = File(...), langue: str = Form("fr")):
    """Aperçu d'un plan comptable, avant tout envoi de pièces (sans lecture IA,
    donc hors quota) : le comptable voit tout de suite les comptes reconnus."""
    langue = "en" if langue == "en" else "fr"
    with tempfile.TemporaryDirectory() as tmp:
        comptes = await _lire_plan(plan, tmp) if plan.filename else []
    if not comptes:
        return JSONResponse({"comptes": [], "erreur": (
            "Chart of accounts unreadable (Excel or CSV with a code and a name column)"
            if langue == "en" else
            "Plan comptable illisible (Excel ou CSV avec une colonne code et une colonne nom)")},
            status_code=400)
    return {"comptes": _plan_affiche(comptes)}


@app.get("/api/plan-standard")
def plan_standard(langue: str = "fr"):
    langue = "en" if langue == "en" else "fr"
    return {"comptes": _plan_affiche(fve.plan_par_defaut(langue))}


@app.post("/api/extraire")
async def extraire(request: Request,
                   fichiers: Optional[list[UploadFile]] = File(None),
                   base: Optional[UploadFile] = File(None),
                   plan: Optional[UploadFile] = File(None),
                   releve: Optional[UploadFile] = File(None),
                   langue: str = Form("fr"),
                   qbo: str = Form("")):
    langue = "en" if langue == "en" else "fr"
    veut_qbo = qbo in ("1", "true", "on")
    _purger_telechargements()
    fichiers = [f for f in (fichiers or []) if f.filename]
    if releve is not None and not releve.filename:
        releve = None
    if not fichiers and releve is None:
        return JSONResponse({"erreurs": [{"fichier": "", "message":
                             _message("rien", langue)}]}, status_code=400)
    if len(fichiers) > MAX_FICHIERS_REQUETE:
        return JSONResponse({"erreurs": [{"fichier": "", "message":
                             _message("trop_de_fichiers", langue)}]}, status_code=413)
    if not _reserver_quota(_visiteur(request), len(fichiers)):
        return JSONResponse({"quota_atteint": True}, status_code=429)
    factures, erreurs = [], []
    with tempfile.TemporaryDirectory() as tmp:
        # Excel existant optionnel : les nouvelles lignes y seront ajoutées.
        base_path = None
        if base is not None and base.filename:
            contenu = await _lire_limite(base)
            if contenu is None:
                erreurs.append({"fichier": base.filename,
                                "message": _message("trop_gros", langue)})
            else:
                base_path = os.path.join(tmp, "base.xlsx")
                with open(base_path, "wb") as f:
                    f.write(contenu)

        # Plan comptable : fichier fourni > onglet du classeur de base > défaut.
        plan_comptable = []
        if plan is not None and plan.filename:
            plan_comptable = await _lire_plan(plan, tmp)
            if not plan_comptable:
                erreurs.append({"fichier": plan.filename, "message": (
                    "Chart of accounts unreadable, default plan used" if langue == "en"
                    else "Plan comptable illisible, plan par défaut utilisé")})
        if not plan_comptable and base_path:
            try:
                plan_comptable = fve.plan_du_classeur(base_path)
            except Exception:
                pass
        plan_comptable = plan_comptable or fve.plan_par_defaut(langue)

        chemins, recus = [], []
        for up in fichiers:
            contenu = await _lire_limite(up)
            if contenu is None:          # trop gros : signalé, le lot continue
                erreurs.append({"fichier": up.filename,
                                "message": _message("trop_gros", langue)})
                continue
            ext = os.path.splitext(up.filename or "")[1][:10] or ".bin"
            dest = os.path.join(tmp, f"{uuid.uuid4().hex}{ext}")
            with open(dest, "wb") as f:
                f.write(contenu)
            chemins.append(dest)
            recus.append(up)

        # Lecture en parallèle, hors de la boucle d'événements du serveur.
        resultats = await run_in_threadpool(
            fve.traiter_lot, chemins, paralleles=PARALLELES, plan=plan_comptable)
        for up, res in zip(recus, resultats):
            # Détail technique dans les journaux ; message générique à l'écran
            # (ni nom du moteur, ni réponse brute de l'API).
            if isinstance(res, fve.ErreurService):
                _journal.error("Service de lecture : %s", res)
                erreurs.append({"fichier": up.filename,
                                "message": _message("service", langue)})
            elif isinstance(res, Exception):
                _journal.warning("Fichier illisible %s : %r", up.filename, res)
                erreurs.append({"fichier": up.filename,
                                "message": _message("illisible", langue)})
            else:
                res.fichier = up.filename or res.fichier
                factures.append(res)

        # Relevé bancaire optionnel : rapproché de toutes les pièces du classeur.
        transactions = None
        if releve is not None:
            contenu = await _lire_limite(releve)
            ext = os.path.splitext(releve.filename)[1].lower()
            if contenu is None:
                erreurs.append({"fichier": releve.filename,
                                "message": _message("trop_gros", langue)})
            else:
                chemin_releve = os.path.join(
                    tmp, "releve" + (ext if ext in (".csv", ".ofx", ".qfx", ".txt") else ".csv"))
                with open(chemin_releve, "wb") as f:
                    f.write(contenu)
                try:
                    transactions = rb.lire_releve(chemin_releve)
                except Exception:
                    _journal.warning("Relevé illisible", exc_info=True)
                    erreurs.append({"fichier": releve.filename,
                                    "message": _message("releve", langue)})

        if not factures and transactions is None:
            return JSONResponse({"factures": [], "erreurs": erreurs,
                                 "download_id": None, "ajoute": bool(base_path)})

        factures = fve.regrouper_factures(factures)
        bilan: dict = {}
        jeton = uuid.uuid4().hex
        chemin = os.path.join(_DOSSIER, f"facturo_{jeton}.xlsx")
        try:
            fve.construire_excel(factures, chemin, base_excel=base_path, langue=langue,
                                 plan=plan_comptable, releve=transactions, bilan=bilan)
        except Exception:
            # Excel de base illisible : on repart sur un fichier neuf.
            _journal.warning("Excel de base illisible", exc_info=True)
            erreurs.append({"fichier": base.filename if base else "Excel existant",
                            "message": ("Existing Excel unreadable, new file created"
                                        if langue == "en" else
                                        "Excel existant illisible, nouveau fichier créé")})
            fve.construire_excel(factures, chemin, langue=langue, plan=plan_comptable,
                                 releve=transactions, bilan=bilan)
        _TELECHARGEMENTS[jeton] = (chemin, time.time())
        # Copie de l'Excel de base : il resservira si un compte est corrigé.
        base_copie = None
        if base_path:
            base_copie = chemin[:-5] + "_base.xlsx"
            with open(base_path, "rb") as src, open(base_copie, "wb") as dst:
                dst.write(src.read())
        _SESSIONS[jeton] = {"factures": factures, "plan": plan_comptable, "langue": langue,
                            "base": base_copie, "releve": transactions, "qbo": veut_qbo}
        # CSV QuickBooks Online, en option et en plus de l'Excel (jamais à sa
        # place) : seulement les pièces de cet envoi, pour ne pas importer de
        # doublons quand on ajoute à un Excel existant.
        qbo_pret = False
        if veut_qbo and factures:
            try:
                fve.construire_csv_qbo(factures, chemin[:-5] + "_qbo.csv",
                                       plan=plan_comptable)
                qbo_pret = True
            except Exception:
                _journal.warning("Export QuickBooks impossible", exc_info=True)
                erreurs.append({"fichier": "QuickBooks",
                                "message": _message("illisible", langue)})
        try:
            totaux = fve.totaux_resume(chemin)
        except Exception:
            totaux = []

    return {"factures": _resume(factures, plan_comptable, langue), "erreurs": erreurs,
            "download_id": jeton, "ajoute": bool(base_path), "totaux_fichier": totaux,
            "qbo": qbo_pret, "rapprochement": bilan or None,
            "plan": _plan_affiche(plan_comptable), "plan_fourni": bool(plan is not None
                                                                     and plan.filename),
            "restant": _restant(_visiteur(request))}


def _resume(factures: list, plan: list[dict], langue: str) -> list[dict]:
    return [{
        "fichier": f.fichier, "fournisseur": f.fournisseur, "numero": f.numero,
        "date": f.date, "devise": f.devise, "total_ht": f.total_ht,
        "tps": f.tps, "tvq": f.tvq, "tva": f.tva, "total_ttc": f.total_ttc,
        "nb_lignes": len(f.lignes), "pieces_liees": f.pieces_liees,
        "categorie": f.categorie, "compte": _id_de_libelle(plan, f.categorie),
        "alertes": [fve.traduire_alerte(a, langue) for a in f.alertes],
    } for f in factures]


@app.post("/api/comptes")
def corriger_comptes(donnees: dict = Body(...)):
    """Le comptable choisit le compte de chaque pièce dans SON plan : on
    reconstruit l'Excel (et le CSV QuickBooks) avec ses choix, au même lien."""
    _purger_telechargements()
    jeton = str(donnees.get("jeton", ""))
    session = _SESSIONS.get(jeton)
    chemin = _TELECHARGEMENTS.get(jeton, (None, 0))[0]
    if not session or not chemin:
        return JSONResponse({"erreur": "Fichier introuvable ou expiré."}, status_code=404)
    plan, factures = session["plan"], session["factures"]
    ids = {fve._id_compte(c) for c in plan}
    choix = donnees.get("comptes") or {}
    if not isinstance(choix, dict):
        return JSONResponse({"erreur": "Format invalide."}, status_code=400)
    for cle, ident in choix.items():
        try:
            i = int(cle)
        except (TypeError, ValueError):
            continue
        if 0 <= i < len(factures) and str(ident) in ids:
            libelle = fve.libelle_compte(plan, str(ident))
            # Toute la pièce va dans le compte choisi, articles compris.
            factures[i].categorie = libelle
            for ligne in factures[i].lignes:
                ligne.categorie = libelle
    bilan: dict = {}
    options = {"langue": session["langue"], "plan": plan, "releve": session["releve"]}
    try:
        fve.construire_excel(factures, chemin, base_excel=session["base"], bilan=bilan,
                             **options)
    except Exception:
        # Excel de base illisible (déjà signalé au premier envoi) : fichier neuf.
        _journal.warning("Excel de base illisible", exc_info=True)
        fve.construire_excel(factures, chemin, bilan=bilan, **options)
    if session["qbo"] and factures:
        fve.construire_csv_qbo(factures, chemin[:-5] + "_qbo.csv", plan=plan)
    try:
        totaux = fve.totaux_resume(chemin)
    except Exception:
        totaux = []
    return {"factures": _resume(factures, plan, session["langue"]),
            "totaux_fichier": totaux, "rapprochement": bilan or None}


@app.get("/telecharger/{jeton}")
def telecharger(jeton: str, format: str = "xlsx"):
    _purger_telechargements()
    chemin = _TELECHARGEMENTS.get(jeton, (None, 0))[0]
    if chemin and format == "qbo":
        chemin = chemin[:-5] + "_qbo.csv"
        if not os.path.isfile(chemin):
            return JSONResponse({"erreur": "Fichier introuvable ou expiré."}, status_code=404)
        return FileResponse(chemin, media_type="text/csv",
                            filename="facturo_quickbooks.csv")
    if not chemin or not os.path.isfile(chemin):
        return JSONResponse({"erreur": "Fichier introuvable ou expiré."}, status_code=404)
    return FileResponse(
        chemin,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="facturo.xlsx")


@app.get("/", response_class=HTMLResponse)
def accueil():
    return _PAGE


_PAGE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Facturo — de la photo au tableur</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@100..125,400..800&family=IBM+Plex+Mono:wght@500&display=swap" rel="stylesheet">
<style>
  /* Direction « blocs en relief » : fond vert d'eau (rappel du tableur),
     cartes blanches à bordure épaisse et ombre pleine décalée, accent orange
     pour l'action. Une seule table de travail : deux modes visibles en haut ;
     en mode comptable, trois étapes côte à côte, puis la vérification des comptes. */
  :root{
    --bg:#CFEBDD; --surface:#FFFFFF; --surface2:#F1F8F4; --ink:#0B2B24; --muted:#2E4A42;
    --trait:#0B2B24; --line:#B7D6C8; --pri:#0E6E5E; --pri-fort:#0A5849; --pri-doux:#E1F2EA;
    --acc:#E2552B; --acc-fort:#C7451E; --acc-doux:#FCE6DD; --warn:#9A4A0B; --warn-doux:#FDEBD6;
    --ombre:6px 6px 0 var(--trait); --ombre-p:3px 3px 0 var(--trait);
    --f-titre:"Archivo","Segoe UI",system-ui,sans-serif;
    --f-texte:"Archivo","Segoe UI",system-ui,sans-serif;
    --f-chiffre:"IBM Plex Mono",ui-monospace,Menlo,monospace;
    --r:16px;
  }
  @media (prefers-color-scheme: dark){ :root:not([data-theme="light"]){
    --bg:#0C1A16; --surface:#132520; --surface2:#172C26; --ink:#E4F3EC; --muted:#A5C2B6;
    --trait:#000000; --line:#28433A; --pri:#3BC1A4; --pri-fort:#5BD3B8; --pri-doux:#163A31;
    --acc:#FF7A4D; --acc-fort:#FF956F; --acc-doux:#3A2017; --warn:#F0AE5E; --warn-doux:#35260F; color-scheme:dark }}
  :root[data-theme="dark"]{
    --bg:#0C1A16; --surface:#132520; --surface2:#172C26; --ink:#E4F3EC; --muted:#A5C2B6;
    --trait:#000000; --line:#28433A; --pri:#3BC1A4; --pri-fort:#5BD3B8; --pri-doux:#163A31;
    --acc:#FF7A4D; --acc-fort:#FF956F; --acc-doux:#3A2017; --warn:#F0AE5E; --warn-doux:#35260F; color-scheme:dark }
  *{box-sizing:border-box}
  html,body{margin:0}
  body{font:15px/1.55 var(--f-texte);background:var(--bg);color:var(--ink);min-height:100vh}
  a{color:var(--pri)}
  button{font:inherit}
  :focus-visible{outline:3px solid var(--acc);outline-offset:2px}
  .masque{display:none!important}
  .conteneur{max-width:1100px;margin:0 auto;padding-left:max(18px,env(safe-area-inset-left,0px));
    padding-right:max(18px,env(safe-area-inset-right,0px))}
  .large{font-stretch:125%}

  /* En-tête */
  .barre{display:flex;align-items:center;gap:14px;padding-block:18px}
  .logo{display:flex;align-items:center;gap:10px;font-family:var(--f-titre);font-stretch:125%;font-weight:800;
    font-size:1.25rem;text-decoration:none;color:var(--ink)}
  .droite{margin-left:auto;display:flex;align-items:center;gap:14px}
  .statut{display:inline-flex;align-items:center;gap:7px;font-size:.82rem;color:var(--muted);font-weight:600}
  .pastille{width:9px;height:9px;border-radius:50%;background:var(--line);border:1.5px solid var(--trait)}
  .pastille.on{background:#2BB673} .pastille.off{background:var(--warn)}
  .lang{display:flex;border:2px solid var(--trait);border-radius:10px;overflow:hidden;font-size:.8rem;font-weight:700;
    box-shadow:var(--ombre-p);background:var(--surface)}
  .lang button{border:0;background:transparent;color:var(--ink);padding:6px 11px;cursor:pointer}
  .lang button.actif{background:var(--ink);color:var(--surface)}

  /* Accroche : texte + illustration en relief */
  .hero{display:grid;grid-template-columns:minmax(0,1.1fr) minmax(0,.9fr);gap:28px;align-items:center;padding-block:18px 30px}
  .hero-texte{display:grid;gap:14px;align-content:start}
  .surtitre{font-size:.76rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--pri)}
  .hero h1{font-family:var(--f-titre);font-stretch:125%;font-weight:800;font-size:clamp(2.1rem,5vw,3.4rem);line-height:1;
    letter-spacing:-.02em;margin:0;text-wrap:balance}
  .hero p{margin:0;color:var(--muted);font-size:1.06rem;max-width:52ch}
  .pas-liste{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:6px}
  .pas{background:var(--surface);border:2px solid var(--trait);border-radius:12px;padding:12px;display:grid;gap:3px;
    align-content:start;box-shadow:var(--ombre-p);min-width:0}
  .pas .n{font-stretch:125%;font-weight:800;color:var(--acc)}
  .pas b{font-size:.92rem} .pas span:last-child{color:var(--muted);font-size:.84rem;line-height:1.4}
  .illus{position:relative;height:300px}
  .illus .recu{position:absolute;left:4%;top:16px;width:46%;max-width:200px;height:250px;background:var(--surface);
    border:2.5px solid var(--trait);border-radius:14px;transform:rotate(-7deg);box-shadow:8px 8px 0 var(--trait);
    padding:16px;display:flex;flex-direction:column;gap:9px}
  .illus .recu i{display:block;height:8px;border-radius:4px;background:var(--line)}
  .illus .recu i:first-child{height:11px;width:70%;background:var(--ink)}
  .illus .recu .tx{margin-top:auto;display:flex;justify-content:space-between;font-weight:800;font-size:.85rem}
  .illus .recu .tx em{font-style:normal;background:var(--acc);color:#fff;padding:1px 7px;border-radius:5px}
  .illus .feuille{position:absolute;right:2%;bottom:6px;width:56%;max-width:250px;background:var(--surface);
    border:2.5px solid var(--trait);border-radius:14px;transform:rotate(5deg);box-shadow:8px 8px 0 var(--trait);
    overflow:hidden;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));font-size:.72rem;font-weight:600}
  .illus .feuille span{padding:8px 7px;border-top:1.5px solid var(--trait);font-family:var(--f-chiffre);font-weight:500}
  .illus .feuille span.t{background:var(--pri);color:#fff;border-top:0;font-family:var(--f-texte);font-weight:700}
  .illus .fleche{position:absolute;left:44%;top:120px;width:56px;height:56px;border-radius:50%;background:var(--acc);
    border:2.5px solid var(--trait);box-shadow:4px 4px 0 var(--trait);display:grid;place-items:center;color:#fff}

  /* Carte de travail */
  .atelier{background:var(--surface);border:2.5px solid var(--trait);border-radius:20px;box-shadow:var(--ombre);overflow:hidden}
  .modes{display:grid;grid-template-columns:1fr 1fr;border-bottom:2.5px solid var(--trait)}
  .mode{display:flex;gap:12px;align-items:flex-start;text-align:left;border:0;background:var(--surface2);cursor:pointer;
    padding:16px 20px;color:var(--muted)}
  .mode+.mode{border-left:2.5px solid var(--trait)}
  .mode svg{flex:0 0 22px;width:22px;height:22px;margin-top:2px}
  .mode b{display:block;font-stretch:125%;font-weight:800;font-size:1rem;color:var(--ink)}
  .mode span{font-size:.84rem;line-height:1.4}
  .mode[aria-selected="true"]{background:var(--pri);color:#E6F5EE}
  .mode[aria-selected="true"] b{color:#fff}
  .mode:not([aria-selected="true"]):hover{background:var(--pri-doux)}
  .corps{padding:22px}
  .etapes{display:grid;gap:18px}
  body.comptable .etapes{grid-template-columns:minmax(0,1fr) minmax(0,1.25fr) minmax(0,1fr)}
  .etape{display:grid;gap:12px;align-content:start;min-width:0}
  .etape-tete{display:flex;align-items:center;gap:10px}
  .etape-tete .n{font-stretch:125%;font-weight:800;font-size:.85rem;color:#fff;background:var(--acc);border:2px solid var(--trait);
    border-radius:8px;width:28px;height:28px;display:grid;place-items:center}
  .etape-tete h2{font-stretch:125%;font-weight:800;font-size:1.02rem;margin:0}
  body:not(.comptable) .etape-tete, body:not(.comptable) .seul-comptable{display:none!important}
  .aide{color:var(--muted);font-size:.84rem;margin:0}

  /* Zones de dépôt */
  .zone{border:2.5px dashed var(--trait);border-radius:var(--r);padding:34px 18px;text-align:center;cursor:pointer;
    background:var(--surface2);transition:background .15s,transform .15s}
  .zone:hover,.zone.survol{background:var(--pri-doux)}
  .zone.petite{padding:18px 14px}
  .zone .ico{width:48px;height:48px;margin:0 auto 10px;border-radius:12px;display:grid;place-items:center;
    background:var(--acc);border:2px solid var(--trait);box-shadow:var(--ombre-p);color:#fff}
  .zone .gros{font-weight:700;font-size:1.02rem}
  .zone .petit{color:var(--muted);font-size:.84rem;margin-top:3px}
  ul.fichiers{list-style:none;padding:0;margin:0;display:grid;gap:6px}
  ul.fichiers li,.puce{display:flex;align-items:center;gap:10px;padding:8px 11px;border:2px solid var(--trait);
    border-radius:10px;font-size:.88rem;background:var(--surface);min-width:0}
  ul.fichiers li span.nom,.puce span.nom{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
  .x{margin-left:auto;border:0;background:transparent;color:var(--muted);cursor:pointer;font-weight:700;
    padding:2px 7px;border-radius:6px;flex:0 0 auto}
  .x:hover{background:var(--acc-doux);color:var(--acc)}
  .ext{font-family:var(--f-chiffre);font-size:.7rem;color:var(--ink);background:var(--pri-doux);border-radius:5px;padding:0 5px}

  /* Plan comptable */
  .plan-etat{border:2px solid var(--trait);border-radius:var(--r);padding:12px 14px;display:grid;gap:8px;background:var(--surface2)}
  .plan-etat.fourni{background:var(--pri-doux);box-shadow:var(--ombre-p)}
  .plan-etat .ligne{display:flex;align-items:center;gap:8px;font-size:.9rem;flex-wrap:wrap}
  .plan-etat b{font-weight:700}
  .comptes{display:flex;flex-wrap:wrap;gap:5px}
  .compte{font-size:.76rem;border:1.5px solid var(--trait);background:var(--surface);border-radius:999px;padding:2px 8px;
    max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .compte code{font-family:var(--f-chiffre);color:var(--pri)}
  .lien{border:0;background:none;color:var(--pri);font-weight:700;cursor:pointer;padding:0;font-size:.84rem;text-decoration:underline;text-underline-offset:2px}
  .tous{max-height:180px;overflow:auto;border-top:1.5px solid var(--line);padding-top:8px;font-size:.82rem;display:grid;gap:3px}
  .tous div{display:flex;gap:8px} .tous code{font-family:var(--f-chiffre);color:var(--pri);min-width:3.5em}

  /* Options */
  .option{display:flex;gap:12px;align-items:flex-start;border:2px solid var(--trait);border-radius:12px;padding:11px 13px;
    cursor:pointer;background:var(--surface)}
  .option:hover{background:var(--surface2)}
  .option b{display:block;font-size:.9rem}
  .option .aide{font-size:.8rem}
  .option svg{flex:0 0 20px;width:20px;height:20px;color:var(--pri);margin-top:1px}
  .option.continuer{border-style:dashed;padding:9px 12px} .option.continuer.choisi{border-style:solid}
  .option.choisi{background:var(--pri-doux);box-shadow:var(--ombre-p)}
  .option .nomfic{font-size:.8rem;color:var(--pri-fort);font-weight:700;word-break:break-all}
  .interrupteur{appearance:none;flex:0 0 40px;width:40px;height:22px;border-radius:999px;background:var(--line);position:relative;
    cursor:pointer;margin:1px 0 0;border:2px solid var(--trait);transition:background .15s}
  .interrupteur::after{content:"";position:absolute;top:1px;left:1px;width:16px;height:16px;border-radius:50%;background:#fff;
    border:1.5px solid var(--trait);transition:transform .15s}
  .interrupteur:checked{background:var(--pri)} .interrupteur:checked::after{transform:translateX(18px)}

  /* Action */
  .actions{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin-top:20px;padding-top:18px;border-top:2px dashed var(--line)}
  .btn{font-family:var(--f-titre);font-stretch:125%;font-weight:800;border:2.5px solid var(--trait);border-radius:14px;padding:13px 22px;cursor:pointer;
    display:inline-flex;align-items:center;gap:8px;text-decoration:none;font-size:.98rem;transition:transform .1s,box-shadow .1s}
  .btn.pri{background:var(--acc);color:#fff;box-shadow:4px 4px 0 var(--trait)}
  .btn.pri:hover{transform:translate(-1px,-1px);box-shadow:5px 5px 0 var(--trait)}
  .btn.pri:active{transform:translate(3px,3px);box-shadow:1px 1px 0 var(--trait)}
  .btn.pri:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:2px 2px 0 var(--trait)}
  .btn.sec{background:var(--surface);color:var(--ink);box-shadow:var(--ombre-p)} .btn.sec:hover{background:var(--pri-doux)}
  .etat{color:var(--muted);font-size:.9rem}
  .spin{width:16px;height:16px;border:3px solid var(--line);border-top-color:var(--acc);border-radius:50%;
    animation:t .8s linear infinite;display:inline-block;vertical-align:-3px}
  @keyframes t{to{transform:rotate(360deg)}}
  @media (prefers-reduced-motion:reduce){.spin{animation:none}*{transition:none!important}}
  .quota-note{color:var(--muted);font-size:.84rem;margin-left:auto}
  .quota-note b{color:var(--ink)}
  .err{color:var(--warn);font-size:.88rem;margin-top:12px;font-weight:600}
  .err:empty{display:none}

  /* Résultats */
  .resultats{margin-top:26px;background:var(--surface);border:2.5px solid var(--trait);border-radius:20px;box-shadow:var(--ombre);padding:22px}
  .res-tete{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;flex-wrap:wrap}
  .res-tete h2{font-stretch:125%;font-weight:800;font-size:1.25rem;margin:0}
  .res-tete .lie{color:var(--muted);font-size:.84rem}
  .tels{display:flex;gap:10px;flex-wrap:wrap}
  .chiffres{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:18px 0}
  .chiffre{border:2px solid var(--trait);border-radius:12px;padding:12px 14px;background:var(--surface2);box-shadow:var(--ombre-p)}
  .chiffre span{display:block;font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);font-weight:700}
  .chiffre b{font-family:var(--f-chiffre);font-size:1.15rem;font-weight:500;font-variant-numeric:tabular-nums}
  .verif{display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-radius:12px;padding:11px 14px;margin-bottom:12px;
    background:var(--pri-doux);border:2px solid var(--trait);font-size:.9rem}
  .verif .pill{margin-left:auto;font-size:.78rem;font-weight:700;border-radius:999px;padding:3px 10px;background:var(--surface);
    border:1.5px solid var(--trait);color:var(--pri-fort);white-space:nowrap}
  .verif .pill.encours{color:var(--muted)}
  .table-wrap{overflow-x:auto;border:2px solid var(--trait);border-radius:12px}
  table{width:100%;border-collapse:collapse;font-size:.88rem;min-width:760px}
  th,td{padding:10px 12px;text-align:left;border-bottom:1.5px solid var(--line);vertical-align:top}
  thead th{background:var(--pri);color:#fff;font-weight:700;font-size:.74rem;letter-spacing:.06em;text-transform:uppercase;border-bottom:2px solid var(--trait)}
  th.num{text-align:right} td.date{white-space:nowrap}
  td.num{text-align:right;font-family:var(--f-chiffre);font-variant-numeric:tabular-nums;white-space:nowrap}
  tbody tr:last-child td{border-bottom:0}
  .sous{color:var(--muted);font-size:.78rem;display:block;margin-top:2px}
  .al{color:var(--warn);font-size:.78rem;display:block;margin-top:2px}
  select.choix{font:inherit;font-size:.86rem;width:100%;min-width:220px;padding:7px 9px;border:2px solid var(--trait);
    border-radius:9px;background:var(--surface);color:var(--ink)}
  select.choix.vide{background:var(--warn-doux)}
  .marque{font-size:.74rem;color:var(--muted);display:block;margin-top:3px}
  .marque.modif{color:var(--pri-fort);font-weight:700}
  .cumul{margin-top:14px;font-size:.88rem;line-height:1.7;color:var(--muted)}
  .cumul b{color:var(--ink);font-family:var(--f-chiffre);font-weight:500}

  /* Offre */
  .offre{margin-top:30px;border-radius:20px;padding:24px;display:flex;gap:20px;align-items:center;justify-content:space-between;
    flex-wrap:wrap;background:var(--pri);color:#fff;border:2.5px solid var(--trait);box-shadow:var(--ombre)}
  .offre>div{flex:1 1 320px;min-width:0}
  .offre .surtitre{color:#FFD9C9}
  .offre h3{font-stretch:125%;font-weight:800;font-size:1.25rem;margin:4px 0 6px}
  .offre p{margin:0;opacity:.9;font-size:.92rem}
  .offre .btn{background:#fff;color:#0B2B24;box-shadow:4px 4px 0 var(--trait)}
  .pied{color:var(--muted);font-size:.8rem;text-align:center;padding-block:28px 40px}

  @media (max-width:860px){
    body.comptable .etapes{grid-template-columns:1fr}
    .hero{grid-template-columns:1fr;gap:6px}
    .illus{order:-1;height:220px;max-width:380px;width:100%;margin:0 auto}
    .illus .recu{width:42%;height:180px;left:2%;top:10px;padding:12px}
    .illus .feuille{width:52%;right:3%;bottom:14px}
    .illus .fleche{left:auto;right:30%;top:40px;width:46px;height:46px}
  }
  @media (max-width:640px){
    /* Sur téléphone, chaque pièce devient une fiche : le compte reste visible. */
    table{min-width:0} thead{display:none}
    tbody tr{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px;padding:14px;border-bottom:1.5px solid var(--line)}
    tbody tr:last-child{border-bottom:0}
    td{border:0;padding:0;min-width:0} td.num{text-align:left} td.num:empty{display:none}
    td::before{content:attr(data-l);display:block;font:700 .68rem var(--f-texte);color:var(--muted);
      text-transform:uppercase;letter-spacing:.06em}
    td.plein{grid-column:1/-1} select.choix{min-width:0}
    .tels{width:100%} .tels .btn{flex:1 1 auto;justify-content:center}
    .pas-liste{gap:7px} .pas{padding:10px 9px} .pas b{font-size:.82rem} .pas span:last-child{font-size:.76rem}
  }
  @media (max-width:560px){
    .modes{grid-template-columns:1fr 1fr} .mode{padding:12px} .mode span{display:none} .mode svg{display:none}
    .statut span:last-child{display:none} .corps{padding:16px} .quota-note{margin-left:0}
    #btn{width:100%;justify-content:center}
  }
</style>
</head>
<body>
<header class="conteneur barre">
  <a class="logo" href="/">
    <svg viewBox="0 0 44 44" width="34" height="34" aria-hidden="true">
      <rect x="2" y="2" width="40" height="40" rx="12" fill="var(--pri)"/>
      <path d="M12 11h10.5a2 2 0 0 1 2 2v16l-2.1-1.5-2.1 1.5-2.1-1.5-2.1 1.5-2.1-1.5-1.9 1.35V13a2 2 0 0 1 2-2z" fill="#fff"/>
      <rect x="14.4" y="15.2" width="7.5" height="1.7" rx=".85" fill="var(--pri)"/>
      <rect x="14.4" y="18.6" width="5.5" height="1.7" rx=".85" fill="var(--pri)"/>
      <rect x="24" y="22" width="8.5" height="8.5" rx="2" fill="var(--acc)"/>
      <path d="M28.25 22v8.5M24 26.25h8.5" stroke="#fff" stroke-width="1.1"/>
    </svg>
    <span>Facturo</span>
  </a>
  <div class="droite">
    <span class="statut"><span id="pastille" class="pastille"></span><span id="txtStatut" data-i18n="statut_check">Vérification…</span></span>
    <div class="lang">
      <button id="fr" class="actif" onclick="setLang('fr')">FR</button>
      <button id="en" onclick="setLang('en')">EN</button>
    </div>
  </div>
</header>

<main class="conteneur">
  <section class="hero">
    <div class="hero-texte">
      <span class="surtitre" data-i18n="surtitre">Bilingue · TPS, TVQ, TVH · QuickBooks Online</span>
      <h1 data-i18n="hero">Reçus in, Excel out.</h1>
      <p data-i18n="sous">Fini la saisie à la main. Prenez vos reçus et factures en photo : Facturo en fait un fichier Excel propre, avec la TPS, la TVQ ou la TVH séparées, prêt pour votre comptable.</p>
      <div class="pas-liste">
        <div class="pas"><span class="n">1</span><b data-i18n="p1t">Photo ou PDF</b><span data-i18n="p1d">Un reçu ou tout un lot, en français ou en anglais.</span></div>
        <div class="pas"><span class="n">2</span><b data-i18n="p2t">Lecture auto</b><span data-i18n="p2d">Fournisseur, date, taxes, total.</span></div>
        <div class="pas"><span class="n">3</span><b data-i18n="p3t">Excel prêt</b><span data-i18n="p3d">À garder, à continuer ou à donner au comptable.</span></div>
      </div>
    </div>
    <div class="illus" aria-hidden="true">
      <div class="recu"><i></i><i></i><i style="width:80%"></i><i style="width:88%"></i><i style="width:60%"></i>
        <div class="tx"><span>TPS+TVQ</span><em>2,93</em></div></div>
      <div class="feuille"><span class="t" data-i18n="ill_f">Fourn.</span><span class="t" data-i18n="ill_t">Taxes</span><span class="t">Total</span>
        <span>Boul.</span><span>2,93</span><span>22,43</span><span>Garage</span><span>85,21</span><span>654,21</span></div>
      <div class="fleche"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></div>
    </div>
  </section>

  <section class="atelier">
    <div class="modes" role="tablist">
      <button class="mode" id="modeSimple" role="tab" aria-selected="true" onclick="choisirMode('simple')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 3h9l3 3v15l-2-1.3-2 1.3-2-1.3-2 1.3-2-1.3L6 21z"/><path d="M9 9h6M9 13h4"/></svg>
        <div><b data-i18n="mode_simple">Usage simple</b><span data-i18n="mode_simple_d">Déposez vos reçus, téléchargez l'Excel.</span></div>
      </button>
      <button class="mode" id="modeComptable" role="tab" aria-selected="false" onclick="choisirMode('comptable')">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 19V5M9 8h6M9 12h6"/></svg>
        <div><b data-i18n="mode_comptable">Mode comptable</b><span data-i18n="mode_comptable_d">Votre plan comptable, vérification des comptes, QuickBooks, relevé bancaire.</span></div>
      </button>
    </div>

    <div class="corps">
      <div class="etapes">
        <!-- 1. Plan comptable (mode comptable) -->
        <div class="etape seul-comptable">
          <div class="etape-tete"><span class="n">1</span><h2 data-i18n="e_plan">Votre plan comptable</h2></div>
          <p class="aide" data-i18n="plan_aide">Excel ou CSV avec une colonne code (ex. 5210) et une colonne nom. Chaque pièce sera rangée dans un compte de CE plan, et vous pourrez corriger avant de télécharger.</p>
          <div id="planZone" class="zone petite" tabindex="0" role="button">
            <div class="gros" data-i18n="plan_depot">Déposer mon plan comptable</div>
            <div class="petit">.xlsx · .csv</div>
            <input id="planInput" type="file" accept=".xlsx,.csv" class="masque">
          </div>
          <div id="planEtat" class="plan-etat"></div>
        </div>

        <!-- 2. Pièces (toujours) -->
        <div class="etape">
          <div class="etape-tete"><span class="n">2</span><h2 data-i18n="e_pieces">Vos pièces</h2></div>
          <div id="zone" class="zone" tabindex="0" role="button">
            <div class="ico"><svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 16V4M8 8l4-4 4 4"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg></div>
            <div class="gros" data-i18n="drop_gros">Glissez vos reçus et factures ici</div>
            <div class="petit" data-i18n="drop_petit">ou cliquez pour choisir · photos (JPG, PNG, HEIC) et PDF, plusieurs à la fois</div>
            <input id="input" type="file" multiple accept=".pdf,image/*,.heic,.heif" class="masque">
          </div>
          <ul id="liste" class="fichiers"></ul>
          <label class="option continuer" id="optBase">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M3 9h18M9 21V9"/></svg>
            <div><b data-i18n="base_btn">Continuer un Excel existant</b><span class="aide" data-i18n="base_hint">Facultatif : déposez votre dernier fichier Facturo, les nouvelles pièces s'ajoutent à la suite.</span><span id="baseNom" class="nomfic"></span></div>
            <input id="baseInput" type="file" accept=".xlsx" class="masque">
          </label>
        </div>

        <!-- 3. Options (mode comptable) -->
        <div class="etape seul-comptable">
          <div class="etape-tete"><span class="n">3</span><h2 data-i18n="e_options">Options</h2></div>
          <label class="option" id="optQbo">
            <input id="qboInput" type="checkbox" class="interrupteur">
            <div><b data-i18n="qbo_coche">Fichier d'import QuickBooks Online</b><span class="aide" data-i18n="qbo_hint">CSV de factures fournisseurs, en plus de l'Excel, avec le compte et le code de taxe de chaque ligne.</span></div>
          </label>
          <label class="option" id="optReleve">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 10l9-6 9 6M5 10v8M19 10v8M9 10v8M15 10v8M3 20h18"/></svg>
            <div><b data-i18n="releve_btn">Rapprocher un relevé bancaire</b><span class="aide" data-i18n="releve_hint">CSV ou OFX de votre banque : repère les dépenses sans pièce.</span><span id="releveNom" class="nomfic"></span></div>
            <input id="releveInput" type="file" accept=".csv,.ofx,.qfx" class="masque">
          </label>
        </div>
      </div>

      <div class="actions">
        <button id="btn" class="btn pri" disabled data-i18n="btn">Convertir en Excel</button>
        <span id="etat" class="etat"></span>
        <span id="quotaNote" class="quota-note masque"></span>
      </div>
      <div id="erreurs" class="err"></div>
    </div>
  </section>

  <section id="resultats" class="resultats masque" aria-live="polite">
    <div class="res-tete">
      <div><h2 id="titreRes" data-i18n="res">Résultat</h2><span id="ajouteNote" class="lie"></span></div>
      <div class="tels">
        <a id="tel" class="btn pri" href="#"><span aria-hidden="true">↓</span><span data-i18n="tel">Télécharger l'Excel</span></a>
        <a id="telQbo" class="btn sec masque" href="#"><span aria-hidden="true">↓</span><span data-i18n="tel_qbo">CSV QuickBooks Online</span></a>
      </div>
    </div>
    <div id="chiffres" class="chiffres"></div>
    <div id="verif" class="verif masque">
      <span data-i18n="verif_txt">Vérifiez le compte proposé pour chaque pièce. Chaque changement met à jour l'Excel et le fichier QuickBooks.</span>
      <span id="verifEtat" class="pill"></span>
    </div>
    <div class="table-wrap"><table id="tab"></table></div>
    <div id="cumul" class="cumul masque"></div>
  </section>

  <section id="offre" class="offre">
    <div>
      <span class="surtitre" data-i18n="offre_tag">Facturo Pro</span>
      <h3 data-i18n="offre_titre">Une version sur mesure pour votre entreprise</h3>
      <p data-i18n="offre_texte">Plus de volume, votre plan comptable et vos fichiers Excel existants, un format adapté à votre logiciel comptable, un accès privé pour votre équipe. Conçu et accompagné par un consultant bilingue basé à Ottawa.</p>
    </div>
    <a id="offreLien" class="btn" href="#" data-i18n="offre_btn">Discuter de mon besoin</a>
  </section>
</main>
<footer class="conteneur pied"><span data-i18n="pied">Facturo · vos données ne servent qu'à produire votre tableur, et les fichiers sont effacés après une heure.</span></footer>

<script>
const I18N = {
  fr:{ statut_check:"Vérification…", statut_on:"Service en ligne", statut_off:"Service indisponible",
    surtitre:"Bilingue · TPS, TVQ, TVH · QuickBooks Online",
    hero:"Reçus in, Excel out.", ill_f:"Fourn.", ill_t:"Taxes",
    sous:"Fini la saisie à la main. Prenez vos reçus et factures en photo : Facturo en fait un fichier Excel propre, avec la TPS, la TVQ ou la TVH séparées, prêt pour votre comptable.",
    mode_simple:"Usage simple", mode_simple_d:"Déposez vos reçus, téléchargez l'Excel.",
    mode_comptable:"Mode comptable", mode_comptable_d:"Votre plan comptable, vérification des comptes, QuickBooks, relevé bancaire.",
    e_plan:"Votre plan comptable", e_pieces:"Vos pièces", e_options:"Options",
    plan_aide:"Excel ou CSV avec une colonne code (ex. 5210) et une colonne nom. Chaque pièce sera rangée dans un compte de CE plan, et vous pourrez corriger avant de télécharger.",
    plan_depot:"Déposer mon plan comptable", plan_lecture:"Lecture du plan…",
    plan_std:"Plan standard Facturo utilisé", plan_std_d:"comptes inspirés du formulaire T2125. Déposez le vôtre pour le remplacer.",
    plan_ok:"comptes reconnus dans", plan_voir:"Voir les comptes", plan_cacher:"Masquer", plan_retirer:"Retirer",
    drop_gros:"Glissez vos reçus et factures ici", drop_petit:"ou cliquez pour choisir · photos (JPG, PNG, HEIC) et PDF, plusieurs à la fois",
    btn:"Convertir en Excel", btn_comptable:"Lire les pièces et proposer les comptes", traitement:"Lecture en cours…",
    qbo_coche:"Fichier d'import QuickBooks Online", qbo_hint:"CSV de factures fournisseurs, en plus de l'Excel, avec le compte et le code de taxe de chaque ligne.",
    base_btn:"Continuer un Excel existant", base_hint:"Facultatif : déposez votre dernier fichier Facturo, les nouvelles pièces s'ajoutent à la suite.",
    releve_btn:"Rapprocher un relevé bancaire", releve_hint:"CSV ou OFX de votre banque : repère les dépenses sans pièce.",
    choisi:"Fichier :", ajoute_note:"Nouvelles lignes ajoutées à votre fichier.",
    res:"Résultat", tel:"Télécharger l'Excel", tel_qbo:"CSV QuickBooks Online",
    qbo_aide:"Factures fournisseurs à importer dans QuickBooks Online (Paramètres > Importer des données > Factures) : montants hors taxes, dates JJ/MM/AAAA.",
    verif_txt:"Vérifiez le compte proposé pour chaque pièce. Chaque changement met à jour l'Excel et le fichier QuickBooks.",
    v_ajour:"Excel à jour", v_encours:"Mise à jour…", v_erreur:"Échec de la mise à jour, réessayez",
    suggere:"proposé par Facturo", modifie:"choisi par vous", a_choisir:"— Choisir un compte —",
    k_pieces:"Pièces", k_total:"Total TTC", k_taxes:"Taxes", k_ht:"Hors taxes",
    th_f:"Fournisseur", th_d:"Date", th_ht:"HT", th_tps:"TPS", th_tvq:"TVQ", th_tx:"Taxes", th_ttc:"TTC", th_cat:"Compte",
    regroupe:"regroupé avec", pieces:"pièce(s)", aucune:"Aucune donnée n'a pu être extraite.", reseau:"Erreur réseau : ",
    quota:"Limite de la démo atteinte pour aujourd'hui. Revenez demain !",
    gratuit:"Version gratuite", par_jour:"fichiers par jour", restants:"restant(s) aujourd'hui", voir_pro:"Facturo Pro",
    p1t:"Photo ou PDF", p1d:"Un reçu ou tout un lot, en français ou en anglais.",
    p2t:"Lecture auto", p2d:"Fournisseur, date, taxes, total.",
    p3t:"Excel prêt", p3d:"À garder, à continuer ou à donner au comptable.",
    offre_tag:"Facturo Pro", offre_titre:"Une version sur mesure pour votre entreprise",
    offre_texte:"Plus de volume, votre plan comptable et vos fichiers Excel existants, un format adapté à votre logiciel comptable, un accès privé pour votre équipe. Conçu et accompagné par un consultant bilingue basé à Ottawa.",
    offre_btn:"Discuter de mon besoin", mail_sujet:"Facturo Pro — demande d'information",
    mail_corps:"Bonjour,\n\nJ'ai essayé Facturo et j'aimerais en savoir plus sur une version adaptée à mon entreprise.\n\nEntreprise :\nVolume approximatif (factures par mois) :\nLogiciel comptable utilisé :\n\nMerci !",
    pied:"Facturo · vos données ne servent qu'à produire votre tableur, et les fichiers sont effacés après une heure.",
    rap_titre:"Relevé bancaire", rap_ok:"transaction(s) rapprochée(s)", rap_sans:"dépense(s) sans pièce",
    rap_abs:"facture(s) absente(s) du relevé", rap_voir:"détail dans l'onglet Rapprochement", rap_cat:"catégorie suggérée pour",
    cumul:"Total du fichier", dont_taxes:"dont taxes", nb_fact:"facture(s) depuis le début" },
  en:{ statut_check:"Checking…", statut_on:"Service online", statut_off:"Service unavailable",
    surtitre:"Bilingual · GST, QST, HST · QuickBooks Online",
    hero:"Receipts in, Excel out.", ill_f:"Vendor", ill_t:"Tax",
    sous:"No more typing receipts by hand. Snap your receipts and invoices: Facturo turns them into a clean Excel file, with GST, QST or HST split out, ready for your accountant.",
    mode_simple:"Simple use", mode_simple_d:"Drop your receipts, download the Excel.",
    mode_comptable:"Accountant mode", mode_comptable_d:"Your chart of accounts, account review, QuickBooks, bank statement.",
    e_plan:"Your chart of accounts", e_pieces:"Your documents", e_options:"Options",
    plan_aide:"Excel or CSV with a code column (e.g. 5210) and a name column. Each document is filed under an account from THIS chart, and you can correct it before downloading.",
    plan_depot:"Drop my chart of accounts", plan_lecture:"Reading the chart…",
    plan_std:"Facturo standard chart in use", plan_std_d:"accounts based on form T2125. Drop yours to replace it.",
    plan_ok:"accounts recognized in", plan_voir:"See the accounts", plan_cacher:"Hide", plan_retirer:"Remove",
    drop_gros:"Drag your receipts and invoices here", drop_petit:"or click to choose · photos (JPG, PNG, HEIC) and PDF, several at once",
    btn:"Convert to Excel", btn_comptable:"Read the documents and suggest accounts", traitement:"Reading…",
    qbo_coche:"QuickBooks Online import file", qbo_hint:"Bills CSV, in addition to the Excel, with the account and tax code on each line.",
    base_btn:"Continue an existing Excel", base_hint:"Optional: drop your latest Facturo file, the new documents are added after the existing ones.",
    releve_btn:"Reconcile a bank statement", releve_hint:"CSV or OFX from your bank: flags expenses without a receipt.",
    choisi:"File:", ajoute_note:"New rows appended to your file.",
    res:"Result", tel:"Download the Excel", tel_qbo:"QuickBooks Online CSV",
    qbo_aide:"Bills to import into QuickBooks Online (Settings > Import data > Bills): amounts exclude tax, dates DD/MM/YYYY.",
    verif_txt:"Check the suggested account for each document. Every change updates the Excel and the QuickBooks file.",
    v_ajour:"Excel up to date", v_encours:"Updating…", v_erreur:"Update failed, please try again",
    suggere:"suggested by Facturo", modifie:"chosen by you", a_choisir:"— Choose an account —",
    k_pieces:"Documents", k_total:"Total", k_taxes:"Tax", k_ht:"Before tax",
    th_f:"Vendor", th_d:"Date", th_ht:"Net", th_tps:"GST", th_tvq:"QST", th_tx:"Tax", th_ttc:"Total", th_cat:"Account",
    regroupe:"merged with", pieces:"document(s)", aucune:"No data could be extracted.", reseau:"Network error: ",
    quota:"Today's demo limit has been reached. Please come back tomorrow!",
    gratuit:"Free version", par_jour:"files per day", restants:"left today", voir_pro:"Facturo Pro",
    p1t:"Photo or PDF", p1d:"One receipt or a whole batch, in English or French.",
    p2t:"Auto reading", p2d:"Vendor, date, taxes, total.",
    p3t:"Excel ready", p3d:"Keep it, add to it, or hand it to your accountant.",
    offre_tag:"Facturo Pro", offre_titre:"A custom version for your business",
    offre_texte:"More volume, your chart of accounts and existing Excel files, output tailored to your accounting software, private access for your team. Built and supported by a bilingual consultant based in Ottawa.",
    offre_btn:"Discuss my needs", mail_sujet:"Facturo Pro — information request",
    mail_corps:"Hello,\n\nI tried Facturo and would like to learn more about a version tailored to my business.\n\nCompany:\nApproximate volume (invoices per month):\nAccounting software used:\n\nThank you!",
    pied:"Facturo · your data is only used to produce your spreadsheet, and files are deleted after one hour.",
    rap_titre:"Bank statement", rap_ok:"transaction(s) matched", rap_sans:"expense(s) without a receipt",
    rap_abs:"invoice(s) not on the statement", rap_voir:"details in the Bank reconciliation sheet", rap_cat:"suggested category for",
    cumul:"File total", dont_taxes:"incl. tax", nb_fact:"invoice(s) since the start" }
};
function lire(k){try{return localStorage.getItem(k);}catch(e){return null;}}
function ecrire(k,v){try{localStorage.setItem(k,v);}catch(e){}}
let LANG = lire('facturo_lang') || (navigator.language||'fr').slice(0,2);
if(LANG!=='en') LANG='fr';
let MODE = location.hash==='#comptable' ? 'comptable' : (lire('facturo_mode')||'simple');
let dernier=null, choixInitiaux={};

function T(k){return (I18N[LANG]&&I18N[LANG][k])||k;}
// Tout texte venant d'un fichier (nom, contenu lu) est échappé avant affichage.
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
const $=id=>document.getElementById(id);

function setLang(l){
  LANG=l; ecrire('facturo_lang',l); document.documentElement.lang=l;
  $('fr').classList.toggle('actif',l==='fr'); $('en').classList.toggle('actif',l==='en');
  document.querySelectorAll('[data-i18n]').forEach(e=>{const k=e.getAttribute('data-i18n');if(I18N[l][k])e.textContent=I18N[l][k];});
  if(!planComptes) chargerPlanStandard();
  majStatut(); majBouton(); rendrePlan(); rendreOptions(); majOffre(); if(dernier) afficher(dernier,false);
}
function choisirMode(m){
  MODE=m; ecrire('facturo_mode',m);
  document.body.classList.toggle('comptable',m==='comptable');
  $('modeSimple').setAttribute('aria-selected',m==='simple');
  $('modeComptable').setAttribute('aria-selected',m==='comptable');
  if(m==='comptable'&&location.hash!=='#comptable') history.replaceState(null,'','#comptable');
  if(m==='simple'&&location.hash) history.replaceState(null,'',location.pathname);
  majBouton(); if(dernier) afficher(dernier,false);
}

let statutPret=null, quota=null, restant=null, contact='';
function majStatut(){
  const p=$('pastille'), t=$('txtStatut');
  if(statutPret===null){t.textContent=T('statut_check');p.className='pastille';}
  else if(statutPret){t.textContent=T('statut_on');p.className='pastille on';}
  else{t.textContent=T('statut_off');p.className='pastille off';}
}
fetch('/api/statut').then(r=>r.json()).then(s=>{statutPret=!!s.pret;quota=s.quota;restant=s.restant;contact=s.contact||'';majStatut();majOffre();}).catch(()=>{statutPret=false;majStatut();});
function majOffre(){
  const q=$('quotaNote');
  if(quota!==null&&quota!==undefined){
    q.classList.remove('masque');
    q.innerHTML=`${T('gratuit')} · ${quota} ${T('par_jour')} · <b>${restant}</b> ${T('restants')} · <a href="#offre">${T('voir_pro')}</a>`;
  }
  if(contact) $('offreLien').href=`mailto:${contact}?subject=${encodeURIComponent(T('mail_sujet'))}&body=${encodeURIComponent(T('mail_corps'))}`;
}

// --- Pièces ---
let fichiers=[];
const zone=$('zone'), input=$('input');
function brancherZone(z, inp, rappel){
  z.addEventListener('click',()=>inp.click());
  z.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();inp.click();}});
  ['dragover','dragenter'].forEach(ev=>z.addEventListener(ev,e=>{e.preventDefault();z.classList.add('survol');}));
  ['dragleave','drop'].forEach(ev=>z.addEventListener(ev,e=>{e.preventDefault();z.classList.remove('survol');}));
  z.addEventListener('drop',e=>rappel(e.dataTransfer.files));
  inp.addEventListener('change',()=>{rappel(inp.files);inp.value='';});
}
brancherZone(zone,input,fl=>{for(const f of fl)fichiers.push(f);rendreListe();});
function rendreListe(){
  const ul=$('liste'); ul.innerHTML='';
  fichiers.forEach((f,i)=>{
    const ext=((f.name||'').split('.').pop()||'').slice(0,5);
    const li=document.createElement('li');
    li.innerHTML=`<span class="ext">${esc(ext)}</span><span class="nom">${esc(f.name)}</span><button class="x" aria-label="✕">✕</button>`;
    li.querySelector('.x').onclick=()=>{fichiers.splice(i,1);rendreListe();};
    ul.appendChild(li);
  });
  majBouton();
}
function majBouton(){
  const b=$('btn'); b.textContent=MODE==='comptable'?T('btn_comptable'):T('btn');
  b.disabled=!fichiers.length&&!(MODE==='comptable'&&releveFichier);
}

// --- Plan comptable ---
let planFichier=null, planComptes=null, planStandard=[], planOuvert=false;
function chargerPlanStandard(){
  fetch('/api/plan-standard?langue='+LANG).then(r=>r.json()).then(d=>{planStandard=d.comptes||[];rendrePlan();}).catch(()=>{});
}
brancherZone($('planZone'),$('planInput'),async fl=>{
  const f=fl[0]; if(!f) return;
  $('planEtat').innerHTML=`<div class="ligne"><span class="spin"></span> ${T('plan_lecture')}</div>`;
  const fd=new FormData(); fd.append('plan',f,f.name); fd.append('langue',LANG);
  try{
    const r=await fetch('/api/plan',{method:'POST',body:fd}); const d=await r.json();
    if(!r.ok){planFichier=null;planComptes=null;rendrePlan();$('erreurs').textContent='⚠ '+(d.erreur||'');return;}
    planFichier=f; planComptes=d.comptes; planOuvert=false; $('erreurs').textContent='';
  }catch(e){planFichier=null;planComptes=null;$('erreurs').textContent=T('reseau')+e;}
  rendrePlan();
});
function rendrePlan(){
  const el=$('planEtat'), comptes=planComptes||planStandard;
  el.classList.toggle('fourni',!!planComptes);
  $('planZone').classList.toggle('masque',!!planComptes);
  const tete=planComptes
    ? `<div class="ligne"><b>✓ ${planComptes.length} ${T('plan_ok')}</b><span class="nom">${esc(planFichier.name)}</span><button class="x" id="planRetirer" title="${T('plan_retirer')}">✕</button></div>`
    : `<div class="ligne"><b>${T('plan_std')}</b><span class="aide">${comptes.length} ${T('plan_std_d')}</span></div>`;
  const apercu=comptes.slice(0,6).map(c=>`<span class="compte">${c.id!==c.libelle?`<code>${esc(c.id)}</code> `:''}${esc(c.libelle.replace(c.id+' · ',''))}</span>`).join('');
  const tous=planOuvert?`<div class="tous">${comptes.map(c=>`<div><code>${c.id!==c.libelle?esc(c.id):''}</code><span>${esc(c.libelle.replace(c.id+' · ',''))}</span></div>`).join('')}</div>`:'';
  el.innerHTML=tete+`<div class="comptes">${apercu}</div><div><button class="lien" id="planVoir">${planOuvert?T('plan_cacher'):T('plan_voir')+' ('+comptes.length+')'}</button></div>`+tous;
  $('planVoir').onclick=()=>{planOuvert=!planOuvert;rendrePlan();};
  const x=$('planRetirer'); if(x) x.onclick=()=>{planFichier=null;planComptes=null;rendrePlan();};
}

// --- Options ---
let baseExcel=null, releveFichier=null;
$('baseInput').addEventListener('change',e=>{baseExcel=e.target.files[0]||null;e.target.value='';rendreOptions();});
$('releveInput').addEventListener('change',e=>{releveFichier=e.target.files[0]||null;e.target.value='';rendreOptions();});
const qboInput=$('qboInput');
if(lire('facturo_qbo')==='1') qboInput.checked=true;
qboInput.addEventListener('change',()=>{ecrire('facturo_qbo',qboInput.checked?'1':'0');rendreOptions();});
function nomFichier(el,f,vider){
  el.innerHTML=f?`${T('choisi')} ${esc(f.name)} <button class="x" aria-label="✕">✕</button>`:'';
  if(f) el.querySelector('.x').onclick=e=>{e.preventDefault();e.stopPropagation();vider();};
}
function rendreOptions(){
  nomFichier($('baseNom'),baseExcel,()=>{baseExcel=null;rendreOptions();});
  nomFichier($('releveNom'),releveFichier,()=>{releveFichier=null;rendreOptions();});
  $('optBase').classList.toggle('choisi',!!baseExcel);
  $('optReleve').classList.toggle('choisi',!!releveFichier);
  $('optQbo').classList.toggle('choisi',qboInput.checked);
  majBouton();
}

// --- Envoi ---
$('btn').onclick=async()=>{
  const compta=MODE==='comptable';
  if(!fichiers.length&&!(compta&&releveFichier)) return;
  const b=$('btn'), err=$('erreurs'); b.disabled=true; err.textContent=''; $('resultats').classList.add('masque');
  $('etat').innerHTML='<span class="spin"></span> '+T('traitement');
  const fd=new FormData(); fichiers.forEach(f=>fd.append('fichiers',f,f.name));
  // L'Excel à continuer fait partie des deux usages. En usage simple, les
  // options comptables ne partent pas : elles ne changent jamais cet Excel.
  if(baseExcel) fd.append('base',baseExcel,baseExcel.name);
  if(compta){
    if(planFichier) fd.append('plan',planFichier,planFichier.name);
    if(releveFichier) fd.append('releve',releveFichier,releveFichier.name);
    if(qboInput.checked) fd.append('qbo','1');
  }
  fd.append('langue',LANG);
  try{
    const r=await fetch('/api/extraire',{method:'POST',body:fd});
    if(r.status===429){dernier=null;restant=0;majOffre();err.innerHTML=`${T('quota')} <a href="#offre">${T('voir_pro')} →</a>`;}
    else{dernier=await r.json(); dernier.mode=MODE; choixInitiaux={};
      (dernier.factures||[]).forEach((f,i)=>choixInitiaux[i]=f.compte);
      if(dernier.restant!==undefined){restant=dernier.restant;majOffre();} afficher(dernier);}
  }catch(e){err.textContent=T('reseau')+e;}
  $('etat').textContent=''; majBouton();
};

function fmt(x){return(x===null||x===undefined||x==='')?'':Number(x).toLocaleString(LANG==='fr'?'fr-CA':'en-CA',{minimumFractionDigits:2,maximumFractionDigits:2});}
function somme(l,k){return l.reduce((s,f)=>s+(Number(f[k])||0),0);}
function afficher(d,defiler=true){
  const err=$('erreurs'), res=$('resultats');
  err.innerHTML=(d.erreurs&&d.erreurs.length)?d.erreurs.map(e=>`⚠ ${esc(e.fichier)} : ${esc(e.message)}`).join('<br>'):'';
  if(!d.download_id){res.classList.add('masque');if(!(d.erreurs&&d.erreurs.length))err.textContent=T('aucune');return;}
  const compta=d.mode==='comptable', fs=d.factures||[];
  const devises=[...new Set(fs.map(f=>f.devise).filter(Boolean))], dev=devises.length===1?' '+esc(devises[0]):'';
  $('chiffres').innerHTML=fs.length?[[T('k_pieces'),fs.length],[T('k_ht'),fmt(somme(fs,'total_ht'))+dev],
    [T('k_taxes'),fmt(somme(fs,'tva'))+dev],[T('k_total'),fmt(somme(fs,'total_ttc'))+dev]]
    .map(([k,v])=>`<div class="chiffre"><span>${k}</span><b>${v}</b></div>`).join(''):'';
  $('verif').classList.toggle('masque',!(compta&&fs.length));
  if(compta) majVerif('ajour');
  let h=`<thead><tr><th>${T('th_f')}</th><th>${T('th_d')}</th><th class="num">${T('th_ht')}</th><th class="num">${T('th_tps')}</th>
    <th class="num">${T('th_tvq')}</th><th class="num">${T('th_tx')}</th><th class="num">${T('th_ttc')}</th><th>${T('th_cat')}</th></tr></thead><tbody>`;
  fs.forEach((f,i)=>{
    let extra=f.numero?`<span class="sous">N° ${esc(f.numero)}</span>`:'';
    if(f.pieces_liees&&f.pieces_liees.length) extra+=`<span class="sous">↻ ${T('regroupe')} ${esc(f.pieces_liees.join(', '))}</span>`;
    if(f.alertes&&f.alertes.length) extra+=`<span class="al">⚠ ${esc(f.alertes.join(' | '))}</span>`;
    let cat=esc(f.categorie);
    if(compta&&d.plan&&d.plan.length){
      const opts=(f.compte?'':`<option value="">${T('a_choisir')}</option>`)+d.plan.map(c=>
        `<option value="${esc(c.id)}"${c.id===f.compte?' selected':''}>${esc(c.libelle)}</option>`).join('');
      const modif=choixInitiaux[i]!==f.compte;
      cat=`<select class="choix${f.compte?'':' vide'}" data-i="${i}" aria-label="${T('th_cat')}">${opts}</select>`
        +`<span class="marque${modif?' modif':''}">${f.compte?(modif?T('modifie'):T('suggere')):''}</span>`;
    }
    h+=`<tr><td class="plein" data-l="${T('th_f')}">${esc(f.fournisseur)||'—'}${extra}</td><td class="date" data-l="${T('th_d')}">${esc(f.date)}</td>
        <td class="num" data-l="${T('th_ht')}">${fmt(f.total_ht)}</td><td class="num" data-l="${T('th_tps')}">${fmt(f.tps)}</td><td class="num" data-l="${T('th_tvq')}">${fmt(f.tvq)}</td>
        <td class="num" data-l="${T('th_tx')}">${fmt(f.tva)}</td><td class="num" data-l="${T('th_ttc')}">${fmt(f.total_ttc)} ${esc(f.devise)}</td><td class="plein" data-l="${T('th_cat')}">${cat}</td></tr>`;
  });
  $('tab').innerHTML=fs.length?h+'</tbody>':'';
  $('tab').querySelectorAll('select.choix').forEach(s=>s.addEventListener('change',()=>corriger(+s.dataset.i,s.value)));
  $('titreRes').textContent=fs.length?`${T('res')} · ${fs.length} ${T('pieces')}`:`${T('res')} · ${T('rap_titre')}`;
  $('ajouteNote').textContent=(d.ajoute&&fs.length)?('↳ '+T('ajoute_note')):'';
  const cumul=$('cumul'), tot=d.totaux_fichier||[];
  cumul.innerHTML=tot.map(t=>`${T('cumul')} : <b>${fmt(t.total_ttc)} ${esc(t.devise)}</b> · ${T('dont_taxes')} ${fmt(t.taxes)} ${esc(t.devise)} · ${t.nb} ${T('nb_fact')}`).join('<br>');
  const rap=d.rapprochement;
  if(rap) cumul.innerHTML+=(tot.length?'<br>':'')+`${T('rap_titre')} : <b>${rap.rapprochees}</b> ${T('rap_ok')} · `
    +`<b>${rap.sans_piece}</b> ${T('rap_sans')} (${fmt(rap.montant_sans_piece)}${rap.categorisees?`, ${T('rap_cat')} ${rap.categorisees}`:''}) · <b>${rap.absentes}</b> ${T('rap_abs')} · ${T('rap_voir')}`;
  cumul.classList.toggle('masque',!tot.length&&!rap);
  const q=$('telQbo'); q.classList.toggle('masque',!d.qbo);
  $('tel').href='/telecharger/'+d.download_id; q.href='/telecharger/'+d.download_id+'?format=qbo'; q.title=T('qbo_aide');
  res.classList.remove('masque'); if(defiler) res.scrollIntoView({behavior:'smooth',block:'nearest'});
}
function majVerif(e){const p=$('verifEtat');p.className='pill'+(e==='encours'?' encours':'');
  p.innerHTML=e==='encours'?`<span class="spin"></span> ${T('v_encours')}`:(e==='erreur'?T('v_erreur'):'✓ '+T('v_ajour'));}
async function corriger(i,id){
  if(!dernier||!id) return;
  majVerif('encours'); const tel=$('tel'); tel.style.pointerEvents='none'; tel.style.opacity=.5;
  try{
    const r=await fetch('/api/comptes',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({jeton:dernier.download_id,comptes:{[i]:id}})});
    if(!r.ok) throw new Error(r.status);
    const d=await r.json(); dernier.factures=d.factures; dernier.totaux_fichier=d.totaux_fichier;
    if(d.rapprochement) dernier.rapprochement=d.rapprochement;
    afficher(dernier,false);
  }catch(e){majVerif('erreur');}
  tel.style.pointerEvents=''; tel.style.opacity='';
}

document.body.classList.toggle('comptable',MODE==='comptable');
choisirMode(MODE); rendreOptions(); setLang(LANG);
</script>
</body>
</html>"""


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "7860"))
    uvicorn.run(app, host="0.0.0.0", port=port)
