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

import os
import tempfile
import threading
import uuid
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from starlette.concurrency import run_in_threadpool

import facture_vers_excel as fve

app = FastAPI(title="Facturo")

_TELECHARGEMENTS: dict[str, str] = {}
_DOSSIER = tempfile.mkdtemp(prefix="facturo_")


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
    # Derrière le proxy de l'hébergeur, l'IP réelle est dans X-Forwarded-For.
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


@app.post("/api/extraire")
async def extraire(request: Request,
                   fichiers: list[UploadFile] = File(...),
                   base: Optional[UploadFile] = File(None),
                   plan: Optional[UploadFile] = File(None),
                   langue: str = Form("fr")):
    langue = "en" if langue == "en" else "fr"
    if not _reserver_quota(_visiteur(request), len(fichiers)):
        return JSONResponse({"quota_atteint": True}, status_code=429)
    factures, erreurs = [], []
    with tempfile.TemporaryDirectory() as tmp:
        # Excel existant optionnel : les nouvelles lignes y seront ajoutées.
        base_path = None
        if base is not None and base.filename:
            base_path = os.path.join(tmp, "base.xlsx")
            with open(base_path, "wb") as f:
                f.write(await base.read())

        # Plan comptable : fichier fourni > onglet du classeur de base > défaut.
        plan_comptable = []
        if plan is not None and plan.filename:
            ext = ".csv" if plan.filename.lower().endswith(".csv") else ".xlsx"
            plan_path = os.path.join(tmp, "plan" + ext)
            with open(plan_path, "wb") as f:
                f.write(await plan.read())
            try:
                plan_comptable = fve.lire_plan_comptable(plan_path)
            except Exception:
                pass
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

        chemins = []
        for up in fichiers:
            ext = os.path.splitext(up.filename or "")[1] or ".bin"
            dest = os.path.join(tmp, f"{uuid.uuid4().hex}{ext}")
            with open(dest, "wb") as f:
                f.write(await up.read())
            chemins.append(dest)

        # Lecture en parallèle, hors de la boucle d'événements du serveur.
        resultats = await run_in_threadpool(
            fve.traiter_lot, chemins, paralleles=PARALLELES, plan=plan_comptable)
        for up, res in zip(fichiers, resultats):
            if isinstance(res, fve.ErreurLLM):
                erreurs.append({"fichier": up.filename, "message": str(res)})
            elif isinstance(res, Exception):
                erreurs.append({"fichier": up.filename, "message": f"Erreur : {res}"})
            else:
                res.fichier = up.filename or res.fichier
                factures.append(res)

        if not factures:
            return JSONResponse({"factures": [], "erreurs": erreurs,
                                 "download_id": None, "ajoute": bool(base_path)})

        factures = fve.regrouper_factures(factures)
        jeton = uuid.uuid4().hex
        chemin = os.path.join(_DOSSIER, f"facturo_{jeton}.xlsx")
        try:
            fve.construire_excel(factures, chemin, base_excel=base_path, langue=langue,
                                 plan=plan_comptable)
        except Exception as e:
            # Excel de base illisible : on repart sur un fichier neuf.
            erreurs.append({"fichier": base.filename if base else "Excel existant",
                            "message": (f"Existing Excel unreadable, new file created ({e})"
                                        if langue == "en" else
                                        f"Excel existant illisible, nouveau fichier créé ({e})")})
            fve.construire_excel(factures, chemin, langue=langue, plan=plan_comptable)
        _TELECHARGEMENTS[jeton] = chemin
        # CSV QuickBooks Online : seulement les pièces de cet envoi (pas de doublon
        # à l'import quand on ajoute à un Excel existant).
        try:
            fve.construire_csv_qbo(factures, chemin[:-5] + "_qbo.csv", plan=plan_comptable)
        except Exception as e:
            erreurs.append({"fichier": "QuickBooks", "message": str(e)})
        try:
            totaux = fve.totaux_resume(chemin)
        except Exception:
            totaux = []

    resume = [{
        "fichier": f.fichier, "fournisseur": f.fournisseur, "numero": f.numero,
        "date": f.date, "devise": f.devise, "total_ht": f.total_ht,
        "tps": f.tps, "tvq": f.tvq, "tva": f.tva, "total_ttc": f.total_ttc,
        "nb_lignes": len(f.lignes), "pieces_liees": f.pieces_liees,
        "categorie": f.categorie,
        "alertes": [fve.traduire_alerte(a, langue) for a in f.alertes],
    } for f in factures]
    return {"factures": resume, "erreurs": erreurs, "download_id": jeton,
            "ajoute": bool(base_path), "totaux_fichier": totaux,
            "restant": _restant(_visiteur(request))}


@app.get("/telecharger/{jeton}")
def telecharger(jeton: str, format: str = "xlsx"):
    chemin = _TELECHARGEMENTS.get(jeton)
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
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --paper:#F4F1EA; --card:#FFFFFF; --ink:#1B2024; --muted:#6C7176;
    --bord:#E7E1D5; --pri:#0E6E5E; --pri-fort:#0A5849; --acc:#EE6B3F;
    --acc-doux:#FCE9DF; --ok:#0E6E5E; --alerte:#B25B1E; --ombre:0 10px 30px rgba(20,30,28,.08);
    --r:16px;
  }
  :root[data-theme="dark"], html[data-auto-dark]{}
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      --paper:#121618; --card:#1A2023; --ink:#ECE9E1; --muted:#98A0A2;
      --bord:#2A3236; --pri:#2FBBA0; --pri-fort:#26A38B; --acc:#FF7E56;
      --acc-doux:#33241d; --bord2:#2A3236; --ombre:0 10px 30px rgba(0,0,0,.35);
    }
  }
  *{box-sizing:border-box}
  html,body{margin:0}
  body{font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
       background:
         radial-gradient(1100px 500px at 88% -8%, color-mix(in srgb, var(--pri) 12%, transparent), transparent 60%),
         radial-gradient(900px 500px at 5% 8%, color-mix(in srgb, var(--acc) 10%, transparent), transparent 55%),
         var(--paper);
       color:var(--ink);line-height:1.55;min-height:100vh}
  a{color:inherit}
  .barre{max-width:960px;margin:0 auto;padding:20px 18px;display:flex;align-items:center;gap:12px}
  .logo{display:flex;align-items:center;gap:11px;font-family:"Space Grotesk",sans-serif}
  .logo .nom{font-weight:700;font-size:1.28rem;letter-spacing:-.02em}
  .logo .nom b{color:var(--pri)}
  .droite{margin-left:auto;display:flex;align-items:center;gap:14px}
  .statut{display:inline-flex;align-items:center;gap:7px;font-size:.82rem;color:var(--muted)}
  .pastille{width:9px;height:9px;border-radius:50%;background:#c9ccce}
  .pastille.on{background:var(--ok);box-shadow:0 0 0 4px color-mix(in srgb,var(--ok) 20%,transparent)}
  .pastille.off{background:var(--alerte);box-shadow:0 0 0 4px color-mix(in srgb,var(--alerte) 20%,transparent)}
  .lang{display:flex;border:1px solid var(--bord);border-radius:999px;overflow:hidden;font-size:.8rem;font-weight:600}
  .lang button{border:none;background:transparent;color:var(--muted);padding:6px 12px;cursor:pointer;font:inherit}
  .lang button.actif{background:var(--pri);color:#fff}
  .wrap{max-width:960px;margin:0 auto;padding:10px 18px 70px}
  .hero{padding:20px 4px 26px}
  .hero h1{font-family:"Space Grotesk",sans-serif;font-weight:700;font-size:2.35rem;line-height:1.08;
           letter-spacing:-.03em;margin:.2em 0 .35em}
  .hero h1 .surligne{background:linear-gradient(transparent 62%, color-mix(in srgb,var(--acc) 45%,transparent) 0);padding:0 .06em}
  .hero p{font-size:1.08rem;color:var(--muted);max-width:640px;margin:0}
  .carte{background:var(--card);border:1px solid var(--bord);border-radius:var(--r);box-shadow:var(--ombre)}
  .p24{padding:24px}
  #zone{border:2px dashed color-mix(in srgb,var(--pri) 35%,var(--bord));border-radius:14px;padding:40px 20px;
        text-align:center;cursor:pointer;transition:.15s;background:color-mix(in srgb,var(--pri) 4%,transparent)}
  #zone:hover,#zone.survol{border-color:var(--acc);background:var(--acc-doux)}
  #zone .cercle{width:60px;height:60px;border-radius:50%;background:var(--card);border:1px solid var(--bord);
                display:flex;align-items:center;justify-content:center;margin:0 auto 12px;box-shadow:var(--ombre)}
  #zone .cercle svg{width:28px;height:28px}
  #zone .gros{font-weight:600;font-size:1.05rem}
  #zone .petit{color:var(--muted);font-size:.86rem;margin-top:4px}
  ul.fichiers{list-style:none;padding:0;margin:16px 0 0;display:grid;gap:8px}
  ul.fichiers li{display:flex;align-items:center;gap:11px;padding:10px 13px;border:1px solid var(--bord);
                 border-radius:11px;font-size:.92rem;background:var(--card)}
  ul.fichiers .ico{width:22px;text-align:center}
  ul.fichiers .x{margin-left:auto;color:var(--muted);cursor:pointer;font-weight:700;padding:2px 6px;border-radius:6px}
  ul.fichiers .x:hover{background:var(--acc-doux);color:var(--acc)}
  .base-zone{margin-top:14px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
  .base-btn{display:inline-flex;align-items:center;gap:6px;border:1px dashed var(--bord);border-radius:10px;
            padding:9px 14px;cursor:pointer;font-size:.88rem;font-weight:600;color:var(--pri);background:transparent}
  .base-btn:hover{border-color:var(--pri);background:color-mix(in srgb,var(--pri) 5%,transparent)}
  .base-hint{color:var(--muted);font-size:.8rem}
  .base-nom{width:100%;display:flex;align-items:center;gap:9px;font-size:.88rem;padding:9px 13px;
            border:1px solid var(--bord);border-radius:10px;background:var(--acc-doux)}
  .base-nom .x{margin-left:auto;cursor:pointer;font-weight:700;color:var(--muted);padding:2px 6px;border-radius:6px}
  .base-nom .x:hover{color:var(--acc)}
  .actions{margin-top:18px;display:flex;gap:14px;align-items:center;flex-wrap:wrap}
  .btn{font:inherit;font-weight:600;border:none;border-radius:12px;padding:13px 24px;cursor:pointer;
       font-family:"Space Grotesk",sans-serif}
  .btn.pri{background:var(--pri);color:#fff;box-shadow:0 6px 16px color-mix(in srgb,var(--pri) 35%,transparent)}
  .btn.pri:hover{background:var(--pri-fort)}
  .btn.pri:disabled{opacity:.45;cursor:not-allowed;box-shadow:none}
  .tel{background:var(--acc);color:#fff;text-decoration:none;border-radius:12px;padding:13px 22px;font-weight:600;
       font-family:"Space Grotesk",sans-serif;display:inline-flex;gap:8px;align-items:center;
       box-shadow:0 6px 16px color-mix(in srgb,var(--acc) 38%,transparent)}
  .tel:hover{filter:brightness(.96)}
  .tels{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
  .tel2{color:var(--pri);text-decoration:none;border:1px solid var(--bord);border-radius:12px;padding:12px 18px;
        font-weight:600;font-family:"Space Grotesk",sans-serif;display:inline-flex;gap:8px;align-items:center}
  .tel2:hover{border-color:var(--pri)}
  .spin{width:17px;height:17px;border:3px solid var(--bord);border-top-color:var(--pri);border-radius:50%;
        animation:t .8s linear infinite;display:inline-block;vertical-align:middle}
  @keyframes t{to{transform:rotate(360deg)}}
  .etat{color:var(--muted);font-size:.92rem}
  .quota-note{margin-top:10px;color:var(--muted);font-size:.84rem}
  .quota-note b{color:var(--ink)}
  .quota-note a,.err a{color:var(--pri);font-weight:600}
  .offre{margin-top:22px;padding:22px 24px;display:flex;gap:20px;align-items:center;flex-wrap:wrap;
    justify-content:space-between;border-left:4px solid var(--acc)}
  .offre>div{flex:1 1 260px;min-width:0}
  .offre h3{margin:4px 0 6px;font-family:'Space Grotesk',sans-serif;font-size:1.15rem}
  .offre p{margin:0;color:var(--muted);font-size:.92rem;line-height:1.55}
  .offre-tag{display:inline-block;font-size:.72rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
    color:var(--acc);background:var(--acc-doux);padding:3px 9px;border-radius:999px}
  .offre-btn{text-decoration:none;white-space:nowrap}
  .etapes{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:22px 0 0}
  .etape{display:flex;gap:11px;align-items:flex-start;font-size:.9rem}
  .etape .n{flex:0 0 26px;height:26px;border-radius:8px;background:var(--acc-doux);color:var(--acc);
            font-weight:700;display:flex;align-items:center;justify-content:center;font-family:"Space Grotesk",sans-serif}
  .etape b{display:block}
  .etape span{color:var(--muted)}
  .resultats{margin-top:26px}
  .res-tete{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;margin-bottom:14px}
  .res-tete h2{font-family:"Space Grotesk",sans-serif;font-size:1.2rem;margin:0}
  table{width:100%;border-collapse:collapse;font-size:.88rem;overflow:hidden;border-radius:12px;border:1px solid var(--bord)}
  th,td{padding:10px 12px;text-align:left;border-bottom:1px solid var(--bord)}
  thead th{background:color-mix(in srgb,var(--pri) 10%,var(--card));color:var(--pri-fort);font-weight:600;
           font-family:"Space Grotesk",sans-serif;letter-spacing:.01em}
  td.num{text-align:right;font-variant-numeric:tabular-nums}
  tbody tr:last-child td{border-bottom:none}
  tbody tr:hover{background:color-mix(in srgb,var(--pri) 4%,transparent)}
  .lie{color:var(--muted);font-size:.78rem;display:block;margin-top:2px}
  .al{color:var(--alerte);font-size:.78rem;display:block;margin-top:2px}
  .cumul{margin-top:14px;padding:12px 14px;border-radius:12px;font-size:.92rem;line-height:1.7;
    background:color-mix(in srgb,var(--pri) 8%,var(--card));border:1px solid var(--bord)}
  .err{color:var(--alerte);margin-top:14px;font-size:.9rem}
  .pied{max-width:960px;margin:0 auto;padding:8px 18px 40px;color:var(--muted);font-size:.8rem;text-align:center}
  .masque{display:none}
  @media (max-width:640px){ .hero h1{font-size:1.9rem} .etapes{grid-template-columns:1fr} }
</style>
</head>
<body>
  <div class="barre">
    <div class="logo">
      <svg class="mark" viewBox="0 0 44 44" width="38" height="38" aria-hidden="true">
        <rect x="2" y="2" width="40" height="40" rx="12" fill="var(--pri)"/>
        <path d="M12 11h10.5a2 2 0 0 1 2 2v16l-2.1-1.5-2.1 1.5-2.1-1.5-2.1 1.5-2.1-1.5-1.9 1.35V13a2 2 0 0 1 2-2z" fill="#fff"/>
        <rect x="14.4" y="15.2" width="7.5" height="1.7" rx=".85" fill="var(--pri)"/>
        <rect x="14.4" y="18.6" width="5.5" height="1.7" rx=".85" fill="var(--pri)"/>
        <rect x="24" y="22" width="8.5" height="8.5" rx="2" fill="var(--acc)"/>
        <path d="M28.25 22v8.5M24 26.25h8.5" stroke="#fff" stroke-width="1.1"/>
      </svg>
      <span class="nom">Factur<b>o</b></span>
    </div>
    <div class="droite">
      <span class="statut"><span id="pastille" class="pastille"></span><span id="txtStatut" data-i18n="statut_check">Vérification…</span></span>
      <div class="lang">
        <button id="fr" class="actif" onclick="setLang('fr')">FR</button>
        <button id="en" onclick="setLang('en')">EN</button>
      </div>
    </div>
  </div>

  <div class="wrap">
    <div class="hero">
      <h1><span data-i18n="hero1">Vos factures et reçus,</span><br>
          <span class="surligne" data-i18n="hero2">en Excel — d'une photo.</span></h1>
      <p data-i18n="sous">Déposez une photo ou un PDF. Vous récupérez un tableur propre, prêt pour la compta — en quelques secondes.</p>
    </div>

    <div class="carte p24">
      <div id="zone">
        <div class="cercle">
          <svg viewBox="0 0 24 24" fill="none" stroke="var(--pri)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 16V4M8 8l4-4 4 4"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>
        </div>
        <div class="gros" data-i18n="drop_gros">Glissez vos fichiers ici</div>
        <div class="petit" data-i18n="drop_petit">ou cliquez pour choisir — photos (JPG, PNG, HEIC…) et PDF, plusieurs à la fois</div>
        <input id="input" type="file" multiple accept=".pdf,image/*,.heic,.heif" class="masque">
      </div>
      <ul id="liste" class="fichiers"></ul>

      <div class="base-zone">
        <label class="base-btn">
          <input id="baseInput" type="file" accept=".xlsx" class="masque">
          <span>＋ <span data-i18n="base_btn">Ajouter à un Excel existant</span></span>
        </label>
        <span class="base-hint" data-i18n="base_hint">optionnel — les nouvelles lignes s'ajoutent à la fin de votre fichier</span>
        <div id="baseNom" class="base-nom masque"></div>
      </div>
      <div class="base-zone">
        <label class="base-btn">
          <input id="planInput" type="file" accept=".xlsx,.csv" class="masque">
          <span>＋ <span data-i18n="plan_btn">Importer mon plan comptable</span></span>
        </label>
        <span class="base-hint" data-i18n="plan_hint">optionnel — Excel ou CSV ; sinon un plan standard est utilisé pour catégoriser</span>
        <div id="planNom" class="base-nom masque"></div>
      </div>

      <div class="actions">
        <button id="btn" class="btn pri" disabled data-i18n="btn">Convertir en Excel</button>
        <span id="etat" class="etat"></span>
      </div>
      <div id="quotaNote" class="quota-note masque"></div>
      <div id="erreurs" class="err"></div>

      <div class="etapes">
        <div class="etape"><div class="n">1</div><div><b data-i18n="e1t">Déposez</b><span data-i18n="e1d">Photo ou PDF, en lot.</span></div></div>
        <div class="etape"><div class="n">2</div><div><b data-i18n="e2t">On lit tout</b><span data-i18n="e2d">Fournisseur, dates, taxes, lignes.</span></div></div>
        <div class="etape"><div class="n">3</div><div><b data-i18n="e3t">Téléchargez</b><span data-i18n="e3d">Un Excel propre, deux feuilles.</span></div></div>
      </div>
    </div>

    <div id="resultats" class="resultats masque">
      <div class="carte p24">
        <div class="res-tete">
          <div><h2 id="titreRes" data-i18n="res">Résultat</h2><span id="ajouteNote" class="lie"></span></div>
          <div class="tels">
            <a id="tel" class="tel" href="#"><span>⬇</span><span data-i18n="tel">Télécharger l'Excel</span></a>
            <a id="telQbo" class="tel2" href="#" title=""><span>⬇</span><span data-i18n="tel_qbo">CSV QuickBooks Online</span></a>
          </div>
        </div>
        <div style="overflow-x:auto"><table id="tab"></table></div>
        <div id="cumul" class="cumul masque"></div>
      </div>
    </div>

    <div id="offre" class="carte offre">
      <div>
        <div class="offre-tag" data-i18n="offre_tag">Facturo Pro</div>
        <h3 data-i18n="offre_titre">Une version sur mesure pour votre entreprise</h3>
        <p data-i18n="offre_texte">Plus de volume, votre plan comptable et vos fichiers Excel existants, un format adapté à votre logiciel comptable, un accès privé pour votre équipe. Conçu et accompagné par un consultant bilingue basé à Ottawa.</p>
      </div>
      <a id="offreLien" class="btn pri offre-btn" href="#" data-i18n="offre_btn">Discuter de mon besoin</a>
    </div>
  </div>

  <div class="pied"><span data-i18n="pied">Facturo — vos données ne servent qu'à produire votre tableur.</span></div>

<script>
const I18N = {
  fr:{ statut_check:"Vérification…", statut_on:"Service en ligne", statut_off:"Service indisponible",
    hero1:"Vos factures et reçus,", hero2:"en Excel — d'une photo.",
    sous:"Déposez une photo ou un PDF. Vous récupérez un tableur propre, prêt pour la compta — en quelques secondes.",
    drop_gros:"Glissez vos fichiers ici", drop_petit:"ou cliquez pour choisir — photos (JPG, PNG, HEIC…) et PDF, plusieurs à la fois",
    btn:"Convertir en Excel", traitement:"Lecture en cours…",
    e1t:"Déposez", e1d:"Photo ou PDF, en lot.", e2t:"On lit tout", e2d:"Fournisseur, dates, taxes, lignes.",
    e3t:"Téléchargez", e3d:"Un Excel propre, prêt pour la compta.",
    base_btn:"Ajouter à un Excel existant", base_hint:"optionnel — les nouvelles lignes s'ajoutent à la fin de votre fichier",
    base_choisi:"Excel de base :", ajoute_note:"Nouvelles lignes ajoutées à votre fichier.",
    res:"Résultat", tel:"Télécharger l'Excel", tel_qbo:"CSV QuickBooks Online",
    qbo_aide:"Factures fournisseurs à importer dans QuickBooks Online (Paramètres > Importer des données > Factures) : montants hors taxes, dates JJ/MM/AAAA.", pied:"Facturo — vos données ne servent qu'à produire votre tableur.",
    th_f:"Fournisseur", th_n:"N°", th_d:"Date", th_ht:"HT", th_tps:"TPS", th_tvq:"TVQ", th_tx:"Taxes", th_ttc:"TTC", th_l:"Lignes",
    regroupe:"regroupé avec", pieces:"pièce(s)", aucune:"Aucune donnée n'a pu être extraite.", reseau:"Erreur réseau : ",
    quota:"Limite de la démo atteinte pour aujourd'hui. Revenez demain !",
    gratuit:"Version gratuite", par_jour:"fichiers par jour", restants:"restant(s) aujourd'hui",
    plus:"Besoin de plus ?", voir_pro:"Découvrir Facturo Pro",
    offre_tag:"Facturo Pro", offre_titre:"Une version sur mesure pour votre entreprise",
    offre_texte:"Plus de volume, votre plan comptable et vos fichiers Excel existants, un format adapté à votre logiciel comptable, un accès privé pour votre équipe. Conçu et accompagné par un consultant bilingue basé à Ottawa.",
    offre_btn:"Discuter de mon besoin", mail_sujet:"Facturo Pro — demande d'information",
    mail_corps:"Bonjour,\n\nJ'ai essayé Facturo et j'aimerais en savoir plus sur une version adaptée à mon entreprise.\n\nEntreprise :\nVolume approximatif (factures par mois) :\nLogiciel comptable utilisé :\n\nMerci !",
    plan_btn:"Importer mon plan comptable", plan_hint:"optionnel — Excel ou CSV ; sinon un plan standard est utilisé pour catégoriser",
    plan_choisi:"Plan comptable :", th_cat:"Catégorie",
    cumul:"Total du fichier", dont_taxes:"dont taxes", nb_fact:"facture(s) depuis le début" },
  en:{ statut_check:"Checking…", statut_on:"Service online", statut_off:"Service unavailable",
    hero1:"Your invoices and receipts,", hero2:"in Excel — from a photo.",
    sous:"Drop a photo or a PDF. You get a clean spreadsheet, ready for bookkeeping — in seconds.",
    drop_gros:"Drag your files here", drop_petit:"or click to choose — photos (JPG, PNG, HEIC…) and PDF, several at once",
    btn:"Convert to Excel", traitement:"Reading…",
    e1t:"Drop", e1d:"Photo or PDF, in batches.", e2t:"We read it all", e2d:"Vendor, dates, taxes, line items.",
    e3t:"Download", e3d:"A clean Excel, ready for your books.",
    base_btn:"Add to an existing Excel", base_hint:"optional — new rows are appended to the end of your file",
    base_choisi:"Base Excel:", ajoute_note:"New rows appended to your file.",
    res:"Result", tel:"Download the Excel", tel_qbo:"QuickBooks Online CSV",
    qbo_aide:"Bills to import into QuickBooks Online (Settings > Import data > Bills): amounts exclude tax, dates DD/MM/YYYY.", pied:"Facturo — your data is only used to produce your spreadsheet.",
    th_f:"Vendor", th_n:"No.", th_d:"Date", th_ht:"Net", th_tps:"GST", th_tvq:"QST", th_tx:"Tax", th_ttc:"Total", th_l:"Items",
    regroupe:"merged with", pieces:"item(s)", aucune:"No data could be extracted.", reseau:"Network error: ",
    quota:"Today's demo limit has been reached. Please come back tomorrow!",
    gratuit:"Free version", par_jour:"files per day", restants:"left today",
    plus:"Need more?", voir_pro:"Discover Facturo Pro",
    offre_tag:"Facturo Pro", offre_titre:"A custom version for your business",
    offre_texte:"More volume, your chart of accounts and existing Excel files, output tailored to your accounting software, private access for your team. Built and supported by a bilingual consultant based in Ottawa.",
    offre_btn:"Discuss my needs", mail_sujet:"Facturo Pro — information request",
    mail_corps:"Hello,\n\nI tried Facturo and would like to learn more about a version tailored to my business.\n\nCompany:\nApproximate volume (invoices per month):\nAccounting software used:\n\nThank you!",
    plan_btn:"Import my chart of accounts", plan_hint:"optional — Excel or CSV; otherwise a standard chart is used to categorize",
    plan_choisi:"Chart of accounts:", th_cat:"Category",
    cumul:"File total", dont_taxes:"incl. tax", nb_fact:"invoice(s) since the start" }
};
let LANG = localStorage.getItem('facturo_lang') || (navigator.language||'fr').slice(0,2);
if(LANG!=='en') LANG='fr';
let dernier = null;   // dernier résultat, pour re-rendre au changement de langue

function T(k){return (I18N[LANG]&&I18N[LANG][k])||k;}
function setLang(l){
  LANG=l; localStorage.setItem('facturo_lang',l);
  document.documentElement.lang=l;
  document.getElementById('fr').classList.toggle('actif',l==='fr');
  document.getElementById('en').classList.toggle('actif',l==='en');
  document.querySelectorAll('[data-i18n]').forEach(e=>{const k=e.getAttribute('data-i18n');if(I18N[l][k])e.textContent=I18N[l][k];});
  majStatut(); rendreBase(); rendrePlan(); if(typeof majOffre==='function') majOffre(); if(dernier) afficher(dernier);
}

const input=document.getElementById('input'), zone=document.getElementById('zone'),
      liste=document.getElementById('liste'), btn=document.getElementById('btn'),
      etat=document.getElementById('etat'), res=document.getElementById('resultats'),
      tab=document.getElementById('tab'), tel=document.getElementById('tel'),
      errBox=document.getElementById('erreurs');
let fichiers=[], statutPret=null;

function majStatut(){
  const p=document.getElementById('pastille'), t=document.getElementById('txtStatut');
  if(statutPret===null){t.textContent=T('statut_check');p.className='pastille';}
  else if(statutPret){t.textContent=T('statut_on');p.className='pastille on';}
  else{t.textContent=T('statut_off');p.className='pastille off';}
}
let quota=null, restant=null, contact='';
fetch('/api/statut').then(r=>r.json()).then(s=>{statutPret=!!s.pret;quota=s.quota;restant=s.restant;contact=s.contact||'';majStatut();majOffre();}).catch(()=>{statutPret=false;majStatut();});
function majOffre(){
  const q=document.getElementById('quotaNote');
  if(quota!==null&&quota!==undefined){
    q.classList.remove('masque');
    q.innerHTML=`${T('gratuit')} · ${quota} ${T('par_jour')} · <b>${restant}</b> ${T('restants')} · ${T('plus')} <a href="#offre">${T('voir_pro')}</a>`;
  }
  const l=document.getElementById('offreLien');
  if(contact) l.href=`mailto:${contact}?subject=${encodeURIComponent(T('mail_sujet'))}&body=${encodeURIComponent(T('mail_corps'))}`;
}

zone.onclick=()=>input.click();
['dragover','dragenter'].forEach(e=>zone.addEventListener(e,ev=>{ev.preventDefault();zone.classList.add('survol')}));
['dragleave','drop'].forEach(e=>zone.addEventListener(e,ev=>{ev.preventDefault();zone.classList.remove('survol')}));
zone.addEventListener('drop',ev=>ajouter(ev.dataTransfer.files));
input.addEventListener('change',()=>ajouter(input.files));
function ajouter(fl){for(const f of fl)fichiers.push(f);rendreListe();}
function retirer(i){fichiers.splice(i,1);rendreListe();}

let baseExcel=null;
const baseInput=document.getElementById('baseInput'), baseNom=document.getElementById('baseNom');
baseInput.addEventListener('change',()=>{ if(baseInput.files[0]){baseExcel=baseInput.files[0];rendreBase();} });
function retirerBase(){baseExcel=null;baseInput.value='';rendreBase();}
function rendreBase(){
  if(!baseExcel){baseNom.classList.add('masque');baseNom.innerHTML='';return;}
  baseNom.classList.remove('masque');
  baseNom.innerHTML=`<span>📗</span><span>${T('base_choisi')} ${baseExcel.name}</span><span class="x" title="✕">✕</span>`;
  baseNom.querySelector('.x').onclick=retirerBase;
}
let planFichier=null;
const planInput=document.getElementById('planInput'), planNom=document.getElementById('planNom');
planInput.addEventListener('change',()=>{ if(planInput.files[0]){planFichier=planInput.files[0];rendrePlan();} });
function rendrePlan(){
  if(!planFichier){planNom.classList.add('masque');planNom.innerHTML='';return;}
  planNom.classList.remove('masque');
  planNom.innerHTML=`<span>🗂️</span><span>${T('plan_choisi')} ${planFichier.name}</span><span class="x" title="✕">✕</span>`;
  planNom.querySelector('.x').onclick=()=>{planFichier=null;planInput.value='';rendrePlan();};
}
function rendreListe(){
  liste.innerHTML='';
  fichiers.forEach((f,i)=>{
    const li=document.createElement('li');
    const ic=(f.name||'').toLowerCase().endsWith('.pdf')?'📄':'🖼️';
    li.innerHTML=`<span class="ico">${ic}</span><span>${f.name}</span><span class="x" title="✕">✕</span>`;
    li.querySelector('.x').onclick=()=>retirer(i);
    liste.appendChild(li);
  });
  btn.disabled=fichiers.length===0;
}

btn.onclick=async()=>{
  if(!fichiers.length)return;
  btn.disabled=true;errBox.textContent='';res.classList.add('masque');
  etat.innerHTML='<span class="spin"></span> '+T('traitement');
  const fd=new FormData(); fichiers.forEach(f=>fd.append('fichiers',f,f.name));
  if(baseExcel)fd.append('base',baseExcel,baseExcel.name);
  if(planFichier)fd.append('plan',planFichier,planFichier.name);
  fd.append('langue',LANG);
  try{
    const r=await fetch('/api/extraire',{method:'POST',body:fd});
    if(r.status===429){dernier=null;restant=0;majOffre();
      errBox.innerHTML=`${T('quota')} <a href="#offre">${T('voir_pro')} →</a>`;}
    else{dernier=await r.json(); if(dernier.restant!==undefined){restant=dernier.restant;majOffre();} afficher(dernier);}
  }catch(e){errBox.textContent=T('reseau')+e;}
  etat.textContent=''; btn.disabled=false;
};

function fmt(x){return(x===null||x===undefined||x==='')?'':Number(x).toLocaleString(LANG==='fr'?'fr-CA':'en-CA',{minimumFractionDigits:2,maximumFractionDigits:2});}
function afficher(d){
  errBox.innerHTML=(d.erreurs&&d.erreurs.length)?d.erreurs.map(e=>`⚠ ${e.fichier} : ${e.message}`).join('<br>'):'';
  if(!d.factures||!d.factures.length){res.classList.add('masque');if(d.erreurs&&d.erreurs.length){}else errBox.textContent=T('aucune');return;}
  let h=`<thead><tr><th>${T('th_f')}</th><th>${T('th_n')}</th><th>${T('th_d')}</th>
    <th>${T('th_ht')}</th><th>${T('th_tps')}</th><th>${T('th_tvq')}</th><th>${T('th_tx')}</th>
    <th>${T('th_ttc')}</th><th>${T('th_l')}</th><th>${T('th_cat')}</th></tr></thead><tbody>`;
  d.factures.forEach(f=>{
    let extra='';
    if(f.pieces_liees&&f.pieces_liees.length)extra+=`<span class="lie">↻ ${T('regroupe')} ${f.pieces_liees.join(', ')}</span>`;
    if(f.alertes&&f.alertes.length)extra+=`<span class="al">⚠ ${f.alertes.join(' | ')}</span>`;
    h+=`<tr><td>${f.fournisseur||'—'}${extra}</td><td>${f.numero||''}</td><td>${f.date||''}</td>
        <td class="num">${fmt(f.total_ht)}</td><td class="num">${fmt(f.tps)}</td><td class="num">${fmt(f.tvq)}</td>
        <td class="num">${fmt(f.tva)}</td><td class="num">${fmt(f.total_ttc)} ${f.devise||''}</td>
        <td class="num">${f.nb_lignes}</td><td>${f.categorie||''}</td></tr>`;
  });
  tab.innerHTML=h+'</tbody>';
  document.getElementById('titreRes').textContent=`${T('res')} — ${d.factures.length} ${T('pieces')}`;
  document.getElementById('ajouteNote').textContent=d.ajoute?('↳ '+T('ajoute_note')):'';
  const cumul=document.getElementById('cumul'), tot=d.totaux_fichier||[];
  cumul.innerHTML=tot.map(t=>`Σ ${T('cumul')} : <b>${fmt(t.total_ttc)} ${t.devise}</b> · ${T('dont_taxes')} ${fmt(t.taxes)} ${t.devise} · ${t.nb} ${T('nb_fact')}`).join('<br>');
  cumul.classList.toggle('masque',!tot.length);
  if(d.download_id){tel.href='/telecharger/'+d.download_id;
    const q=document.getElementById('telQbo'); q.href='/telecharger/'+d.download_id+'?format=qbo'; q.title=T('qbo_aide');}
  res.classList.remove('masque'); res.scrollIntoView({behavior:'smooth',block:'nearest'});
}

setLang(LANG);
</script>
</body>
</html>"""


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "7860"))
    uvicorn.run(app, host="0.0.0.0", port=port)
