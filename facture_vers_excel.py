#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
facture_vers_excel.py  —  V2.1
=================================================================
Extrait les données de factures et de reçus et les exporte dans un
tableur Excel propre (2 feuilles : « Resume » et « Details »), et en
option dans un CSV d'import de factures fournisseurs QuickBooks Online.

Deux entrées acceptées, sans conversion préalable :
    - PDF « texte » (factures générées par ordinateur)  -> extraction texte
    - PHOTO / IMAGE d'une facture ou d'un reçu           -> extraction vision
      Formats : .jpg .jpeg .png .webp .gif .bmp .tiff .heic .heif
      (les photos HEIC d'iPhone sont converties automatiquement)
    - PDF scanné (sans texte)                            -> pages rendues en
      images puis extraction vision (si PyMuPDF est installé)

L'extraction fiable passe par un LLM (Claude ou OpenAI) qui renvoie un
JSON structuré et validé. Repli hors ligne « tables » pour les PDF texte
lorsqu'aucune clé API n'est configurée.

Usage rapide :
    export ANTHROPIC_API_KEY="sk-ant-..."          # ou OPENAI_API_KEY
    python facture_vers_excel.py facture.pdf recu.jpg photo.heic -o sortie.xlsx

Dépendances :
    pip install pdfplumber openpyxl requests pillow
    pip install pillow-heif        # pour les photos HEIC (iPhone)
    pip install pymupdf            # optionnel : PDF scannés -> images
=================================================================
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import sys
from dataclasses import dataclass, field, asdict
from typing import Any, Optional

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    print("ERREUR : module 'pdfplumber' manquant -> pip install pdfplumber",
          file=sys.stderr)
    raise

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    print("ERREUR : module 'openpyxl' manquant -> pip install openpyxl",
          file=sys.stderr)
    raise


# Extensions image reconnues.
_EXT_IMAGE = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff",
              ".heic", ".heif"}
# Formats acceptés tels quels par les API vision (sinon on convertit en JPEG).
_MEDIA_API = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
              ".webp": "image/webp", ".gif": "image/gif"}
_TAILLE_MAX_IMG = 1568  # côté long max recommandé pour la vision (px)


# =================================================================
#  1. MODÈLE DE DONNÉES
# =================================================================

@dataclass
class LigneFacture:
    description: str = ""
    quantite: Optional[float] = None
    prix_unitaire: Optional[float] = None
    montant: Optional[float] = None
    categorie: str = ""      # compte du plan comptable de cet article


@dataclass
class Facture:
    fichier: str = ""
    fournisseur: str = ""
    numero: str = ""
    date: str = ""
    total_ht: Optional[float] = None
    tps: Optional[float] = None      # TPS/GST (Canada)
    tvq: Optional[float] = None      # TVQ/QST (Québec)
    tva: Optional[float] = None      # total des taxes (TPS+TVQ, ou TVA/VAT)
    total_ttc: Optional[float] = None
    devise: str = ""
    lignes: list[LigneFacture] = field(default_factory=list)
    moteur: str = ""
    alertes: list[str] = field(default_factory=list)
    pieces_liees: list[str] = field(default_factory=list)   # autres fichiers du même achat
    numeros_lies: list[str] = field(default_factory=list)   # leurs numéros
    categorie: str = ""      # compte du plan comptable (« code · nom »)

    def valider_coherence(self) -> None:
        def egal(a: float, b: float, tol: float = 0.02) -> bool:
            return abs(a - b) <= max(0.02, abs(b) * tol)

        # Si TPS et TVQ sont donnés, leur somme doit égaler la TVA totale.
        if self.tps is not None and self.tvq is not None:
            somme_taxes = self.tps + self.tvq
            if self.tva is None:
                self.tva = round(somme_taxes, 2)
            elif not egal(somme_taxes, self.tva):
                self.alertes.append(
                    f"TPS+TVQ ({somme_taxes:.2f}) ≠ TVA totale ({self.tva:.2f})")

        if (self.total_ht is not None and self.tva is not None
                and self.total_ttc is not None):
            if not egal(self.total_ht + self.tva, self.total_ttc):
                self.alertes.append(
                    f"HT+TVA ({self.total_ht + self.tva:.2f}) ≠ TTC ({self.total_ttc:.2f})")

        if self.lignes and self.total_ht is not None:
            s = sum(l.montant for l in self.lignes if l.montant is not None)
            if s and not egal(s, self.total_ht, tol=0.05):
                self.alertes.append(
                    f"Somme des lignes ({s:.2f}) ≠ HT ({self.total_ht:.2f})")


# =================================================================
#  2. LECTURE DES SOURCES (PDF texte / images)
# =================================================================

def extraire_texte_pdf(chemin: str) -> str:
    morceaux: list[str] = []
    with pdfplumber.open(chemin) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            morceaux.append(f"--- Page {i} ---\n{page.extract_text() or ''}")
    return "\n\n".join(morceaux).strip()


def _pdf_a_du_texte(chemin: str) -> bool:
    try:
        return len(extraire_texte_pdf(chemin)) >= 25
    except Exception:
        return False


def _redimensionner(img):
    """Réduit une image PIL pour que son côté long ≤ _TAILLE_MAX_IMG."""
    r = _TAILLE_MAX_IMG / max(img.size)
    if r < 1:
        img = img.resize((max(1, int(img.size[0] * r)),
                          max(1, int(img.size[1] * r))))
    return img


def preparer_image(chemin: str) -> tuple[str, str]:
    """
    Renvoie (media_type, donnees_base64) prêtes pour une API vision.
    Convertit HEIC et tout format non standard en JPEG, et redimensionne.
    """
    ext = os.path.splitext(chemin)[1].lower()
    if ext in (".heic", ".heif"):
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except ImportError as e:
            raise ErreurService(
                "Photo HEIC détectée mais 'pillow-heif' n'est pas installé "
                "-> pip install pillow-heif") from e

    try:
        from PIL import Image
    except ImportError as e:
        raise ErreurService("Module 'pillow' manquant -> pip install pillow") from e

    # Format déjà accepté par l'API et image raisonnable : envoi direct.
    if ext in _MEDIA_API:
        with Image.open(chemin) as im:
            besoin_reduction = max(im.size) > _TAILLE_MAX_IMG
        if not besoin_reduction and os.path.getsize(chemin) < 4_500_000:
            with open(chemin, "rb") as f:
                return _MEDIA_API[ext], base64.b64encode(f.read()).decode()

    # Sinon : (re)encodage en JPEG après redimensionnement.
    with Image.open(chemin) as im:
        im = im.convert("RGB")
        im = _redimensionner(im)
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=85)
    return "image/jpeg", base64.b64encode(buf.getvalue()).decode()


def _rendre_pdf_en_images(chemin: str) -> list[tuple[str, str]]:
    """PDF scanné -> liste d'images base64 (via PyMuPDF si dispo)."""
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise ErreurService(
            "PDF scanné (sans texte). Installez 'pymupdf' pour le traiter, "
            "ou envoyez-le en photo -> pip install pymupdf") from e
    from PIL import Image
    images = []
    doc = fitz.open(chemin)
    for page in doc:
        pix = page.get_pixmap(dpi=150)
        im = Image.open(io.BytesIO(pix.tobytes("png")))
        im = _redimensionner(im.convert("RGB"))
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=85)
        images.append(("image/jpeg", base64.b64encode(buf.getvalue()).decode()))
    return images


# =================================================================
#  3. EXTRACTION PAR LLM (texte OU image)
# =================================================================

_SCHEMA_FACTURE = {
    "type": "object",
    "properties": {
        "fournisseur": {"type": "string", "description": "Nom du commerce/entreprise émetteur"},
        "numero": {"type": "string", "description": "Numéro de facture, reçu ou transaction"},
        "date": {"type": "string", "description": "Date au format AAAA-MM-JJ si possible"},
        "devise": {"type": "string", "description": "Code ISO 4217 (CAD, USD, EUR...). Un « $ » seul, sans mention US/USD, = CAD."},
        "total_ht": {"type": ["number", "null"], "description": "Sous-total avant taxes"},
        "tps": {"type": ["number", "null"], "description": "Montant TPS/GST (taxe fédérale, 5%)"},
        "tvq": {"type": ["number", "null"], "description": "Montant TVQ/QST (taxe Québec, 9.975%)"},
        "tva": {"type": ["number", "null"], "description": "Total de TOUTES les taxes (TPS+TVQ, ou TVA/VAT)"},
        "total_ttc": {"type": ["number", "null"], "description": "Montant total payé, taxes comprises"},
        "lignes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "description": {"type": "string"},
                    "quantite": {"type": ["number", "null"]},
                    "prix_unitaire": {"type": ["number", "null"]},
                    "montant": {"type": ["number", "null"]},
                },
                "required": ["description"],
            },
        },
    },
    "required": ["fournisseur", "numero", "date", "total_ht", "tps", "tvq",
                 "tva", "total_ttc", "lignes"],
}

_INSTRUCTION = (
    "Tu es un extracteur de données de factures et de reçus très rigoureux. "
    "On te donne le texte OU la photo d'une facture/d'un reçu (français ou anglais). "
    "Tu renvoies UNIQUEMENT les données structurées demandées, sans commentaire. "
    "Règles : "
    "1) Montants = nombres (point décimal, sans symbole ni séparateur de milliers). "
    "2) Valeur absente = null ; ne devine jamais un montant. "
    "3) 'tps' = TPS/GST si elle figure, 'tvq' = TVQ/QST si elle figure, sinon null. "
    "'tva' = TOTAL de toutes les taxes de la pièce (TPS+TVQ, TVH/HST, TPS+TVP/PST, "
    "TVA/VAT…) ; ne le limite jamais à une seule taxe quand il y en a plusieurs. "
    "4) Date au format AAAA-MM-JJ si possible. "
    "5) Extrais chaque ligne (description, quantité, prix unitaire, montant). "
    "Une ligne '3 @ 4.95' = quantité 3, prix unitaire 4.95, montant 14.85. "
    "6) N'invente aucune ligne ; si pas de détail, renvoie une liste vide. "
    "7) Ne mets jamais un numéro de carte ou de compte comme numéro de facture. "
    "8) Devise = code ISO (CAD, USD, EUR...). Un « $ » sans mention explicite "
    "« USD », « US$ » ou « US » = CAD : la clientèle est canadienne, et des taxes "
    "TPS/TVQ/TVH (5 %, 13 %, 14,975 %...) confirment le CAD."
)

_PROMPT_TXT = ("Voici le texte de la facture. Extrais les données structurées.\n\n"
               "<<<FACTURE\n{texte}\nFACTURE>>>")
_PROMPT_IMG = "Voici la photo d'une facture ou d'un reçu. Extrais les données structurées."


class ErreurLLM(RuntimeError):
    pass


class ErreurService(ErreurLLM):
    """Panne du service de lecture (clé, API, module manquant) : le détail
    est pour les journaux du serveur, pas pour l'utilisateur."""


def extraire_avec_llm(
    texte: Optional[str] = None,
    images: Optional[list[tuple[str, str]]] = None,
    fournisseur: str = "auto",
    modele: Optional[str] = None,
    timeout: int = 120,
    plan: Optional[list[dict]] = None,
) -> dict[str, Any]:
    """
    Envoie du texte OU des images à un LLM et renvoie un dict structuré.
    Avec un `plan` comptable, le LLM choisit aussi le compte de la pièce.
    """
    if not texte and not images:
        raise ErreurLLM("Aucun contenu à extraire.")
    fournisseur = _resoudre_fournisseur(fournisseur)
    schema, consigne = _SCHEMA_FACTURE, ""
    if plan:
        schema = json.loads(json.dumps(_SCHEMA_FACTURE))
        schema["properties"]["compte"] = {
            "type": "string", "enum": [_id_compte(c) for c in plan],
            "description": "Compte du plan comptable le plus adapté à la pièce"}
        schema["required"] = schema["required"] + ["compte"]
        article = schema["properties"]["lignes"]["items"]
        article["properties"]["compte"] = {
            "type": "string", "enum": [_id_compte(c) for c in plan],
            "description": "Compte du plan comptable de cet article"}
        article["required"] = article["required"] + ["compte"]
        consigne = _consigne_plan(plan)
    if fournisseur == "claude":
        return _appeler_claude(texte, images, modele, timeout, schema, consigne)
    if fournisseur == "openai":
        return _appeler_openai(texte, images, modele, timeout, schema, consigne)
    raise ErreurService(f"Fournisseur LLM inconnu : {fournisseur}")


def _resoudre_fournisseur(fournisseur: str) -> str:
    fournisseur = (fournisseur or "auto").lower()
    if fournisseur != "auto":
        return fournisseur
    if os.getenv("ANTHROPIC_API_KEY"):
        return "claude"
    if os.getenv("OPENAI_API_KEY"):
        return "openai"
    raise ErreurService(
        "Aucune clé API trouvée. Définissez ANTHROPIC_API_KEY ou OPENAI_API_KEY. "
        "(Le mode --moteur tables hors ligne ne fonctionne que sur les PDF texte.)")


def _appeler_claude(texte, images, modele, timeout, schema=_SCHEMA_FACTURE, consigne=""):
    import requests
    cle = os.getenv("ANTHROPIC_API_KEY")
    if not cle:
        raise ErreurService("ANTHROPIC_API_KEY non définie.")
    base = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com").rstrip("/")
    modele = modele or os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5")

    contenu: list[dict[str, Any]] = []
    for media_type, b64 in (images or []):
        contenu.append({"type": "image", "source": {
            "type": "base64", "media_type": media_type, "data": b64}})
    contenu.append({"type": "text",
                    "text": (_PROMPT_TXT.format(texte=texte) if texte else _PROMPT_IMG)
                    + consigne})

    corps = {
        "model": modele, "max_tokens": 4096, "system": _INSTRUCTION,
        "tools": [{"name": "enregistrer_facture",
                   "description": "Enregistre les données extraites.",
                   "input_schema": schema}],
        "tool_choice": {"type": "tool", "name": "enregistrer_facture"},
        "messages": [{"role": "user", "content": contenu}],
    }
    rep = requests.post(f"{base}/v1/messages",
                        headers={"x-api-key": cle,
                                 "anthropic-version": "2023-06-01",
                                 "content-type": "application/json"},
                        json=corps, timeout=timeout)
    if rep.status_code >= 400:
        raise ErreurService(f"Erreur API Claude {rep.status_code} : {rep.text[:400]}")
    for bloc in rep.json().get("content", []):
        if bloc.get("type") == "tool_use":
            return bloc["input"]
    raise ErreurService("Réponse Claude sans appel d'outil exploitable.")


def _appeler_openai(texte, images, modele, timeout, schema=_SCHEMA_FACTURE, consigne=""):
    import requests
    cle = os.getenv("OPENAI_API_KEY")
    if not cle:
        raise ErreurService("OPENAI_API_KEY non définie.")
    base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    modele = modele or os.getenv("OPENAI_MODEL", "gpt-4o-mini")

    contenu: list[dict[str, Any]] = [
        {"type": "text",
         "text": (_PROMPT_TXT.format(texte=texte) if texte else _PROMPT_IMG) + consigne}]
    for media_type, b64 in (images or []):
        contenu.append({"type": "image_url",
                        "image_url": {"url": f"data:{media_type};base64,{b64}"}})

    schema = json.loads(json.dumps(schema))
    schema["additionalProperties"] = False
    corps = {
        "model": modele,
        "messages": [{"role": "system", "content": _INSTRUCTION},
                     {"role": "user", "content": contenu}],
        "response_format": {"type": "json_schema",
                            "json_schema": {"name": "facture", "schema": schema}},
    }
    rep = requests.post(f"{base}/chat/completions",
                        headers={"Authorization": f"Bearer {cle}",
                                 "content-type": "application/json"},
                        json=corps, timeout=timeout)
    if rep.status_code >= 400:
        raise ErreurService(f"Erreur API OpenAI {rep.status_code} : {rep.text[:400]}")
    return json.loads(rep.json()["choices"][0]["message"]["content"])


# =================================================================
#  4. EXTRACTION HORS LIGNE (PDF texte uniquement)
# =================================================================

_MOTIFS = {
    "numero": re.compile(
        r"(?:facture|invoice|inv|n[o°º]|numéro|number)\s*[:#]?\s*([A-Z0-9\-\/]{3,})",
        re.IGNORECASE),
    "date": re.compile(
        r"(?:date)\s*[:#]?\s*([0-9]{1,4}[\-\/\.][0-9]{1,2}[\-\/\.][0-9]{1,4}"
        r"|[0-9]{1,2}\s+\w+\s+[0-9]{4})", re.IGNORECASE),
    "total_ht": re.compile(
        r"(?:total\s*(?:ht|hors\s*taxe|before\s*tax)|sous[\-\s]*total|subtotal)"
        r"\s*[:#]?\s*([0-9][0-9\s.,]*)", re.IGNORECASE),
    "tva": re.compile(
        r"(?:tva|t\.v\.a|vat)\s*(?:\([0-9.,%]+\))?\s*[:#]?\s*([0-9][0-9\s.,]*)",
        re.IGNORECASE),
    "tps": re.compile(r"(?:tps|gst)\s*(?:\([0-9.,%]+\))?\s*[:#]?\s*([0-9][0-9\s.,]*)",
                      re.IGNORECASE),
    "tvq": re.compile(r"(?:tvq|qst)\s*(?:\([0-9.,%]+\))?\s*[:#]?\s*([0-9][0-9\s.,]*)",
                      re.IGNORECASE),
    "total_ttc": re.compile(
        r"(?:total\s*(?:ttc|toutes\s*taxes)|montant\s*(?:total|dû|du)|grand\s*total"
        r"|total\s*due|amount\s*due|total)\s*[:#$]?\s*([0-9][0-9\s.,]*)", re.IGNORECASE),
}
_MOTIF_DEVISE = re.compile(r"(CAD|EUR|USD|GBP|\$|€|£)")


def _nombre(brut) -> Optional[float]:
    if brut is None:
        return None
    if isinstance(brut, (int, float)):
        return float(brut)
    s = re.sub(r"[^\d,.\-]", "", str(brut).replace(" ", "")
               .replace(" ", "").replace("\xa0", ""))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") \
            else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def extraire_avec_tables(chemin: str) -> dict[str, Any]:
    texte = extraire_texte_pdf(chemin)
    res: dict[str, Any] = {k: None for k in
                           ("total_ht", "tps", "tvq", "tva", "total_ttc")}
    res.update({"fournisseur": "", "numero": "", "date": "", "devise": "", "lignes": []})

    for l in (x.strip() for x in texte.splitlines()):
        if l and not l.startswith("--- Page") and not re.match(
                r"(facture|invoice)", l, re.IGNORECASE) and len(l) > 2:
            res["fournisseur"] = l[:80]
            break
    for cle, motif in _MOTIFS.items():
        m = motif.search(texte)
        if m:
            v = m.group(1).strip()
            res[cle] = _nombre(v) if cle in ("total_ht", "tps", "tvq", "tva",
                                             "total_ttc") else v
    md = _MOTIF_DEVISE.search(texte)
    if md:
        res["devise"] = md.group(1)

    with pdfplumber.open(chemin) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for rangee in table:
                    cell = [(c or "").strip() for c in rangee]
                    nums = [_nombre(c) for c in cell if _nombre(c) is not None]
                    desc = next((c for c in cell if c and _nombre(c) is None), "")
                    if desc and nums:
                        res["lignes"].append({
                            "description": desc,
                            "quantite": nums[0] if len(nums) >= 3 else None,
                            "prix_unitaire": nums[-2] if len(nums) >= 2 else None,
                            "montant": nums[-1]})
    return res


# =================================================================
#  5. ORCHESTRATION
# =================================================================

_DEVISES = {"$": "CAD", "CA$": "CAD", "C$": "CAD", "CAD$": "CAD", "$CA": "CAD",
            "$CAN": "CAD", "CAN$": "CAD", "US$": "USD", "USD$": "USD", "$US": "USD",
            "€": "EUR", "£": "GBP"}


def normaliser_devise(v) -> str:
    """Ramène une devise à son code ISO : « $ » seul → CAD (marché canadien)."""
    d = str(v or "").strip().upper().replace(" ", "")
    return _DEVISES.get(d, d)


def _vers_facture(chemin: str, brut: dict[str, Any], moteur: str,
                  plan: Optional[list[dict]] = None) -> Facture:
    lignes = [LigneFacture(
        description=str(l.get("description", "")).strip(),
        quantite=_nombre(l.get("quantite")),
        prix_unitaire=_nombre(l.get("prix_unitaire")),
        montant=_nombre(l.get("montant")),
        categorie=libelle_compte(plan, l.get("compte")))
        for l in (brut.get("lignes") or [])]
    f = Facture(
        fichier=os.path.basename(chemin),
        fournisseur=str(brut.get("fournisseur", "")).strip(),
        numero=str(brut.get("numero", "")).strip(),
        date=str(brut.get("date", "")).strip(),
        total_ht=_nombre(brut.get("total_ht")),
        tps=_nombre(brut.get("tps")),
        tvq=_nombre(brut.get("tvq")),
        tva=_nombre(brut.get("tva")),
        total_ttc=_nombre(brut.get("total_ttc")),
        devise=normaliser_devise(brut.get("devise")),
        lignes=lignes, moteur=moteur,
        categorie=libelle_compte(plan, brut.get("compte")))
    # La pièce prend la catégorie qui pèse le plus dans ses articles.
    poids: dict[str, float] = {}
    for l in lignes:
        if l.categorie and l.montant:
            poids[l.categorie] = poids.get(l.categorie, 0) + l.montant
    if poids:
        f.categorie = max(poids, key=poids.get)
    f.valider_coherence()
    return f


def _tokens(s: str) -> set:
    """Tokens significatifs d'un nom de fournisseur (mots de 3+ caractères)."""
    return set(re.findall(r"[a-z0-9àâäéèêëïîôöùûüç]{3,}", (s or "").lower()))


def _remplissage(f: "Facture") -> int:
    """Nombre de champs d'entête renseignés (pour choisir la pièce principale)."""
    champs = (f.fournisseur, f.numero, f.date, f.total_ht, f.tps, f.tvq,
              f.tva, f.total_ttc, f.devise)
    return sum(1 for c in champs if c not in (None, ""))


def _memes_pieces(a: "Facture", b: "Facture") -> bool:
    """Deux pièces d'un même achat : même total, même date, fournisseur proche."""
    if a.total_ttc is None or b.total_ttc is None:
        return False
    if abs(a.total_ttc - b.total_ttc) > 0.01:
        return False
    if a.date and b.date and a.date != b.date:
        return False
    return bool(_tokens(a.fournisseur) & _tokens(b.fournisseur))


def _fusionner(groupe: list["Facture"]) -> "Facture":
    """Fusionne un groupe de pièces en gardant la plus détaillée comme principale."""
    if len(groupe) == 1:
        return groupe[0]
    principale = max(groupe, key=lambda x: (len(x.lignes), _remplissage(x)))
    for autre in groupe:
        if autre is principale:
            continue
        for champ in ("fournisseur", "numero", "date", "total_ht", "tps", "tvq",
                      "tva", "total_ttc", "devise"):
            if not getattr(principale, champ) and getattr(autre, champ):
                setattr(principale, champ, getattr(autre, champ))
        principale.pieces_liees.append(autre.fichier)
        if autre.numero and autre.numero != principale.numero:
            principale.numeros_lies.append(autre.numero)
    principale.alertes = []           # on revalide proprement après fusion
    principale.valider_coherence()
    return principale


def regrouper_factures(factures: list["Facture"], actif: bool = True) -> list["Facture"]:
    """Regroupe les pièces d'un même achat (reçu détaillé + relevé de paiement)."""
    if not actif:
        return factures
    utilises = [False] * len(factures)
    resultat: list["Facture"] = []
    for i, f in enumerate(factures):
        if utilises[i]:
            continue
        groupe = [f]
        for j in range(i + 1, len(factures)):
            if not utilises[j] and _memes_pieces(f, factures[j]):
                groupe.append(factures[j])
                utilises[j] = True
        utilises[i] = True
        resultat.append(_fusionner(groupe))
    return resultat


def _type_source(chemin: str) -> str:
    ext = os.path.splitext(chemin)[1].lower()
    if ext in _EXT_IMAGE:
        return "image"
    if ext == ".pdf":
        return "pdf"
    return "inconnu"


def traiter_facture(chemin, moteur="auto", fournisseur_llm="auto", modele=None,
                    plan=None):
    """Traite un PDF ou une image et renvoie une Facture."""
    type_src = _type_source(chemin)
    cle_dispo = bool(os.getenv("ANTHROPIC_API_KEY") or os.getenv("OPENAI_API_KEY"))

    # --- IMAGES : uniquement via vision ---
    if type_src == "image":
        if moteur == "tables":
            raise ErreurLLM(
                f"{os.path.basename(chemin)} est une image : le mode 'tables' "
                "hors ligne ne peut pas la lire. Utilisez le moteur LLM (clé API).")
        images = [preparer_image(chemin)]
        brut = extraire_avec_llm(images=images, fournisseur=fournisseur_llm,
                                 modele=modele, plan=plan)
        return _vers_facture(chemin, brut, f"vision:{_resoudre_fournisseur(fournisseur_llm)}",
                            plan)

    # --- PDF ---
    if type_src == "pdf":
        pdf_texte = _pdf_a_du_texte(chemin)
        veut_llm = moteur == "llm" or (moteur == "auto" and cle_dispo)

        if not pdf_texte:  # PDF scanné -> vision
            if moteur == "tables":
                f = _vers_facture(chemin, {}, "tables")
                f.alertes.append("PDF scanné (aucun texte) : illisible en mode 'tables'.")
                return f
            images = _rendre_pdf_en_images(chemin)
            brut = extraire_avec_llm(images=images, fournisseur=fournisseur_llm,
                                 modele=modele, plan=plan)
            return _vers_facture(chemin, brut, f"vision:{_resoudre_fournisseur(fournisseur_llm)}",
                            plan)

        if veut_llm:
            try:
                brut = extraire_avec_llm(texte=extraire_texte_pdf(chemin),
                                         fournisseur=fournisseur_llm, modele=modele,
                                         plan=plan)
                return _vers_facture(chemin, brut, f"llm:{_resoudre_fournisseur(fournisseur_llm)}",
                                     plan)
            except ErreurLLM:
                if moteur == "llm":
                    raise
                f = _vers_facture(chemin, extraire_avec_tables(chemin), "tables")
                f.alertes.append("Repli sur 'tables' (LLM indisponible).")
                return f
        return _vers_facture(chemin, extraire_avec_tables(chemin), "tables")

    raise ErreurLLM(f"Type de fichier non pris en charge : {os.path.basename(chemin)}")


def traiter_lot(chemins: list[str], paralleles: int = 4, **options) -> list:
    """
    Traite plusieurs fichiers en parallèle (le temps est surtout passé à
    attendre l'API). Renvoie, dans l'ordre des chemins, une Facture ou
    l'exception levée pour ce fichier : un fichier illisible ne bloque pas le lot.
    """
    from concurrent.futures import ThreadPoolExecutor

    def un(chemin):
        try:
            return traiter_facture(chemin, **options)
        except Exception as e:      # noqa: BLE001 — rapporté par fichier
            return e
    with ThreadPoolExecutor(max_workers=max(1, paralleles)) as ex:
        return list(ex.map(un, chemins))


# =================================================================
#  5 bis. PLAN COMPTABLE (catégorisation des pièces)
# =================================================================
# Un plan = liste de comptes {code, nom, type}. Il vient, par ordre de
# priorité : d'un fichier fourni (Excel/CSV), de l'onglet « Plan comptable »
# d'un classeur existant, ou du plan par défaut ci-dessous (catégories de
# dépenses usuelles d'une petite entreprise canadienne, inspirées du T2125).

_PLAN_DEFAUT = [  # (code, nom FR, nom EN, type)
    ("1500", "Équipement et mobilier", "Equipment and furniture", "actif"),
    ("1550", "Matériel informatique", "Computer equipment", "actif"),
    ("4000", "Ventes et revenus", "Sales and revenue", "revenu"),
    ("5100", "Achats de marchandises", "Purchases of goods for resale", "depense"),
    ("5200", "Publicité et marketing", "Advertising and marketing", "depense"),
    ("5210", "Repas et représentation", "Meals and entertainment", "depense"),
    ("5220", "Assurances", "Insurance", "depense"),
    ("5230", "Intérêts et frais bancaires", "Interest and bank charges", "depense"),
    ("5240", "Fournitures de bureau", "Office supplies", "depense"),
    ("5250", "Logiciels et abonnements", "Software and subscriptions", "depense"),
    ("5260", "Honoraires professionnels", "Professional fees", "depense"),
    ("5270", "Loyer", "Rent", "depense"),
    ("5280", "Entretien et réparations", "Repairs and maintenance", "depense"),
    ("5290", "Frais de déplacement", "Travel", "depense"),
    ("5300", "Télécommunications et services publics", "Telephone and utilities", "depense"),
    ("5310", "Carburant", "Fuel", "depense"),
    ("5320", "Location de véhicule", "Vehicle rental", "depense"),
    ("5330", "Frais de véhicule (entretien, permis)", "Motor vehicle expenses", "depense"),
    ("5340", "Livraison et transport", "Delivery and freight", "depense"),
    ("5350", "Formation", "Training", "depense"),
    ("5360", "Fournitures et matériel (autres)", "Supplies and materials (other)", "depense"),
    ("5900", "Autres dépenses", "Other expenses", "depense"),
]
_TYPES = {"actif": ("Actif", "Asset"), "passif": ("Passif", "Liability"),
          "capitaux": ("Capitaux propres", "Equity"), "revenu": ("Revenus", "Revenue"),
          "depense": ("Dépenses", "Expense")}


def plan_par_defaut(langue: str = "fr") -> list[dict]:
    i = 1 if langue == "en" else 0
    return [{"code": c, "nom": (fr, en)[i], "type": _TYPES[t][i]}
            for c, fr, en, t in _PLAN_DEFAUT]


def _id_compte(c: dict) -> str:
    """Identifiant d'un compte transmis au LLM : son code, sinon son nom."""
    return c["code"] or c["nom"]


def libelle_compte(plan: Optional[list[dict]], ident) -> str:
    """« 5320 · Location de véhicule » à partir de l'identifiant choisi."""
    for c in plan or []:
        if ident is not None and _id_compte(c) == str(ident).strip():
            return f"{c['code']} · {c['nom']}" if c["code"] else c["nom"]
    return ""


def _consigne_plan(plan: list[dict]) -> str:
    comptes = "\n".join(f"- {_id_compte(c)} : {c['nom']}"
                        + (f" ({c['type']})" if c["type"] else "") for c in plan)
    return ("\n\nCatégorise aussi la pièce : dans 'compte', donne l'identifiant "
            "du compte le plus adapté de ce plan comptable, d'après le fournisseur "
            "et les articles achetés (une facture d'achat va en général dans une "
            "dépense ; un bien durable coûteux peut être un actif). Donne aussi "
            "le 'compte' de CHAQUE ligne : sur un reçu mixte (ex. fournitures + "
            "repas), chaque article garde son propre compte, et le 'compte' de la "
            "pièce est celui du plus gros montant. Plan comptable :\n" + comptes)


_TITRES_PLAN = {
    "code": ["code", "n°", "no", "numéro", "numero", "n° compte", "no compte",
             "numéro de compte", "account no", "account no.", "account number",
             "number", "gl", "#"],
    "nom": ["compte", "nom", "libellé", "libelle", "intitulé", "intitule",
            "nom du compte", "description", "account", "account name", "name",
            "title"],
    "type": ["type", "classe", "class", "catégorie", "categorie", "category",
             "nature"],
}


def _plan_depuis_lignes(lignes: list[list]) -> list[dict]:
    """Construit un plan à partir de lignes de tableau (avec ou sans titres)."""
    lignes = [[("" if v is None else str(v).strip()) for v in l] for l in lignes]
    lignes = [l for l in lignes if any(l)]
    if not lignes:
        return []
    titres = [v.lower() for v in lignes[0]]
    col = {}
    for cle, noms in _TITRES_PLAN.items():
        for i, t in enumerate(titres):
            if t in noms and i not in col.values():
                col[cle] = i
                break
    if "nom" in col:
        donnees = lignes[1:]
    else:
        # Sans titres reconnus : « code, nom[, type] » si la 1re colonne est
        # numérique, sinon une simple liste de noms.
        donnees = lignes
        if all(re.fullmatch(r"[0-9][0-9.\-]*", l[0] or "0") for l in donnees) \
                and max(len(l) for l in donnees) > 1:
            col = {"code": 0, "nom": 1, "type": 2}
        else:
            col = {"nom": 0}

    def val(l, cle):
        i = col.get(cle)
        return l[i] if i is not None and i < len(l) else ""

    plan, vus = [], set()
    for l in donnees:
        c = {"code": val(l, "code"), "nom": val(l, "nom"), "type": val(l, "type")}
        if c["nom"] and _id_compte(c) not in vus:
            vus.add(_id_compte(c))
            plan.append(c)
    return plan[:400]


def lire_plan_comptable(chemin: str) -> list[dict]:
    """Lit un plan comptable fourni en Excel (.xlsx) ou CSV."""
    if chemin.lower().endswith(".csv"):
        import csv
        with open(chemin, newline="", encoding="utf-8-sig") as fh:
            echantillon = fh.read(4096)
            fh.seek(0)
            try:
                dialecte = csv.Sniffer().sniff(echantillon, delimiters=",;\t")
            except csv.Error:
                dialecte = csv.excel
            return _plan_depuis_lignes(list(csv.reader(fh, dialecte)))
    from openpyxl import load_workbook
    wb = load_workbook(chemin, read_only=True, data_only=True)
    ws = next((wb[n] for n in wb.sheetnames
               if n.strip().lower() in _noms_connus("Plan comptable")), wb.worksheets[0])
    return _plan_depuis_lignes([list(r) for r in ws.iter_rows(values_only=True)])


def _feuille_plan(wb):
    return next((wb[n] for n in wb.sheetnames
                 if n.strip().lower() in _noms_connus("Plan comptable")), None)


def _plan_de_classeur(wb) -> list[dict]:
    ws = _feuille_plan(wb)
    if ws is None:
        return []
    return _plan_depuis_lignes([list(r)[:3] for r in ws.iter_rows(values_only=True)])


def plan_du_classeur(chemin: str) -> list[dict]:
    """Plan comptable enregistré dans l'onglet dédié d'un classeur Facturo."""
    from openpyxl import load_workbook
    return _plan_de_classeur(load_workbook(chemin, read_only=True, data_only=True))


# =================================================================
#  6. EXPORT EXCEL
# =================================================================

_ENTETE = Font(bold=True, color="FFFFFF", size=11)
_FOND = PatternFill("solid", fgColor="2F5496")
_ALERTE = PatternFill("solid", fgColor="FCE4D6")
_FOND_TOTAL = PatternFill("solid", fgColor="DDEBF7")
_BORD = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
_MON = '#,##0.00'


def _entete(ws, n):
    for c in range(1, n + 1):
        cell = ws.cell(row=1, column=c)
        cell.font, cell.fill = _ENTETE, _FOND
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _BORD
    ws.freeze_panes = "A2"


def _largeurs(ws, maxi=60):
    for col in ws.columns:
        lg = max((len(str(c.value)) for c in col
                  if c.value is not None and not str(c.value).startswith("=")),
                 default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(maxi, lg + 2)


# Les titres français servent de clés internes ; _EN donne leur version
# anglaise. L'Excel est écrit dans la langue demandée (« fr » ou « en »).
_COLS_RESUME = ["Fichier", "Fournisseur", "N°", "Date", "Devise", "Total HT",
                "TPS", "TVQ", "Taxes", "Total TTC", "Nb lignes",
                "Regroupé avec", "Alertes", "Catégorie"]
_COLS_DETAILS = ["Fichier", "Fournisseur", "N°", "Description", "Quantité",
                 "Prix unitaire", "Montant", "Catégorie"]
_EN = {
    "Resume": "Summary", "Details": "Details",
    "Fichier": "File", "Fournisseur": "Vendor", "N°": "Invoice No.",
    "Date": "Date", "Devise": "Currency", "Total HT": "Subtotal",
    "TPS": "GST", "TVQ": "QST", "Taxes": "Tax", "Total TTC": "Total",
    "Nb lignes": "Items", "Regroupé avec": "Grouped with", "Alertes": "Warnings",
    "Description": "Description", "Quantité": "Quantity",
    "Prix unitaire": "Unit price", "Montant": "Amount",
    "Catégorie": "Category", "Plan comptable": "Chart of accounts",
    "Totaux par catégorie": "Totals by category", "Nb factures": "Invoices",
    "Code": "Code", "Compte": "Account", "Type": "Type",
}

# Autres titres acceptés pour une colonne, quand on ajoute à un classeur
# existant (ancienne version de Facturo ou tableau d'un client).
_SYNONYMES = {
    "Resume": ["résumé"],
    "N°": ["n° facture", "no facture", "numéro", "numero", "invoice no",
           "invoice #", "invoice number", "no."],
    "Total HT": ["ht", "sous-total", "net"],
    "Taxes": ["tva", "total taxes", "vat", "total tax"],
    "Total TTC": ["ttc", "montant total"],
    "Nb lignes": ["lignes"],
    "Alertes": ["alerts"],
    "Catégorie": ["categorie", "catégorie comptable", "compte comptable"],
    "Plan comptable": ["plan", "accounts", "coa"],
    "Totaux par catégorie": ["totaux categories", "category totals"],
}


def _titre(cle: str, langue: str) -> str:
    return _EN[cle] if langue == "en" else cle


def _noms_connus(cle: str) -> list[str]:
    """Titres reconnus pour une clé, en minuscules : FR, EN, puis synonymes."""
    return [cle.lower(), _EN[cle].lower()] + _SYNONYMES.get(cle, [])


def _cle_titre(v) -> str:
    return str(v or "").strip().lower()


_TRADUCTIONS_ALERTES = [
    (r"TPS\+TVQ \((.+?)\) ≠ TVA totale \((.+?)\)", r"GST+QST (\1) ≠ total tax (\2)"),
    (r"HT\+TVA \((.+?)\) ≠ TTC \((.+?)\)", r"Subtotal+tax (\1) ≠ total (\2)"),
    (r"Somme des lignes \((.+?)\) ≠ HT \((.+?)\)", r"Sum of lines (\1) ≠ subtotal (\2)"),
    (r"PDF scanné \(aucun texte\) : illisible en mode 'tables'\.",
     "Scanned PDF (no text): unreadable in 'tables' mode."),
    (r"Repli sur 'tables' \(LLM indisponible\)\.",
     "Fell back to 'tables' mode (AI service unavailable)."),
]


def traduire_alerte(texte: str, langue: str) -> str:
    """Traduit une alerte (générée en français) dans la langue demandée."""
    if langue != "en":
        return texte
    for motif, remplacement in _TRADUCTIONS_ALERTES:
        texte = re.sub(motif, remplacement, texte)
    return texte


def _obtenir_feuille(wb, cle: str, cols: list[str], langue: str):
    """
    Récupère (ou crée) la feuille `cle` et renvoie (feuille, n° de colonne de
    chaque clé de `cols`). Une feuille existante est reconnue en français comme
    en anglais, puis passe dans la langue demandée : son nom et les titres
    standard de Facturo sont traduits, les titres personnalisés sont conservés.
    Colonnes retrouvées par leur titre ; titres manquants ajoutés à droite.
    """
    noms = _noms_connus(cle)
    ws = next((wb[n] for n in wb.sheetnames if n.strip().lower() in noms), None)
    if ws is None:
        ws = wb.create_sheet(_titre(cle, langue))
        ws.append([_titre(c, langue) for c in cols])
        _entete(ws, len(cols))
        return ws, {c: i + 1 for i, c in enumerate(cols)}
    if _titre(cle, langue) not in wb.sheetnames:
        ws.title = _titre(cle, langue)

    existants = {_cle_titre(c.value): c.column for c in ws[1] if c.value is not None}
    derniere = max(existants.values(), default=0)
    positions = {}
    for c in cols:
        for nom in _noms_connus(c):
            if nom in existants and existants[nom] not in positions.values():
                positions[c] = existants[nom]
                if nom in (c.lower(), _EN[c].lower()):     # titre standard
                    ws.cell(row=1, column=positions[c], value=_titre(c, langue))
                break
        else:
            derniere += 1
            ws.cell(row=1, column=derniere, value=_titre(c, langue))
            positions[c] = derniere
    _entete(ws, derniere)
    return ws, positions


def _ajouter_ligne(ws, positions: dict, valeurs: dict) -> int:
    """Écrit une ligne sous la dernière ligne remplie ; renvoie son numéro."""
    r = ws.max_row + 1
    for cle, v in valeurs.items():
        cell = ws.cell(row=r, column=positions[cle], value=v)
        # Texte lu sur une pièce (donc non fiable) commençant par « = » :
        # écrit comme texte, jamais comme formule exécutable.
        if cell.data_type == "f":
            cell.data_type = "s"
    return r


_COLS_MONTANTS = ("Total HT", "TPS", "TVQ", "Taxes", "Total TTC")
_MOTIF_TOTAL = re.compile(r"^TOTAL(\s+\S{1,5})?$")


def _est_ligne_total(v) -> bool:
    return isinstance(v, str) and bool(_MOTIF_TOTAL.match(v.strip()))


def _retirer_totaux(ws, positions: dict) -> None:
    """Retire les lignes TOTAL (elles seront recalculées sous les nouvelles)."""
    for r in range(ws.max_row, 1, -1):
        if _est_ligne_total(ws.cell(row=r, column=positions["Fichier"]).value):
            ws.delete_rows(r)


def _calculer_totaux(ws, col: dict) -> list[dict]:
    """
    Cumul des pièces de la feuille (hors lignes TOTAL), par devise.
    `col` : n° de colonne (base 1) de chaque clé présente.
    """
    groupes: dict[str, dict] = {}
    for r in range(2, ws.max_row + 1):
        valeurs = {cle: ws.cell(row=r, column=c).value for cle, c in col.items()}
        if not any(v is not None for v in valeurs.values()) \
                or _est_ligne_total(valeurs.get("Fichier")):
            continue
        d = normaliser_devise(valeurs.get("Devise"))
        g = groupes.setdefault(d, {"devise": d, "nb": 0, "total_ht": 0.0, "tps": 0.0,
                                   "tvq": 0.0, "taxes": 0.0, "total_ttc": 0.0})
        g["nb"] += 1
        for cle, champ in zip(_COLS_MONTANTS, ("total_ht", "tps", "tvq", "taxes", "total_ttc")):
            v = valeurs.get(cle)
            if isinstance(v, (int, float)):
                g[champ] = round(g[champ] + v, 2)
    return list(groupes.values())


def _ajouter_totaux(ws, positions: dict, langue: str) -> None:
    """
    Ajoute en bas du Resume une ligne TOTAL par devise : nombre de pièces et
    sommes HT / TPS / TVQ / taxes / TTC. Valeurs calculées (et non formules)
    pour s'afficher dans tous les lecteurs ; recalculées à chaque ajout.
    """
    mot = "invoice(s)" if langue == "en" else "facture(s)"
    for t in _calculer_totaux(ws, positions):
        r = ws.max_row + 1
        ws.cell(row=r, column=positions["Fichier"], value=f"TOTAL {t['devise']}".strip())
        ws.cell(row=r, column=positions["Fournisseur"], value=f"{t['nb']} {mot}")
        ws.cell(row=r, column=positions["Devise"], value=t["devise"] or None)
        for cle, champ in zip(_COLS_MONTANTS, ("total_ht", "tps", "tvq", "taxes", "total_ttc")):
            ws.cell(row=r, column=positions[cle], value=t[champ]).number_format = _MON
        for cell in ws[r]:
            cell.font, cell.fill, cell.border = Font(bold=True), _FOND_TOTAL, _BORD


def totaux_resume(chemin: str) -> list[dict]:
    """
    Cumul de toutes les pièces du classeur (hors lignes TOTAL), par devise :
    [{devise, nb, total_ht, tps, tvq, taxes, total_ttc}]. Sert à l'affichage web.
    """
    from openpyxl import load_workbook
    wb = load_workbook(chemin)
    ws = next((wb[n] for n in wb.sheetnames
               if n.strip().lower() in _noms_connus("Resume")), None)
    if ws is None:
        return []
    col = {}
    for c in ws[1]:
        for cle in ("Fichier", "Devise") + _COLS_MONTANTS:
            if cle not in col and _cle_titre(c.value) in _noms_connus(cle):
                col[cle] = c.column
    return _calculer_totaux(ws, col) if "Fichier" in col else []


def _ecrire_plan(wb, plan: list[dict], langue: str):
    """(Ré)écrit l'onglet du plan comptable dans la langue demandée ; la 4e
    colonne contient les libellés proposés dans la liste déroulante."""
    ancien = _feuille_plan(wb)
    if ancien is not None:
        wb.remove(ancien)
    cols = ["Code", "Compte", "Type", "Catégorie"]
    ws = wb.create_sheet(_titre("Plan comptable", langue))
    ws.append([_titre(c, langue) for c in cols])
    _entete(ws, len(cols))
    for c in plan:
        ws.append([c["code"] or None, c["nom"], c["type"] or None,
                   libelle_compte(plan, _id_compte(c))])
    _largeurs(ws)
    return ws


def _liste_deroulante(ws, colonne: int, plan_ws, nb: int) -> None:
    """Liste déroulante des catégories (saisie libre toujours permise)."""
    from openpyxl.worksheet.datavalidation import DataValidation
    noms_plan = _noms_connus("Plan comptable")
    ws.data_validations.dataValidation = [
        v for v in ws.data_validations.dataValidation
        if not any(n in (v.formula1 or "").lower() for n in noms_plan)]
    dv = DataValidation(type="list", allow_blank=True, showErrorMessage=False,
                        formula1=f"'{plan_ws.title}'!$D$2:$D${nb + 1}")
    lettre = get_column_letter(colonne)
    dv.add(f"{lettre}2:{lettre}5000")
    ws.add_data_validation(dv)


_SANS_CATEGORIE = {"fr": "(Sans catégorie)", "en": "(Uncategorized)"}


def _repartition_categories(ws, pos_r, wd, pos_d, langue: str) -> list[dict]:
    """
    Montants par (catégorie, devise) de toutes les pièces du classeur.
    Une pièce dont les articles relèvent de plusieurs catégories est répartie
    au prorata du montant des articles (HT, taxes et TTC) ; sinon elle va
    entière dans la catégorie de sa ligne de résumé (modifiable par le
    comptable). Les taxes d'un reçu mixte sont donc réparties approximativement.
    """
    sans = _SANS_CATEGORIE.get(langue, _SANS_CATEGORIE["fr"])
    champs = ("total_ht", "tps", "tvq", "taxes", "total_ttc")
    articles: dict[tuple, list] = {}
    for r in range(2, wd.max_row + 1):
        cle = (wd.cell(row=r, column=pos_d["Fichier"]).value,
               wd.cell(row=r, column=pos_d["N°"]).value)
        m = wd.cell(row=r, column=pos_d["Montant"]).value
        if isinstance(m, (int, float)):
            articles.setdefault(cle, []).append(
                (m, wd.cell(row=r, column=pos_d["Catégorie"]).value))

    groupes: dict[tuple, dict] = {}
    for r in range(2, ws.max_row + 1):
        val = lambda cle: ws.cell(row=r, column=pos_r[cle]).value
        if val("Fichier") is None or _est_ligne_total(val("Fichier")):
            continue
        cat_piece = val("Catégorie") or sans
        devise = normaliser_devise(val("Devise"))
        lignes = articles.get((val("Fichier"), val("N°")), [])
        somme = sum(m for m, _ in lignes)
        cats = {c for _, c in lignes if c}
        if len(cats) >= 2 and somme > 0:
            parts: dict[str, float] = {}
            for m, c in lignes:
                parts[c or cat_piece] = parts.get(c or cat_piece, 0) + m / somme
        else:
            parts = {cat_piece: 1.0}
        for cat, part in parts.items():
            g = groupes.setdefault((cat, devise), dict(
                {"categorie": cat, "devise": devise, "nb": 0}, **{c: 0.0 for c in champs}))
            g["nb"] += 1
            for champ, cle in zip(champs, _COLS_MONTANTS):
                v = val(cle)
                if isinstance(v, (int, float)):
                    g[champ] += v * part
    res = sorted(groupes.values(), key=lambda g: (g["devise"], g["categorie"] == sans,
                                                   g["categorie"]))
    for g in res:
        for c in champs:
            g[c] = round(g[c], 2)
    return res


def _ecrire_totaux_categories(wb, ws, pos_r, wd, pos_d, langue: str) -> None:
    """(Ré)écrit l'onglet « Totaux par catégorie » à partir de tout le classeur."""
    for n in list(wb.sheetnames):
        if n.strip().lower() in _noms_connus("Totaux par catégorie"):
            wb.remove(wb[n])
    cols = ["Catégorie", "Devise", "Nb factures"] + list(_COLS_MONTANTS)
    wt = wb.create_sheet(_titre("Totaux par catégorie", langue), index=2)
    wt.append([_titre(c, langue) for c in cols])
    _entete(wt, len(cols))
    groupes = _repartition_categories(ws, pos_r, wd, pos_d, langue)
    champs = ("total_ht", "tps", "tvq", "taxes", "total_ttc")
    for g in groupes:
        wt.append([g["categorie"], g["devise"] or None, g["nb"]] + [g[c] for c in champs])
    # Une ligne TOTAL par devise (les mêmes montants que le résumé).
    mot = "invoice(s)" if langue == "en" else "facture(s)"
    for t in _calculer_totaux(ws, pos_r):
        wt.append([f"TOTAL {t['devise']}".strip(), t["devise"] or None, t["nb"]]
                  + [t[c] for c in ("total_ht", "tps", "tvq", "taxes", "total_ttc")])
        for cell in wt[wt.max_row]:
            cell.font, cell.fill, cell.border = Font(bold=True), _FOND_TOTAL, _BORD
    for row in wt.iter_rows(min_row=2, min_col=4, max_col=len(cols)):
        for cell in row:
            cell.number_format = _MON
    _largeurs(wt)


def construire_excel(factures: list[Facture], chemin_sortie: str,
                     base_excel: Optional[str] = None, langue: str = "fr",
                     plan: Optional[list[dict]] = None) -> str:
    """
    Écrit les factures dans un classeur Excel, titres dans `langue` (« fr »
    ou « en »). Si `base_excel` pointe vers un classeur existant, les nouvelles
    lignes y sont AJOUTÉES à la suite (feuilles Resume/Summary et Details),
    sans écraser l'existant ; le classeur prend la langue de cet ajout.
    Sinon un nouveau classeur est créé. Le `plan` comptable (sinon celui déjà
    enregistré dans le classeur) est écrit dans un onglet dédié et proposé en
    liste déroulante dans la colonne Catégorie.
    """
    if base_excel and os.path.isfile(base_excel):
        from openpyxl import load_workbook
        wb = load_workbook(base_excel)
    else:
        wb = Workbook()
    avant = list(wb.worksheets)
    ws, pos_r = _obtenir_feuille(wb, "Resume", _COLS_RESUME, langue)
    wd, pos_d = _obtenir_feuille(wb, "Details", _COLS_DETAILS, langue)
    # Supprime la feuille par défaut vide d'un classeur neuf ("Sheet").
    for f in avant:
        if f is not ws and f is not wd and f.max_row <= 1 \
                and f.max_column <= 1 and f.cell(1, 1).value is None:
            wb.remove(f)

    _retirer_totaux(ws, pos_r)
    lies = "n° liés" if langue != "en" else "linked no."
    for f in factures:
        regroupe = " ; ".join(f.pieces_liees)
        if f.numeros_lies:
            regroupe += f"  ({lies} : {', '.join(f.numeros_lies)})"
        alertes = [traduire_alerte(a, langue) for a in f.alertes]
        r = _ajouter_ligne(ws, pos_r, {
            "Fichier": f.fichier, "Fournisseur": f.fournisseur, "N°": f.numero,
            "Date": f.date, "Devise": f.devise, "Total HT": f.total_ht,
            "TPS": f.tps, "TVQ": f.tvq, "Taxes": f.tva, "Total TTC": f.total_ttc,
            "Nb lignes": len(f.lignes), "Regroupé avec": regroupe,
            "Alertes": " | ".join(alertes), "Catégorie": f.categorie or None})
        for cle in ("Total HT", "TPS", "TVQ", "Taxes", "Total TTC"):
            ws.cell(row=r, column=pos_r[cle]).number_format = _MON
        if alertes:
            ws.cell(row=r, column=pos_r["Alertes"]).fill = _ALERTE
    _ajouter_totaux(ws, pos_r, langue)
    _largeurs(ws)

    plan = plan or _plan_de_classeur(wb)
    if plan:
        plan_ws = _ecrire_plan(wb, plan, langue)
        _liste_deroulante(ws, pos_r["Catégorie"], plan_ws, len(plan))
        _liste_deroulante(wd, pos_d["Catégorie"], plan_ws, len(plan))

    for f in factures:
        for l in f.lignes:
            r = _ajouter_ligne(wd, pos_d, {
                "Fichier": f.fichier, "Fournisseur": f.fournisseur,
                "N°": f.numero, "Description": l.description,
                "Quantité": l.quantite, "Prix unitaire": l.prix_unitaire,
                "Montant": l.montant, "Catégorie": l.categorie or None})
            for cle in ("Prix unitaire", "Montant"):
                wd.cell(row=r, column=pos_d[cle]).number_format = _MON
    _largeurs(wd)

    # Après l'écriture du résumé ET des articles (base de la répartition).
    _ecrire_totaux_categories(wb, ws, pos_r, wd, pos_d, langue)

    wb.save(chemin_sortie)
    return chemin_sortie


# =================================================================
#  6 bis. EXPORT QUICKBOOKS ONLINE (Canada) — import de factures fournisseurs
# =================================================================
# Format du CSV « Import data > Bills » de QuickBooks Online Canada.
# Colonnes obligatoires : Bill no., Supplier, Bill Date, Due Date, Account,
# Line Amount, Line Tax Code. Une facture à plusieurs articles = plusieurs
# lignes répétant n°, fournisseur et date. Montants HORS taxes : à l'import,
# choisir « Exclusive » et le format de date JJ/MM/AAAA. Fournisseurs et
# comptes doivent exister dans QuickBooks (d'où l'import du plan comptable
# exporté de QuickBooks : la colonne Account reprend le NOM du compte).

COLS_QBO = ["Bill no.", "Supplier", "Bill Date", "Due Date", "Memo", "Account",
            "Line Description", "Line Amount", "Line Tax Code", "Line Tax Amount",
            "Currency Code"]

# Codes de taxe standard de QuickBooks Online Canada (modifiables par client).
CODES_TAXE_QBO = {
    "GST": "GST",                       # TPS seule (5 %)
    "GST/QST": "GST/QST QC - 9.975",    # TPS + TVQ (14,975 %)
    "HST ON": "HST ON",                 # TVH Ontario (13 %)
    "EXEMPT": "Exempt",                 # aucune taxe facturée
    "OUT OF SCOPE": "Out of Scope",     # achat hors Canada
}
# Taux reconnus (taxes / HT) → clé de CODES_TAXE_QBO.
_TAUX_QBO = [(0.05, "GST"), (0.13, "HST ON"), (0.14975, "GST/QST")]


def code_taxe_qbo(f: Facture) -> str:
    """Code de taxe QuickBooks d'une pièce, déduit de ses taxes (vide si inconnu)."""
    taxes = f.tva if f.tva is not None else (
        (f.tps or 0) + (f.tvq or 0) if (f.tps or f.tvq) else None)
    if f.tvq:
        return CODES_TAXE_QBO["GST/QST"]
    if not taxes:
        if f.devise and f.devise != "CAD":
            return CODES_TAXE_QBO["OUT OF SCOPE"]
        # Taxes inconnues (None) : on ne présume pas d'exonération.
        return CODES_TAXE_QBO["EXEMPT"] if taxes == 0 else ""
    if f.tps and abs(f.tps - taxes) <= 0.02:
        return CODES_TAXE_QBO["GST"]
    if f.total_ht:
        taux = taxes / f.total_ht
        for t, cle in _TAUX_QBO:
            if abs(taux - t) <= 0.004:
                return CODES_TAXE_QBO[cle]
    return ""


def _date_qbo(d: str) -> str:
    """AAAA-MM-JJ → JJ/MM/AAAA ; toute autre forme est laissée telle quelle."""
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", (d or "").strip())
    return f"{int(m[3]):02d}/{int(m[2]):02d}/{m[1]}" if m else (d or "")


def _nom_compte(plan: Optional[list[dict]], libelle: str) -> str:
    """« 5320 · Location de véhicule » → nom du compte tel que dans QuickBooks."""
    for c in plan or []:
        if libelle and libelle == libelle_compte(plan, _id_compte(c)):
            return c["nom"]
    return re.sub(r"^\S+\s·\s", "", libelle or "")


def lignes_qbo(f: Facture, plan: Optional[list[dict]] = None) -> list[dict]:
    """
    Lignes du CSV QuickBooks pour une pièce : une par article quand leur
    somme correspond au HT (chaque article garde son compte), sinon une
    seule ligne au montant HT dans le compte de la pièce.
    """
    taxes = f.tva if f.tva is not None else (
        round((f.tps or 0) + (f.tvq or 0), 2) if (f.tps or f.tvq) else None)
    ht = f.total_ht
    if ht is None and f.total_ttc is not None:
        ht = round(f.total_ttc - (taxes or 0), 2)
    numero = f.numero or os.path.splitext(f.fichier)[0]
    memo = " | ".join(x for x in (f.fichier, *f.alertes) if x)
    commun = {"Bill no.": numero, "Supplier": f.fournisseur,
              "Bill Date": _date_qbo(f.date), "Due Date": _date_qbo(f.date),
              "Memo": memo, "Line Tax Code": code_taxe_qbo(f),
              "Currency Code": f.devise}

    articles = [l for l in f.lignes if l.montant is not None]
    somme = round(sum(l.montant for l in articles), 2)
    if articles and ht is not None and abs(somme - ht) <= max(0.02, abs(ht) * 0.005):
        res, reste = [], taxes
        for i, l in enumerate(articles):
            part = None
            if taxes is not None and somme:
                part = round(taxes * l.montant / somme, 2)
                if i == len(articles) - 1:      # le dernier absorbe l'arrondi
                    part = round(reste, 2)
                reste -= part
            res.append(dict(commun, **{
                "Account": _nom_compte(plan, l.categorie or f.categorie),
                "Line Description": l.description, "Line Amount": l.montant,
                "Line Tax Amount": part}))
        return res
    desc = ", ".join(l.description for l in f.lignes if l.description)[:4000]
    return [dict(commun, **{"Account": _nom_compte(plan, f.categorie),
                            "Line Description": desc, "Line Amount": ht,
                            "Line Tax Amount": taxes})]


_DEBUTS_FORMULE = ("=", "+", "-", "@", "\t", "\r")


def _texte_csv_sur(v: str) -> str:
    """Neutralise un texte qu'un tableur prendrait pour une formule (injection
    CSV) : apostrophe en tête, sauf pour un simple nombre comme « -12.50 »."""
    if v.startswith(_DEBUTS_FORMULE) and not re.fullmatch(r"[-+]?\d+(\.\d+)?", v):
        return "'" + v
    return v


def construire_csv_qbo(factures: list[Facture], chemin_sortie: str,
                       plan: Optional[list[dict]] = None) -> str:
    """Écrit le CSV d'import de factures fournisseurs de QuickBooks Online."""
    import csv

    def texte(v):
        if v is None:
            return ""
        if isinstance(v, float):
            return f"{v:.2f}"
        return _texte_csv_sur(str(v))
    with open(chemin_sortie, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS_QBO)
        w.writeheader()
        for f in factures:
            for ligne in lignes_qbo(f, plan):
                w.writerow({k: texte(v) for k, v in ligne.items()})
    return chemin_sortie


# =================================================================
#  7. CLI
# =================================================================

def _parseur():
    p = argparse.ArgumentParser(
        description="Factures/reçus (PDF ou PHOTO) -> Excel. Extraction LLM (V2.1).")
    p.add_argument("fichiers", nargs="+",
                   help="PDF et/ou images (.jpg .png .heic .webp ...).")
    p.add_argument("-o", "--sortie", default="factures.xlsx")
    p.add_argument("--ajouter-a", default=None, metavar="EXISTANT.xlsx",
                   help="Ajoute les nouvelles lignes à un classeur Excel existant.")
    p.add_argument("--moteur", choices=["auto", "llm", "tables"], default="auto")
    p.add_argument("--fournisseur", choices=["auto", "claude", "openai"], default="auto")
    p.add_argument("--modele", default=None)
    p.add_argument("--pas-de-regroupement", action="store_true",
                   help="Désactive le regroupement des pièces d'un même achat.")
    p.add_argument("--plan", default=None, metavar="PLAN.xlsx|csv",
                   help="Plan comptable pour catégoriser les pièces "
                        "(défaut : celui du classeur --ajouter-a, sinon plan standard).")
    p.add_argument("--paralleles", type=int, default=4,
                   help="Nombre de fichiers lus en même temps (défaut 4).")
    p.add_argument("--langue", choices=["fr", "en"], default="fr",
                   help="Langue des titres de l'Excel (fr ou en).")
    p.add_argument("--qbo", default=None, metavar="FACTURES_QBO.csv",
                   help="Écrit aussi le CSV d'import de factures fournisseurs "
                        "de QuickBooks Online (Canada).")
    p.add_argument("--json", action="store_true")
    return p


def main(argv=None):
    args = _parseur().parse_args(argv)
    factures = []
    plan = lire_plan_comptable(args.plan) if args.plan else []
    if not plan and getattr(args, "ajouter_a", None) and os.path.isfile(args.ajouter_a):
        plan = plan_du_classeur(args.ajouter_a)
    plan = plan or plan_par_defaut(args.langue)
    chemins = []
    for chemin in args.fichiers:
        if not os.path.isfile(chemin):
            print(f"  ✗ Introuvable : {chemin}", file=sys.stderr)
        else:
            chemins.append(chemin)
    print(f"  → {len(chemins)} fichier(s), {args.paralleles} en parallèle ...",
          file=sys.stderr)
    resultats = traiter_lot(chemins, paralleles=args.paralleles, moteur=args.moteur,
                            fournisseur_llm=args.fournisseur, modele=args.modele,
                            plan=plan)
    for chemin, f in zip(chemins, resultats):
        if isinstance(f, Exception):
            print(f"  ✗ {os.path.basename(chemin)} : {f}", file=sys.stderr)
            continue
        factures.append(f)
        etat = "⚠ " + " | ".join(f.alertes) if f.alertes else "ok"
        print(f"     [{f.moteur}] {f.fournisseur or '?'} — {f.numero or '?'} — "
              f"TTC {f.total_ttc} {f.devise} — {len(f.lignes)} ligne(s) — {etat}",
              file=sys.stderr)
        if args.json:
            print(json.dumps(asdict(f), ensure_ascii=False, indent=2))

    if not factures:
        print("Aucune facture traitée.", file=sys.stderr)
        return 1
    avant = len(factures)
    factures = regrouper_factures(factures, actif=not args.pas_de_regroupement)
    if len(factures) < avant:
        print(f"  ⓘ {avant - len(factures)} pièce(s) regroupée(s) (même achat).",
              file=sys.stderr)
    base = getattr(args, "ajouter_a", None)
    chemin = construire_excel(factures, args.sortie, base_excel=base,
                              langue=args.langue, plan=plan)
    suffixe = f" (ajoutées à {os.path.basename(base)})" if base else ""
    print(f"\n✓ {len(factures)} facture(s) -> {chemin}{suffixe}", file=sys.stderr)
    if args.qbo:
        construire_csv_qbo(factures, args.qbo, plan=plan)
        print(f"✓ Import QuickBooks Online -> {args.qbo}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
