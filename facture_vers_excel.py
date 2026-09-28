#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
facture_vers_excel.py  —  V2.1
=================================================================
Extrait les données de factures et de reçus et les exporte dans un
tableur Excel propre (2 feuilles : « Resume » et « Details »).

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
            raise ErreurLLM(
                "Photo HEIC détectée mais 'pillow-heif' n'est pas installé "
                "-> pip install pillow-heif") from e

    try:
        from PIL import Image
    except ImportError as e:
        raise ErreurLLM("Module 'pillow' manquant -> pip install pillow") from e

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
        raise ErreurLLM(
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
        "devise": {"type": "string", "description": "Code ou symbole (CAD, USD, EUR, $, €...)"},
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
    "3) Au Québec : 'tps' = TPS/GST, 'tvq' = TVQ/QST, et 'tva' = TPS+TVQ. "
    "Ailleurs : 'tva' = total des taxes (TVA/VAT), 'tps' et 'tvq' à null. "
    "4) Date au format AAAA-MM-JJ si possible. "
    "5) Extrais chaque ligne (description, quantité, prix unitaire, montant). "
    "Une ligne '3 @ 4.95' = quantité 3, prix unitaire 4.95, montant 14.85. "
    "6) N'invente aucune ligne ; si pas de détail, renvoie une liste vide. "
    "7) Ne mets jamais un numéro de carte ou de compte comme numéro de facture."
)

_PROMPT_TXT = ("Voici le texte de la facture. Extrais les données structurées.\n\n"
               "<<<FACTURE\n{texte}\nFACTURE>>>")
_PROMPT_IMG = "Voici la photo d'une facture ou d'un reçu. Extrais les données structurées."


class ErreurLLM(RuntimeError):
    pass


def extraire_avec_llm(
    texte: Optional[str] = None,
    images: Optional[list[tuple[str, str]]] = None,
    fournisseur: str = "auto",
    modele: Optional[str] = None,
    timeout: int = 120,
) -> dict[str, Any]:
    """Envoie du texte OU des images à un LLM et renvoie un dict structuré."""
    if not texte and not images:
        raise ErreurLLM("Aucun contenu à extraire.")
    fournisseur = _resoudre_fournisseur(fournisseur)
    if fournisseur == "claude":
        return _appeler_claude(texte, images, modele, timeout)
    if fournisseur == "openai":
        return _appeler_openai(texte, images, modele, timeout)
    raise ErreurLLM(f"Fournisseur LLM inconnu : {fournisseur}")


def _resoudre_fournisseur(fournisseur: str) -> str:
    fournisseur = (fournisseur or "auto").lower()
    if fournisseur != "auto":
        return fournisseur
    if os.getenv("ANTHROPIC_API_KEY"):
        return "claude"
    if os.getenv("OPENAI_API_KEY"):
        return "openai"
    raise ErreurLLM(
        "Aucune clé API trouvée. Définissez ANTHROPIC_API_KEY ou OPENAI_API_KEY. "
        "(Le mode --moteur tables hors ligne ne fonctionne que sur les PDF texte.)")


def _appeler_claude(texte, images, modele, timeout):
    import requests
    cle = os.getenv("ANTHROPIC_API_KEY")
    if not cle:
        raise ErreurLLM("ANTHROPIC_API_KEY non définie.")
    base = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com").rstrip("/")
    modele = modele or os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-latest")

    contenu: list[dict[str, Any]] = []
    for media_type, b64 in (images or []):
        contenu.append({"type": "image", "source": {
            "type": "base64", "media_type": media_type, "data": b64}})
    contenu.append({"type": "text",
                    "text": _PROMPT_TXT.format(texte=texte) if texte else _PROMPT_IMG})

    corps = {
        "model": modele, "max_tokens": 4096, "system": _INSTRUCTION,
        "tools": [{"name": "enregistrer_facture",
                   "description": "Enregistre les données extraites.",
                   "input_schema": _SCHEMA_FACTURE}],
        "tool_choice": {"type": "tool", "name": "enregistrer_facture"},
        "messages": [{"role": "user", "content": contenu}],
    }
    rep = requests.post(f"{base}/v1/messages",
                        headers={"x-api-key": cle,
                                 "anthropic-version": "2023-06-01",
                                 "content-type": "application/json"},
                        json=corps, timeout=timeout)
    if rep.status_code >= 400:
        raise ErreurLLM(f"Erreur API Claude {rep.status_code} : {rep.text[:400]}")
    for bloc in rep.json().get("content", []):
        if bloc.get("type") == "tool_use":
            return bloc["input"]
    raise ErreurLLM("Réponse Claude sans appel d'outil exploitable.")


def _appeler_openai(texte, images, modele, timeout):
    import requests
    cle = os.getenv("OPENAI_API_KEY")
    if not cle:
        raise ErreurLLM("OPENAI_API_KEY non définie.")
    base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    modele = modele or os.getenv("OPENAI_MODEL", "gpt-4o-mini")

    contenu: list[dict[str, Any]] = [
        {"type": "text", "text": _PROMPT_TXT.format(texte=texte) if texte else _PROMPT_IMG}]
    for media_type, b64 in (images or []):
        contenu.append({"type": "image_url",
                        "image_url": {"url": f"data:{media_type};base64,{b64}"}})

    schema = json.loads(json.dumps(_SCHEMA_FACTURE))
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
        raise ErreurLLM(f"Erreur API OpenAI {rep.status_code} : {rep.text[:400]}")
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

def _vers_facture(chemin: str, brut: dict[str, Any], moteur: str) -> Facture:
    lignes = [LigneFacture(
        description=str(l.get("description", "")).strip(),
        quantite=_nombre(l.get("quantite")),
        prix_unitaire=_nombre(l.get("prix_unitaire")),
        montant=_nombre(l.get("montant"))) for l in (brut.get("lignes") or [])]
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
        devise=str(brut.get("devise", "")).strip(),
        lignes=lignes, moteur=moteur)
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


def traiter_facture(chemin, moteur="auto", fournisseur_llm="auto", modele=None):
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
        brut = extraire_avec_llm(images=images, fournisseur=fournisseur_llm, modele=modele)
        return _vers_facture(chemin, brut, f"vision:{_resoudre_fournisseur(fournisseur_llm)}")

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
            brut = extraire_avec_llm(images=images, fournisseur=fournisseur_llm, modele=modele)
            return _vers_facture(chemin, brut, f"vision:{_resoudre_fournisseur(fournisseur_llm)}")

        if veut_llm:
            try:
                brut = extraire_avec_llm(texte=extraire_texte_pdf(chemin),
                                         fournisseur=fournisseur_llm, modele=modele)
                return _vers_facture(chemin, brut, f"llm:{_resoudre_fournisseur(fournisseur_llm)}")
            except ErreurLLM:
                if moteur == "llm":
                    raise
                f = _vers_facture(chemin, extraire_avec_tables(chemin), "tables")
                f.alertes.append("Repli sur 'tables' (LLM indisponible).")
                return f
        return _vers_facture(chemin, extraire_avec_tables(chemin), "tables")

    raise ErreurLLM(f"Type de fichier non pris en charge : {os.path.basename(chemin)}")


# =================================================================
#  6. EXPORT EXCEL
# =================================================================

_ENTETE = Font(bold=True, color="FFFFFF", size=11)
_FOND = PatternFill("solid", fgColor="2F5496")
_ALERTE = PatternFill("solid", fgColor="FCE4D6")
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
        lg = max((len(str(c.value)) for c in col if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(maxi, lg + 2)


_COLS_RESUME = ["Fichier", "Fournisseur", "N°", "Date", "Devise", "Total HT",
                "TPS", "TVQ", "Taxes", "Total TTC", "Nb lignes", "Moteur",
                "Regroupé avec", "Alertes"]
_COLS_DETAILS = ["Fichier", "Fournisseur", "N°", "Description", "Quantité",
                 "Prix unitaire", "Montant"]


def _obtenir_feuille(wb, nom: str, cols: list[str]):
    """Récupère une feuille existante, ou la crée avec ses entêtes."""
    if nom in wb.sheetnames:
        return wb[nom]
    ws = wb.create_sheet(nom)
    ws.append(cols)
    _entete(ws, len(cols))
    return ws


def construire_excel(factures: list[Facture], chemin_sortie: str,
                     base_excel: Optional[str] = None) -> str:
    """
    Écrit les factures dans un classeur Excel.
    Si `base_excel` pointe vers un classeur existant, les nouvelles lignes y
    sont AJOUTÉES à la suite (feuilles Resume et Details), sans écraser
    l'existant. Sinon un nouveau classeur est créé.
    """
    if base_excel and os.path.isfile(base_excel):
        from openpyxl import load_workbook
        wb = load_workbook(base_excel)
    else:
        wb = Workbook()
    noms_avant = set(wb.sheetnames)
    ws = _obtenir_feuille(wb, "Resume", _COLS_RESUME)
    wd = _obtenir_feuille(wb, "Details", _COLS_DETAILS)
    # Supprime la feuille par défaut vide d'un classeur neuf ("Sheet").
    for nom in list(wb.sheetnames):
        if nom in noms_avant and nom not in ("Resume", "Details"):
            f = wb[nom]
            if f.max_row <= 1 and (f.max_column <= 1 and f.cell(1, 1).value is None):
                del wb[nom]

    debut_r = ws.max_row + 1
    for f in factures:
        regroupe = " ; ".join(f.pieces_liees)
        if f.numeros_lies:
            regroupe += f"  (n° liés : {', '.join(f.numeros_lies)})"
        ws.append([f.fichier, f.fournisseur, f.numero, f.date, f.devise,
                   f.total_ht, f.tps, f.tvq, f.tva, f.total_ttc,
                   len(f.lignes), f.moteur, regroupe, " | ".join(f.alertes)])
    for i, f in enumerate(factures):
        r = debut_r + i
        for c in (6, 7, 8, 9, 10):
            ws.cell(row=r, column=c).number_format = _MON
        if f.alertes:
            ws.cell(row=r, column=14).fill = _ALERTE
    _largeurs(ws)

    debut_d = wd.max_row + 1
    for f in factures:
        for l in f.lignes:
            wd.append([f.fichier, f.fournisseur, f.numero, l.description,
                       l.quantite, l.prix_unitaire, l.montant])
    for r in range(debut_d, wd.max_row + 1):
        for c in (6, 7):
            wd.cell(row=r, column=c).number_format = _MON
    _largeurs(wd)

    wb.save(chemin_sortie)
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
    p.add_argument("--json", action="store_true")
    return p


def main(argv=None):
    args = _parseur().parse_args(argv)
    factures = []
    for chemin in args.fichiers:
        if not os.path.isfile(chemin):
            print(f"  ✗ Introuvable : {chemin}", file=sys.stderr)
            continue
        print(f"  → {os.path.basename(chemin)} [{_type_source(chemin)}] ...",
              file=sys.stderr)
        try:
            f = traiter_facture(chemin, moteur=args.moteur,
                                fournisseur_llm=args.fournisseur, modele=args.modele)
        except ErreurLLM as e:
            print(f"  ✗ {e}", file=sys.stderr)
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
    chemin = construire_excel(factures, args.sortie, base_excel=base)
    suffixe = f" (ajoutées à {os.path.basename(base)})" if base else ""
    print(f"\n✓ {len(factures)} facture(s) -> {chemin}{suffixe}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
