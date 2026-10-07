# -*- coding: utf-8 -*-
"""
releves_bancaires.py — relevés bancaires (CSV / OFX) et rapprochement
=================================================================
Lit un relevé exporté de la banque et rapproche ses transactions des
factures : chaque facture est associée à un débit de même montant, à une
date plausible, de préférence au nom du même fournisseur.

Formats lus :
    - OFX / QFX (1.x SGML et 2.x XML), offert par toutes les banques ;
    - CSV, avec ou sans ligne de titres (FR/EN), montant signé ou colonnes
      débit / crédit séparées, dates AAAA-MM-JJ, JJ/MM/AAAA ou MM/JJ/AAAA.

Le résultat sert à l'onglet « Rapprochement » de l'Excel : transactions
rapprochées, dépenses sans pièce, factures absentes du relevé.
=================================================================
"""
from __future__ import annotations

import csv
import io
import os
import re
from dataclasses import dataclass
from datetime import date
from typing import Optional


@dataclass
class Transaction:
    date: str = ""            # AAAA-MM-JJ
    description: str = ""
    montant: float = 0.0      # négatif = sortie d'argent (débit)
    devise: str = ""
    reference: str = ""


class ErreurReleve(ValueError):
    pass


# =================================================================
#  1. DATES ET MONTANTS
# =================================================================

_ISO = re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})")             # OFX : 20261003120000
_JMA = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})")   # JJ/MM ou MM/JJ


def _vers_iso(a: int, m: int, j: int) -> str:
    try:
        return date(a, m, j).isoformat()
    except ValueError:
        return ""


def ordre_des_dates(valeurs: list[str]) -> str:
    """« jm » (JJ/MM/AAAA) ou « mj » (MM/JJ/AAAA) d'après les valeurs d'une
    colonne ; sans indice (tous les nombres ≤ 12), « mj », l'usage des
    exports bancaires canadiens."""
    for v in valeurs:
        m = _JMA.match(str(v or "").strip())
        if m and int(m[1]) > 12:
            return "jm"
        if m and int(m[2]) > 12:
            return "mj"
    return "mj"


def lire_date(v, ordre: str = "mj") -> str:
    """Date d'un relevé → AAAA-MM-JJ (vide si illisible)."""
    s = str(v or "").strip()
    for motif in (_ISO, _COMPACT):
        m = motif.match(s)
        if m:
            return _vers_iso(int(m[1]), int(m[2]), int(m[3]))
    m = _JMA.match(s)
    if m:
        a, b = int(m[1]), int(m[2])
        j, mo = (a, b) if ordre == "jm" else (b, a)
        return _vers_iso(int(m[3]), mo, j)
    return ""


_MONETAIRE = re.compile(r"^[-+(]?\s*\$?\s*-?[\d\s .,]*[.,]\d{2}\s*\$?\s*\)?$")


def lire_montant(v) -> Optional[float]:
    """« 1 234,56 », « -12.50 », « (12.50) », « $45.00 » → nombre."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", " ")
    if not s:
        return None
    negatif = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in "-.,":
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") \
            else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return -abs(n) if negatif else n


# =================================================================
#  2. LECTURE OFX / QFX
# =================================================================

def _champ(bloc: str, balise: str) -> str:
    m = re.search(rf"<{balise}>([^<\r\n]*)", bloc, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def lire_ofx(texte: str) -> list[Transaction]:
    devise = _champ(texte, "CURDEF").upper()
    blocs = re.findall(r"<STMTTRN>(.*?)(?=</STMTTRN>|<STMTTRN>|</BANKTRANLIST>|$)",
                       texte, re.IGNORECASE | re.DOTALL)
    res = []
    for b in blocs:
        montant = lire_montant(_champ(b, "TRNAMT"))
        d = lire_date(_champ(b, "DTPOSTED"))
        if montant is None or not d:
            continue
        nom, memo = _champ(b, "NAME"), _champ(b, "MEMO")
        desc = nom if not memo or memo in nom else f"{nom} {memo}".strip()
        res.append(Transaction(date=d, description=desc, montant=montant,
                               devise=devise, reference=_champ(b, "FITID")))
    return res


# =================================================================
#  3. LECTURE CSV (titres FR/EN, ou colonnes devinées)
# =================================================================

_TITRES = {
    "date": ["date", "transaction date", "date de transaction", "date de l'opération",
             "date d'opération", "date de l’opération", "posting date", "posted date",
             "date comptable", "date d'inscription", "date inscrite"],
    "description": ["description", "description 1", "libellé", "libelle", "détails",
                    "details", "transaction", "payee", "bénéficiaire", "memo", "nom",
                    "name", "merchant", "marchand", "description de la transaction"],
    "description2": ["description 2", "memo 2", "détails 2"],
    "montant": ["amount", "montant", "cad$", "cad", "transaction amount",
                "montant de la transaction"],
    "debit": ["debit", "débit", "withdrawals", "withdrawal", "retrait", "retraits",
              "paiement", "paiements", "sortie", "sorties", "money out"],
    "credit": ["credit", "crédit", "deposits", "deposit", "dépôt", "dépôts",
               "entrée", "entrées", "money in"],
}


def _norm_titre(v) -> str:
    t = str(v or "").strip().lower()
    return re.sub(r"\s*\(.*?\)\s*", "", t).strip()


def _trouver_titres(lignes: list[list[str]]) -> tuple[int, dict]:
    """(n° de la ligne de titres, colonne de chaque clé) ou (-1, {})."""
    for i, l in enumerate(lignes[:15]):
        titres = [_norm_titre(c) for c in l]
        col = {}
        for cle, noms in _TITRES.items():
            for j, t in enumerate(titres):
                if t in noms and j not in col.values():
                    col[cle] = j
                    break
        if "date" in col and ("montant" in col or "debit" in col or "credit" in col):
            return i, col
    return -1, {}


def _deviner_colonnes(lignes: list[list[str]]) -> dict:
    """Sans titres (TD, Desjardins...) : colonne de date, de description et
    colonnes monétaires (valeurs à 2 décimales) situées après la description."""
    def cell(l, j):
        return l[j].strip() if j < len(l) else ""
    largeur = max(len(l) for l in lignes)
    datees = [l for l in lignes if any(lire_date(c) or _JMA.match(c.strip())
                                       for c in l)]
    if not datees:
        raise ErreurReleve("Aucune date reconnue dans le relevé.")
    col_date = next((j for j in range(largeur)
                     if sum(1 for l in datees if _JMA.match(cell(l, j)) or
                            _ISO.match(cell(l, j))) >= len(datees) / 2), None)
    if col_date is None:
        raise ErreurReleve("Aucune colonne de dates reconnue dans le relevé.")

    def texte(c):
        return c and not _MONETAIRE.match(c) and not re.fullmatch(r"[\d\s\-/.]+", c)
    longueurs = [sum(len(cell(l, j)) for l in datees if texte(cell(l, j)))
                 for j in range(largeur)]
    col_desc = max(range(largeur), key=lambda j: longueurs[j])
    monetaires = [j for j in range(col_desc + 1, largeur)
                  if any(_MONETAIRE.match(cell(l, j)) and cell(l, j) for l in datees)]
    col = {"date": col_date, "description": col_desc}
    if not monetaires:
        raise ErreurReleve("Aucune colonne de montant reconnue dans le relevé.")
    if len(monetaires) >= 3:                       # débit, crédit, solde
        col.update(debit=monetaires[0], credit=monetaires[1])
    elif len(monetaires) == 2 and all(
            not (cell(l, monetaires[0]) and cell(l, monetaires[1])) for l in datees):
        col.update(debit=monetaires[0], credit=monetaires[1])
    else:                                           # montant signé (+ solde)
        col["montant"] = monetaires[0]
    return col


def lire_csv(texte: str, ordre_date: str = "auto") -> list[Transaction]:
    try:
        dialecte = csv.Sniffer().sniff(texte[:4096], delimiters=",;\t")
    except csv.Error:
        dialecte = csv.excel
    lignes = [l for l in csv.reader(io.StringIO(texte), dialecte)
              if any(c.strip() for c in l)]
    if not lignes:
        raise ErreurReleve("Relevé vide.")
    i, col = _trouver_titres(lignes)
    donnees = lignes[i + 1:] if i >= 0 else lignes
    if i < 0:
        col = _deviner_colonnes(lignes)

    def cell(l, cle):
        j = col.get(cle)
        return l[j].strip() if j is not None and j < len(l) else ""
    ordre = ordre_date if ordre_date in ("jm", "mj") else \
        ordre_des_dates([cell(l, "date") for l in donnees])
    res = []
    for l in donnees:
        d = lire_date(cell(l, "date"), ordre)
        if not d:
            continue
        if "montant" in col:
            montant = lire_montant(cell(l, "montant"))
        else:
            debit, credit = lire_montant(cell(l, "debit")), lire_montant(cell(l, "credit"))
            if debit is None and credit is None:
                continue
            montant = abs(credit or 0) - abs(debit or 0)
        if montant is None:
            continue
        desc = " ".join(x for x in (cell(l, "description"), cell(l, "description2")) if x)
        res.append(Transaction(date=d, description=desc, montant=round(montant, 2)))
    return res


def lire_releve(chemin: str, ordre_date: str = "auto") -> list[Transaction]:
    """Lit un relevé CSV ou OFX/QFX ; lève ErreurReleve s'il est illisible."""
    with open(chemin, "rb") as fh:
        brut = fh.read()
    for codage in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texte = brut.decode(codage)
            break
        except UnicodeDecodeError:
            continue
    ext = os.path.splitext(chemin)[1].lower()
    if ext in (".ofx", ".qfx") or re.search(r"<OFX>", texte[:5000], re.IGNORECASE):
        res = lire_ofx(texte)
    else:
        res = lire_csv(texte, ordre_date)
    if not res:
        raise ErreurReleve("Aucune transaction lisible dans le relevé.")
    return res


# =================================================================
#  4. RAPPROCHEMENT TRANSACTIONS ↔ FACTURES
# =================================================================

_MOTS_VIDES = {"inc", "ltd", "ltée", "ltee", "llc", "corp", "the", "les", "des",
               "and", "et", "cie", "co", "enr", "senc", "inc.", "s.a."}
JOURS_AVANT = 3       # paiement un peu avant la date de facture (préautorisation)
JOURS_APRES = 45      # facture fournisseur réglée jusqu'à 45 jours plus tard
JOURS_SANS_NOM = 10   # sans nom reconnu, on n'accepte qu'un écart court


def _mots(s: str) -> set:
    return {m for m in re.findall(r"[a-z0-9àâäéèêëïîôöùûüç]{3,}", (s or "").lower())
            if m not in _MOTS_VIDES}


def meme_fournisseur(fournisseur: str, description: str) -> bool:
    """Nom de la facture reconnu dans le libellé bancaire (souvent tronqué :
    « PAPETERIE GATIN » pour « Papeterie Gatineau »)."""
    desc = _mots(description)
    for f in _mots(fournisseur):
        for d in desc:
            if len(f) >= 4 and len(d) >= 4 and (d.startswith(f[:5]) or f.startswith(d)):
                return True
            if f == d:
                return True
    return False


def _jours(a: str, b: str) -> Optional[int]:
    try:
        return (date.fromisoformat(a) - date.fromisoformat(b)).days
    except (TypeError, ValueError):
        return None


# Fourchettes de change plausibles (CAD pour 1 unité) pour reconnaître une
# facture en devise étrangère payée depuis un compte en dollars canadiens.
# Larges exprès : on ne connaît pas le taux du jour ni les frais de la banque,
# c'est le nom du fournisseur et la date qui confirment le rapprochement.
TAUX_PLAUSIBLES = {"USD": (1.20, 1.55), "EUR": (1.35, 1.70), "GBP": (1.60, 2.00)}


def _montant_compatible(t: Transaction, p: dict) -> Optional[str]:
    """« exact », « converti » (devise étrangère payée en CAD) ou None."""
    total = p.get("total_ttc")
    if not isinstance(total, (int, float)) or total <= 0:
        return None
    paye = abs(t.montant)
    dev_p, dev_t = p.get("devise") or "", t.devise or ""
    if not dev_p or not dev_t or dev_p == dev_t or (dev_p == "CAD" and not dev_t):
        if abs(paye - total) <= 0.01:
            return "exact"
    if dev_p in TAUX_PLAUSIBLES and dev_t in ("", "CAD"):
        bas, haut = TAUX_PLAUSIBLES[dev_p]
        if bas <= paye / total <= haut:
            return "converti"
    return None


def rapprocher(transactions: list[Transaction], pieces: list[dict]) -> dict:
    """
    Associe chaque pièce {fichier, fournisseur, numero, date, devise,
    total_ttc, categorie} à au plus un débit du relevé (et inversement).
    Critères : même montant (à 1 ¢ près), même devise si connue, écart de
    dates plausible ; le nom du fournisseur dans le libellé départage et
    élargit la fenêtre. Une facture en USD/EUR/GBP payée depuis un compte en
    CAD est reconnue (confiance « converti ») si le montant payé tombe dans
    une fourchette de change plausible, avec le nom ET une date proche. Renvoie {"paires": [(i_transaction, i_piece,
    écart, confiance)], "transactions_seules": [...], "pieces_seules": [...]}.
    """
    candidats = []
    for i, t in enumerate(transactions):
        if t.montant >= 0:
            continue
        for k, p in enumerate(pieces):
            montant = _montant_compatible(t, p)
            if montant is None:
                continue
            nom = meme_fournisseur(p.get("fournisseur", ""), t.description)
            ecart = _jours(t.date, p.get("date", ""))
            if montant == "converti":
                # Montant seulement approché : nom du fournisseur ET date proche.
                if not nom or ecart is None or not (-JOURS_AVANT <= ecart <= JOURS_SANS_NOM):
                    continue
                candidats.append((50 - abs(ecart), i, k, ecart, "converti"))
                continue
            if ecart is None:
                if not nom:
                    continue
            elif not (-JOURS_AVANT <= ecart <= (JOURS_APRES if nom else JOURS_SANS_NOM)):
                continue
            confiance = "elevee" if nom and (ecart is not None and abs(ecart) <= 10) \
                else ("moyenne" if nom or (ecart is not None and abs(ecart) <= 3)
                      else "faible")
            score = (100 if nom else 0) - abs(ecart if ecart is not None else 5)
            candidats.append((score, i, k, ecart, confiance))
    paires, t_pris, p_pris = [], set(), set()
    for score, i, k, ecart, confiance in sorted(candidats, key=lambda c: -c[0]):
        if i in t_pris or k in p_pris:
            continue
        t_pris.add(i)
        p_pris.add(k)
        paires.append((i, k, ecart, confiance))
    return {"paires": sorted(paires),
            "transactions_seules": [i for i in range(len(transactions)) if i not in t_pris],
            "pieces_seules": [k for k in range(len(pieces)) if k not in p_pris]}


# =================================================================
#  5. CATÉGORIE SUGGÉRÉE DES DÉPENSES SANS PIÈCE
# =================================================================
# Règles par mots du libellé bancaire (gratuites, sans appel d'IA). Chacune
# donne le code du plan par défaut et des mots qui reconnaissent le même
# compte dans un plan importé (QuickBooks, Sage...). À valider par le comptable.

REGLES_CATEGORIES = [
    (r"esso|petro|shell|ultramar|couche.?tard|irving|husky|pioneer|chevron|essence|gas\b|fuel",
     "5310", ("carburant", "fuel", "essence", "gas")),
    (r"tim hortons|starbucks|second cup|mcdonald|subway|restaurant|resto|bistro|caf[eé]\b|"
     r"uber ?eats|doordash|skip ?the ?dishes|a&w|pizza|sushi|boulangerie|bakery",
     "5210", ("repas", "meal", "restaurant", "représentation", "entertainment")),
    (r"frais (mensuels|bancaires|de service)|service charge|monthly fee|account fee|"
     r"bank fee|interest|int[ée]r[eê]ts|nsf|frais d'utilisation",
     "5230", ("frais bancaires", "bank charge", "intérêts", "interest")),
    (r"bell\b|rogers|telus|vid[ée]otron|fido|koodo|virgin|freedom mobile|hydro|"
     r"enbridge|[ée]nergir|gazifere|internet",
     "5300", ("télécom", "telephone", "téléphone", "utilities", "services publics", "internet")),
    (r"google|microsoft|adobe|amazon web|aws|zoom|github|dropbox|shopify|slack|"
     r"apple\.com|intuit|quickbooks|canva|notion|wix|godaddy|squarespace",
     "5250", ("logiciel", "software", "abonnement", "subscription")),
    (r"staples|bureau en gros|best buy|office depot|papeterie|buroplus",
     "5240", ("fournitures de bureau", "office")),
    (r"air canada|westjet|porter|via rail|h[oô]tel|airbnb|marriott|hilton|uber(?! ?eats)|"
     r"lyft|taxi|parking|stationnement|oc transpo|stm\b|presto|opus",
     "5290", ("déplacement", "travel", "voyage")),
    (r"postes canada|canada post|purolator|fedex|ups\b|dhl|intelcom",
     "5340", ("livraison", "delivery", "transport", "freight")),
    (r"facebook|meta ?ads|google ads|linkedin|instagram|publicit",
     "5200", ("publicité", "advertising", "marketing")),
    (r"assurance|insurance|intact|aviva|la capitale|desjardins assurances",
     "5220", ("assurance", "insurance")),
    (r"loyer|\brent\b|bail",
     "5270", ("loyer", "rent")),
]


def categorie_suggeree(description: str, plan: list[dict]) -> str:
    """« code · compte » du plan pour une dépense sans pièce, ou vide."""
    desc = (description or "").lower()
    for motif, code, mots in REGLES_CATEGORIES:
        if not re.search(motif, desc):
            continue
        compte = next((c for c in plan if c.get("code") == code), None) or next(
            (c for c in plan if any(m in c.get("nom", "").lower() for m in mots)), None)
        if compte:
            return f"{compte['code']} · {compte['nom']}" if compte.get("code") \
                else compte["nom"]
    return ""
