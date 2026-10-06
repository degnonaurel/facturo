#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère factures_qbo_exemple.csv : un fichier d'import QuickBooks Online
(factures fournisseurs) à faire vérifier par quelqu'un qui a QuickBooks.
Pièces fictives couvrant les cas à valider : TPS+TVQ (Québec), TVH (Ontario),
TPS seule, reçu mixte sur deux comptes, achat exonéré, achat en USD.

    python exemples/quickbooks/generer_exemple_qbo.py
"""
import os
import sys

ICI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.dirname(ICI)))
import facture_vers_excel as fve  # noqa: E402

PLAN = fve.plan_par_defaut("en")

PIECES = [
    ("papeterie_gatineau.jpg", {
        "fournisseur": "Papeterie Gatineau", "numero": "PG-1042", "date": "2026-09-03",
        "devise": "CAD", "total_ht": 120.00, "tps": 6.00, "tvq": 11.97, "tva": 17.97,
        "total_ttc": 137.97, "compte": "5240",
        "lignes": [{"description": "Printer paper (10 reams)", "montant": 80.00, "compte": "5240"},
                   {"description": "Toner cartridge", "montant": 40.00, "compte": "5240"}]}),
    ("resto_hull.jpg", {
        "fournisseur": "Bistro du Ruisseau", "numero": "BR-778", "date": "2026-09-05",
        "devise": "CAD", "total_ht": 64.00, "tps": 3.20, "tvq": 6.38, "tva": 9.58,
        "total_ttc": 73.58, "compte": "5210", "lignes": []}),
    ("quincaillerie_ottawa.pdf", {
        "fournisseur": "Rideau Hardware", "numero": "RH-2291", "date": "2026-09-08",
        "devise": "CAD", "total_ht": 250.00, "tps": None, "tvq": None, "tva": 32.50,
        "total_ttc": 282.50, "compte": "5360",
        "lignes": [{"description": "Shelving brackets", "montant": 150.00, "compte": "5280"},
                   {"description": "Work lunch for crew", "montant": 100.00, "compte": "5210"}]}),
    ("cours_en_ligne.pdf", {
        "fournisseur": "Formation Pro Alberta", "numero": "FPA-55", "date": "2026-09-10",
        "devise": "CAD", "total_ht": 300.00, "tps": 15.00, "tvq": None, "tva": 15.00,
        "total_ttc": 315.00, "compte": "5350", "lignes": []}),
    ("loyer_septembre.pdf", {
        "fournisseur": "Gestion Immobilière Hull", "numero": "LOY-2026-09", "date": "2026-09-01",
        "devise": "CAD", "total_ht": 900.00, "tps": None, "tvq": None, "tva": 0,
        "total_ttc": 900.00, "compte": "5270", "lignes": []}),
    ("logiciel_us.pdf", {
        "fournisseur": "CloudNote Inc.", "numero": "CN-88412", "date": "2026-09-12",
        "devise": "USD", "total_ht": 29.00, "tps": None, "tvq": None, "tva": None,
        "total_ttc": 29.00, "compte": "5250", "lignes": []}),
]


def factures():
    return [fve._vers_facture(nom, brut, "exemple", PLAN) for nom, brut in PIECES]


if __name__ == "__main__":
    sortie = os.path.join(ICI, "factures_qbo_exemple.csv")
    fve.construire_csv_qbo(factures(), sortie, plan=PLAN)
    print(sortie)
