#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génère des factures PDF de démo variées (FR/EN, mises en page différentes)."""
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle)

DOSSIER = "factures_demo"
os.makedirs(DOSSIER, exist_ok=True)
styles = getSampleStyleSheet()
titre = ParagraphStyle("t", parent=styles["Title"], fontSize=20)
petit = ParagraphStyle("p", parent=styles["Normal"], fontSize=9)


def _doc(nom):
    return SimpleDocTemplate(os.path.join(DOSSIER, nom), pagesize=letter,
                             topMargin=2 * cm, bottomMargin=2 * cm)


def _style_table():
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5496")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B0B0")),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
    ])


# --- Facture 1 : FR, avec tableau détecté (comme la démo V1) ---
def facture_fr_tableau():
    d = _doc("facture_fr_boulangerie.pdf")
    e = []
    e.append(Paragraph("Boulangerie Le Croissant Doré", titre))
    e.append(Paragraph("123 rue Wellington, Ottawa, ON — TPS/TVH: 812345678RT", petit))
    e.append(Spacer(1, 0.6 * cm))
    e.append(Paragraph("<b>FACTURE N° :</b> BLG-2026-0142", styles["Normal"]))
    e.append(Paragraph("<b>Date :</b> 2026-09-12", styles["Normal"]))
    e.append(Paragraph("<b>Client :</b> Café du Marché", styles["Normal"]))
    e.append(Spacer(1, 0.5 * cm))
    data = [["Description", "Qté", "Prix unit.", "Montant"],
            ["Baguette tradition", "40", "1,50", "60,00"],
            ["Croissant beurre", "60", "1,25", "75,00"],
            ["Pain au chocolat", "50", "1,40", "70,00"],
            ["Fournée pain de campagne", "20", "3,75", "75,00"]]
    t = Table(data, colWidths=[8 * cm, 2 * cm, 3 * cm, 3 * cm])
    t.setStyle(_style_table())
    e.append(t)
    e.append(Spacer(1, 0.5 * cm))
    tot = [["Total HT", "280,00 $"], ["TVH (13%)", "36,40 $"], ["Total TTC", "316,40 $"]]
    tt = Table(tot, colWidths=[13 * cm, 3 * cm])
    tt.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]))
    e.append(tt)
    d.build(e)


# --- Facture 2 : EN, mise en page « lettre », détail moins tabulaire ---
def facture_en_services():
    d = _doc("invoice_en_consulting.pdf")
    e = []
    e.append(Paragraph("BrightPath Consulting Inc.", titre))
    e.append(Paragraph("400 Bank Street, Ottawa, ON K2P 1Y6", petit))
    e.append(Spacer(1, 0.6 * cm))
    e.append(Paragraph("INVOICE", styles["Heading2"]))
    e.append(Paragraph("Invoice #: BPC-88231", styles["Normal"]))
    e.append(Paragraph("Date: September 8, 2026", styles["Normal"]))
    e.append(Paragraph("Bill to: Nova Retail Group", styles["Normal"]))
    e.append(Spacer(1, 0.5 * cm))
    data = [["Item", "Hours", "Rate", "Amount"],
            ["Website audit", "12", "120.00", "1,440.00"],
            ["SEO optimization", "8", "120.00", "960.00"],
            ["Bilingual content (FR/EN)", "10", "95.00", "950.00"]]
    t = Table(data, colWidths=[8 * cm, 2 * cm, 3 * cm, 3 * cm])
    t.setStyle(_style_table())
    e.append(t)
    e.append(Spacer(1, 0.5 * cm))
    e.append(Paragraph("Subtotal: $3,350.00", styles["Normal"]))
    e.append(Paragraph("HST (13%): $435.50", styles["Normal"]))
    e.append(Paragraph("<b>Total due: $3,785.50 CAD</b>", styles["Normal"]))
    d.build(e)


# --- Facture 3 : FR, format libre sans vrai tableau (piège pour la V1) ---
def facture_fr_libre():
    d = _doc("facture_fr_garage.pdf")
    e = []
    e.append(Paragraph("Garage Mécanique Rivière", titre))
    e.append(Paragraph("77 chemin Montréal, Ottawa, ON", petit))
    e.append(Spacer(1, 0.6 * cm))
    e.append(Paragraph("Facture no GMR/2026/305 — Date: 15/09/2026", styles["Normal"]))
    e.append(Spacer(1, 0.4 * cm))
    # Détail écrit en texte libre, PAS en tableau — casse l'extraction V1.
    corps = (
        "Prestations réalisées sur véhicule Honda Civic 2019 :<br/>"
        "- Vidange huile moteur et filtre — 1 x 89,00 = 89,00<br/>"
        "- Remplacement plaquettes de frein avant — 1 x 210,00 = 210,00<br/>"
        "- Équilibrage 4 roues — 4 x 20,00 = 80,00<br/>"
        "- Main d'œuvre (2h) — 2 x 95,00 = 190,00<br/>"
    )
    e.append(Paragraph(corps, styles["Normal"]))
    e.append(Spacer(1, 0.5 * cm))
    e.append(Paragraph("Sous-total: 569,00 $", styles["Normal"]))
    e.append(Paragraph("TPS (5%): 28,45 $", styles["Normal"]))
    e.append(Paragraph("TVP (8%): 45,52 $", styles["Normal"]))
    e.append(Paragraph("<b>Montant total: 642,97 $</b>", styles["Normal"]))
    d.build(e)


# --- Facture 4 : EN, montants en format US (virgule millier, point décimal) ---
def facture_en_saas():
    d = _doc("invoice_en_saas.pdf")
    e = []
    e.append(Paragraph("CloudSync Software", titre))
    e.append(Spacer(1, 0.6 * cm))
    e.append(Paragraph("Invoice Number: CS-2026-99001", styles["Normal"]))
    e.append(Paragraph("Date: 2026-09-01", styles["Normal"]))
    e.append(Spacer(1, 0.5 * cm))
    data = [["Description", "Qty", "Unit Price", "Amount"],
            ["Pro plan (annual)", "1", "1,200.00", "1,200.00"],
            ["Extra seats", "5", "144.00", "720.00"],
            ["Priority support", "1", "300.00", "300.00"]]
    t = Table(data, colWidths=[8 * cm, 2 * cm, 3 * cm, 3 * cm])
    t.setStyle(_style_table())
    e.append(t)
    e.append(Spacer(1, 0.5 * cm))
    e.append(Paragraph("Subtotal: $2,220.00", styles["Normal"]))
    e.append(Paragraph("Tax (13%): $288.60", styles["Normal"]))
    e.append(Paragraph("<b>Grand Total: $2,508.60</b>", styles["Normal"]))
    d.build(e)


if __name__ == "__main__":
    facture_fr_tableau()
    facture_en_services()
    facture_fr_libre()
    facture_en_saas()
    print("Factures de démo générées dans", DOSSIER)
    for f in sorted(os.listdir(DOSSIER)):
        print("  -", f)
