# -*- coding: utf-8 -*-
"""Outils partagés par les tests : rendent la racine importable et
génèrent des entrées de test (PDF texte, image de reçu) à la volée."""
import os
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
if str(RACINE) not in sys.path:
    sys.path.insert(0, str(RACINE))


def pdf_facture_demo(chemin: str) -> str:
    """Crée un petit PDF texte de facture avec un vrai tableau détectable."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle)
    from reportlab.lib.styles import getSampleStyleSheet
    s = getSampleStyleSheet()
    doc = SimpleDocTemplate(chemin, pagesize=letter)
    e = [Paragraph("Atelier Demo Inc.", s["Title"]),
         Paragraph("Facture N° : DEMO-001", s["Normal"]),
         Paragraph("Date : 2026-09-01", s["Normal"]), Spacer(1, 0.4 * cm)]
    data = [["Description", "Qté", "Prix unit.", "Montant"],
            ["Service A", "2", "50.00", "100.00"],
            ["Service B", "1", "80.00", "80.00"]]
    t = Table(data, colWidths=[8 * cm, 2 * cm, 3 * cm, 3 * cm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
    e += [t, Spacer(1, 0.4 * cm),
          Paragraph("Total HT : 180.00", s["Normal"]),
          Paragraph("TVA : 20.00", s["Normal"]),
          Paragraph("Total TTC : 200.00", s["Normal"])]
    doc.build(e)
    return chemin


def image_recu_demo(chemin: str) -> str:
    """Crée une image JPEG synthétique de reçu (pour tester le chemin vision)."""
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (600, 400), "white")
    d = ImageDraw.Draw(im)
    lignes = ["EPICERIE DEMO", "Facture 555", "Date 2026-09-02",
              "Pommes  3.00", "Pain    2.50", "TOTAL   5.50"]
    y = 20
    for l in lignes:
        d.text((30, y), l, fill="black")
        y += 40
    im.save(chemin, "JPEG")
    return chemin
