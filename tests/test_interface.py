# -*- coding: utf-8 -*-
"""Tests de l'app web (TestClient FastAPI, sans réseau ni clé)."""
import io
from _helpers import RACINE, pdf_facture_demo  # noqa: F401
from fastapi.testclient import TestClient
import interface_web
from openpyxl import load_workbook

client = TestClient(interface_web.app)


def test_page_sans_mention_moteur():
    page = client.get("/").text
    assert "Facturo" in page
    assert "Claude" not in page and "Anthropic" not in page
    assert "data-i18n" in page  # bilingue


def test_statut():
    assert "pret" in client.get("/api/statut").json()


def test_extraire_et_telecharger(tmp_path):
    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    with open(pdf, "rb") as fh:
        files = [("fichiers", ("f.pdf", fh.read(), "application/pdf"))]
    d = client.post("/api/extraire", files=files).json()
    assert len(d["factures"]) == 1 and d["download_id"]
    xl = client.get("/telecharger/" + d["download_id"]).content
    wb = load_workbook(io.BytesIO(xl))
    assert wb.sheetnames == ["Resume", "Details"]


def test_ajout_a_base_via_web(tmp_path):
    # 1) produire un premier Excel
    import facture_vers_excel as fve
    base = str(tmp_path / "base.xlsx")
    d0 = {"fournisseur": "Alpha", "numero": "1", "date": "2026-09-01",
          "devise": "CAD", "total_ht": None, "tps": None, "tvq": None,
          "tva": None, "total_ttc": 10.0, "lignes": [{"description": "x", "montant": 10}]}
    fve.construire_excel([fve._vers_facture("a.pdf", d0, "t")], base)

    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    with open(pdf, "rb") as fh, open(base, "rb") as bh:
        files = [("fichiers", ("f.pdf", fh.read(), "application/pdf")),
                 ("base", ("base.xlsx", bh.read(),
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))]
    d = client.post("/api/extraire", files=files).json()
    assert d["ajoute"] is True
    xl = client.get("/telecharger/" + d["download_id"]).content
    wb = load_workbook(io.BytesIO(xl))
    assert wb["Resume"].max_row - 1 == 2  # 1 existante + 1 nouvelle
