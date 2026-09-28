# -*- coding: utf-8 -*-
"""Tests du moteur : extraction, regroupement, ajout à un Excel existant."""
from _helpers import RACINE, pdf_facture_demo  # noqa: F401  (ajuste sys.path)
import facture_vers_excel as fve
from openpyxl import load_workbook


def _fac(nom, four, ttc, n_lignes, date="2026-09-01"):
    d = {"fournisseur": four, "numero": "X", "date": date, "devise": "CAD",
         "total_ht": None, "tps": None, "tvq": None, "tva": None,
         "total_ttc": ttc, "lignes": [{"description": f"l{i}", "montant": 1.0}
                                       for i in range(n_lignes)]}
    return fve._vers_facture(nom, d, "vision:test")


def test_extraction_tables_pdf(tmp_path):
    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    f = fve.traiter_facture(pdf, moteur="tables")
    assert f.moteur == "tables"
    assert f.total_ttc == 200.0
    assert len(f.lignes) >= 2


def test_extraction_llm_simulee(tmp_path, monkeypatch):
    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    faux = {"fournisseur": "Atelier Demo Inc.", "numero": "DEMO-001",
            "date": "2026-09-01", "devise": "CAD", "total_ht": 180.0,
            "tps": None, "tvq": None, "tva": 20.0, "total_ttc": 200.0,
            "lignes": [{"description": "Service A", "quantite": 2,
                        "prix_unitaire": 50, "montant": 100},
                       {"description": "Service B", "quantite": 1,
                        "prix_unitaire": 80, "montant": 80}]}
    monkeypatch.setenv("ANTHROPIC_API_KEY", "cle-test")
    monkeypatch.setattr(fve, "extraire_avec_llm", lambda *a, **k: faux)
    f = fve.traiter_facture(pdf, moteur="llm")
    assert f.total_ttc == 200.0 and len(f.lignes) == 2
    assert not f.alertes  # HT + TVA == TTC, cohérent


def test_regroupement():
    factures = [_fac("recu.heic", "Renaissance Fripe", 151.5, 21),
                _fac("paiement.heic", "Renaissance", 151.5, 0),
                _fac("pharma.heic", "Pharmaprix", 48.42, 4, date="2026-09-02")]
    groupe = fve.regrouper_factures(factures)
    assert len(groupe) == 2
    ren = [f for f in groupe if "Renaissance" in f.fournisseur][0]
    assert ren.pieces_liees == ["paiement.heic"] and len(ren.lignes) == 21


def test_ajout_a_excel_existant(tmp_path):
    base = str(tmp_path / "compta.xlsx")
    fve.construire_excel([_fac("a.pdf", "Alpha", 10, 2)], base)
    fve.construire_excel([_fac("b.jpg", "Beta", 20, 1)], base, base_excel=base)
    wb = load_workbook(base)
    assert wb.sheetnames == ["Resume", "Details"]
    noms = [wb["Resume"].cell(i, 2).value for i in range(2, wb["Resume"].max_row + 1)]
    assert noms == ["Alpha", "Beta"]
    assert wb["Details"].max_row - 1 == 3  # 2 + 1 lignes de détail
