# -*- coding: utf-8 -*-
"""Tests de l'export CSV QuickBooks Online (import de factures fournisseurs)."""
import csv
import io
from _helpers import RACINE, pdf_facture_demo  # noqa: F401
import facture_vers_excel as fve


def _facture(**champs):
    brut = {"fournisseur": "Bureau en Gros", "numero": "A-1", "date": "2026-09-03",
            "devise": "CAD", "total_ht": 100.0, "tps": 5.0, "tvq": 9.98,
            "tva": 14.98, "total_ttc": 114.98, "lignes": []}
    brut.update(champs)
    return fve._vers_facture("recu.jpg", brut, "t", fve.plan_par_defaut("fr"))


def _lire(chemin):
    with open(chemin, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def test_codes_de_taxe():
    assert fve.code_taxe_qbo(_facture()) == "GST/QST QC - 9.975"
    assert fve.code_taxe_qbo(_facture(tvq=None, tva=5.0, total_ttc=105.0)) == "GST"
    assert fve.code_taxe_qbo(_facture(tps=None, tvq=None, tva=13.0,
                                      total_ttc=113.0)) == "HST ON"
    assert fve.code_taxe_qbo(_facture(tps=None, tvq=None, tva=0,
                                      total_ttc=100.0)) == "Exempt"
    assert fve.code_taxe_qbo(_facture(devise="USD", tps=None, tvq=None, tva=None,
                                      total_ttc=100.0)) == "Out of Scope"


def test_csv_une_ligne_par_article(tmp_path):
    f = _facture(lignes=[
        {"description": "Papier", "montant": 60.0, "compte": "5240"},
        {"description": "Lunch", "montant": 40.0, "compte": "5210"}])
    chemin = fve.construire_csv_qbo([f], str(tmp_path / "qbo.csv"),
                                    plan=fve.plan_par_defaut("fr"))
    lignes = _lire(chemin)
    assert list(lignes[0].keys()) == fve.COLS_QBO
    assert [l["Account"] for l in lignes] == ["Fournitures de bureau",
                                              "Repas et représentation"]
    assert all(l["Bill no."] == "A-1" and l["Supplier"] == "Bureau en Gros"
               and l["Bill Date"] == "03/09/2026" and l["Due Date"] == "03/09/2026"
               for l in lignes)
    assert [l["Line Amount"] for l in lignes] == ["60.00", "40.00"]
    # Taxes réparties au prorata, sans perte d'arrondi.
    assert round(sum(float(l["Line Tax Amount"]) for l in lignes), 2) == 14.98


def test_csv_une_ligne_si_articles_incoherents(tmp_path):
    f = _facture(numero="", lignes=[{"description": "x", "montant": 3.0}])
    lignes = _lire(fve.construire_csv_qbo([f], str(tmp_path / "q.csv")))
    assert len(lignes) == 1 and lignes[0]["Line Amount"] == "100.00"
    assert lignes[0]["Bill no."] == "recu"          # n° absent → nom du fichier
    assert lignes[0]["Line Tax Code"] == "GST/QST QC - 9.975"


def test_csv_via_web(tmp_path):
    from fastapi.testclient import TestClient
    import interface_web
    client = TestClient(interface_web.app)
    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    with open(pdf, "rb") as fh:
        files = [("fichiers", ("f.pdf", fh.read(), "application/pdf"))]
    d = client.post("/api/extraire", files=files).json()
    r = client.get("/telecharger/" + d["download_id"] + "?format=qbo")
    assert r.status_code == 200 and "text/csv" in r.headers["content-type"]
    lignes = list(csv.DictReader(io.StringIO(r.text)))
    assert lignes and lignes[0]["Bill no."] == "DEMO-001"
    assert "QuickBooks" in client.get("/").text
