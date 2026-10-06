# -*- coding: utf-8 -*-
"""Tests du plan comptable : lecture, catégorisation, onglet et liste déroulante."""
from _helpers import RACINE, pdf_facture_demo  # noqa: F401
import facture_vers_excel as fve
from openpyxl import Workbook, load_workbook


def test_plan_par_defaut_bilingue():
    fr, en = fve.plan_par_defaut("fr"), fve.plan_par_defaut("en")
    assert len(fr) == len(en) and fr[0]["code"] == en[0]["code"]
    location = [c for c in fr if c["code"] == "5320"][0]
    assert location["nom"] == "Location de véhicule" and location["type"] == "Dépenses"
    assert [c for c in en if c["code"] == "5320"][0]["nom"] == "Vehicle rental"


def test_lire_plan_csv_avec_titres(tmp_path):
    p = tmp_path / "plan.csv"
    p.write_text("No compte;Libellé;Classe\n6100;Repas d'affaires;Charges\n"
                 "6200;Location auto;Charges\n", encoding="utf-8")
    plan = fve.lire_plan_comptable(str(p))
    assert plan == [{"code": "6100", "nom": "Repas d'affaires", "type": "Charges"},
                    {"code": "6200", "nom": "Location auto", "type": "Charges"}]


def test_lire_plan_excel_sans_titres(tmp_path):
    p = str(tmp_path / "plan.xlsx")
    wb = Workbook()
    for code, nom in ((6100, "Repas"), (6200, "Location auto")):
        wb.active.append([code, nom])
    wb.save(p)
    assert [c["nom"] for c in fve.lire_plan_comptable(p)] == ["Repas", "Location auto"]
    assert fve.lire_plan_comptable(p)[1]["code"] == "6200"


def test_plan_liste_de_noms(tmp_path):
    p = tmp_path / "cats.csv"
    p.write_text("Repas\nLocation auto\n", encoding="utf-8")
    plan = fve.lire_plan_comptable(str(p))
    assert plan[1] == {"code": "", "nom": "Location auto", "type": ""}
    assert fve.libelle_compte(plan, "Location auto") == "Location auto"


def test_schema_envoye_contient_les_comptes(monkeypatch):
    vu = {}

    def faux_claude(texte, images, modele, timeout, schema, consigne):
        vu["schema"], vu["consigne"] = schema, consigne
        return {}
    monkeypatch.setenv("ANTHROPIC_API_KEY", "cle-test")
    monkeypatch.setattr(fve, "_appeler_claude", faux_claude)
    plan = fve.plan_par_defaut("fr")
    fve.extraire_avec_llm(texte="x", plan=plan)
    assert "5320" in vu["schema"]["properties"]["compte"]["enum"]
    assert "compte" in vu["schema"]["required"]
    assert "Location de véhicule" in vu["consigne"]
    assert "compte" not in fve._SCHEMA_FACTURE["properties"]   # schéma de base intact


def test_categorisation_bout_en_bout(tmp_path, monkeypatch):
    """Location auto → « 5320 · Location de véhicule » dans la colonne
    Catégorie, onglet plan + liste déroulante, plan réutilisé à l'ajout."""
    pdf = pdf_facture_demo(str(tmp_path / "f.pdf"))
    faux = {"fournisseur": "Enterprise", "numero": "R-1", "date": "2026-09-01",
            "devise": "CAD", "total_ht": 100.0, "tps": 5.0, "tvq": 9.98,
            "tva": 14.98, "total_ttc": 114.98, "lignes": [], "compte": "5320"}
    monkeypatch.setenv("ANTHROPIC_API_KEY", "cle-test")
    monkeypatch.setattr(fve, "extraire_avec_llm", lambda *a, **k: faux)
    plan = fve.plan_par_defaut("fr")
    f = fve.traiter_facture(pdf, moteur="llm", plan=plan)
    assert f.categorie == "5320 · Location de véhicule"

    base = str(tmp_path / "compta.xlsx")
    fve.construire_excel([f], base, plan=plan)
    wb = load_workbook(base)
    assert wb.sheetnames == ["Resume", "Details", "Plan comptable"]
    ws = wb["Resume"]
    col = [c.value for c in ws[1]].index("Catégorie") + 1
    assert ws.cell(2, col).value == "5320 · Location de véhicule"
    dv = ws.data_validations.dataValidation
    assert len(dv) == 1 and "'Plan comptable'!$D$2" in dv[0].formula1

    # Ajout en anglais sans fournir de plan : celui du classeur est repris.
    assert [c["code"] for c in fve.plan_du_classeur(base)] == [c["code"] for c in plan]
    fve.construire_excel([f], base, base_excel=base, langue="en",
                         plan=fve.plan_du_classeur(base))
    wb = load_workbook(base)
    assert wb.sheetnames == ["Summary", "Details", "Chart of accounts"]
    dv = wb["Summary"].data_validations.dataValidation
    assert len(dv) == 1 and "'Chart of accounts'!" in dv[0].formula1
