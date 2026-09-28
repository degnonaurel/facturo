# -*- coding: utf-8 -*-
"""Tests du chemin image/vision (LLM simulé, sans réseau ni clé réelle)."""
import base64
from _helpers import RACINE, image_recu_demo  # noqa: F401
import facture_vers_excel as fve


def test_type_source(tmp_path):
    img = image_recu_demo(str(tmp_path / "r.jpg"))
    assert fve._type_source(img) == "image"
    assert fve._type_source("x.pdf") == "pdf"


def test_preparer_image(tmp_path):
    img = image_recu_demo(str(tmp_path / "r.jpg"))
    media, b64 = fve.preparer_image(img)
    assert media == "image/jpeg"
    assert len(base64.b64decode(b64)) > 500


def test_extraction_image_simulee(tmp_path, monkeypatch):
    img = image_recu_demo(str(tmp_path / "r.jpg"))
    faux = {"fournisseur": "Epicerie Demo", "numero": "555", "date": "2026-09-02",
            "devise": "CAD", "total_ht": None, "tps": None, "tvq": None,
            "tva": None, "total_ttc": 5.50,
            "lignes": [{"description": "Pommes", "montant": 3.0},
                       {"description": "Pain", "montant": 2.5}]}

    def faux_post(url, headers=None, json=None, timeout=None):
        # vérifie qu'une image est bien envoyée dans la requête
        contenu = json["messages"][0]["content"]
        assert any(b.get("type") == "image" for b in contenu)

        class R:
            status_code = 200
            def json(self):
                return {"content": [{"type": "tool_use",
                                     "name": "enregistrer_facture", "input": faux}]}
        return R()

    monkeypatch.setenv("ANTHROPIC_API_KEY", "cle-test")
    monkeypatch.setattr("requests.post", faux_post)
    f = fve.traiter_facture(img, moteur="llm")
    assert f.moteur.startswith("vision")
    assert f.fournisseur == "Epicerie Demo" and f.total_ttc == 5.5
    assert len(f.lignes) == 2
