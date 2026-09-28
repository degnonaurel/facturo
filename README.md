# Facturo — de la photo au tableur

**▶ Démo en ligne : [facturo-ow5a.onrender.com](https://facturo-ow5a.onrender.com)**
<sub>Hébergement gratuit : le premier chargement peut prendre jusqu'à une minute.
Démo limitée à 10 fichiers par jour et par visiteur.</sub>

Facturo lit vos **factures et reçus** (photos ou PDF, FR/EN) et les exporte
dans un **Excel propre** à deux feuilles : `Resume` (une ligne par pièce) et
`Details` (les lignes de facture). Interface web glisser-déposer bilingue,
plus un outil en ligne de commande.

<p align="center">
  <img src="exemples/apercu_facturo_fr.png" alt="Interface Facturo en français" width="49%">
  <img src="exemples/apercu_facturo_en.png" alt="Facturo interface in English" width="49%">
</p>

## Stack technique

Python 3.11 · FastAPI · LLM multimodal (API Anthropic ou OpenAI, sortie JSON
contrainte par schéma) · pdfplumber / PyMuPDF · Pillow (HEIC) · openpyxl ·
pytest · Docker · déploiement continu sur Render.

## Ce que ça fait

- **Photos ET PDF** — `.jpg .png .heic/.heif .webp` et PDF. Les photos et les
  PDF scannés passent par un modèle de vision ; les PDF texte par extraction
  directe. Aucune conversion manuelle nécessaire.
- **Extraction fiable** via LLM (Claude ou OpenAI) → JSON structuré validé.
- **Taxes québécoises** : TPS et TVQ ventilées séparément + total.
- **Regroupement** des pièces d'un même achat (reçu détaillé + relevé de paiement).
- **Ajout à un Excel existant** : les nouvelles lignes s'ajoutent à la fin,
  sans écraser — pour tenir sa compta au fil de l'eau.
- **Contrôles de cohérence** (HT+taxes ≈ TTC, somme des lignes ≈ HT) signalés
  dans une colonne *Alertes*.

## Démarrage rapide

```bash
python -m venv .venv && source .venv/bin/activate      # Windows : .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # puis renseignez ANTHROPIC_API_KEY (ou OPENAI_API_KEY)

# Interface web (http://localhost:7860)
export $(grep -v '^#' .env | xargs)     # charge la clé
python interface_web.py

# En ligne de commande
python facture_vers_excel.py exemples/factures_demo/*.pdf -o sortie.xlsx
python facture_vers_excel.py recu.heic --ajouter-a ma_compta.xlsx -o ma_compta.xlsx
```

Sans clé API, le mode hors ligne fonctionne uniquement sur les PDF texte :
`--moteur tables`.

## Arborescence

```
facturo/
├── facture_vers_excel.py   # moteur : extraction (PDF/vision) + Excel + CLI
├── interface_web.py        # app web FastAPI « Facturo » (bilingue)
├── requirements.txt
├── Dockerfile              # image de déploiement
├── LICENSE
├── tests/                  # suite de tests portable (pytest)
└── exemples/               # PDF de démo, générateur, aperçus de l'UI
```

## Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

Les tests sont autonomes (ils génèrent leurs propres entrées et simulent le
LLM) : aucun appel réseau, aucune clé requise.

## Déploiement

L'application est conteneurisée (`Dockerfile`) et se déploie telle quelle sur
toute plateforme Docker (Hugging Face Spaces, Render, Fly.io…). Elle écoute
sur `$PORT` (défaut 7860). La clé API est fournie par variable
d'environnement (`ANTHROPIC_API_KEY` ou `OPENAI_API_KEY`), configurée comme
secret de la plateforme — jamais dans le code.

Pour protéger le crédit API d'une démo publique, un quota journalier est
appliqué : `QUOTA_JOUR_TOTAL` (défaut 30 fichiers/jour) et
`QUOTA_JOUR_VISITEUR` (défaut 10 fichiers/jour par adresse IP).

## Licence

© 2026 Mohamed Luc Aurel Degnon — tous droits réservés. Le code est publié
pour consultation ; toute réutilisation, copie ou redistribution nécessite une
autorisation écrite. Voir [LICENSE](LICENSE).
