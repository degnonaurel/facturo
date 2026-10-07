# Facturo — de la photo au tableur

[English](README.md) · **Français**

**▶ Démo en ligne : [facturo-ow5a.onrender.com](https://facturo-ow5a.onrender.com)** — application bilingue, français et anglais
<sub>Hébergement gratuit : le premier chargement peut prendre jusqu'à une minute. Démo limitée à 10 fichiers par jour et par visiteur.</sub>

Facturo lit vos **factures et reçus** (photos ou PDF, en français ou en anglais)
et les exporte dans un **Excel propre** à deux feuilles : `Resume` (une ligne
par pièce) et `Details` (les lignes de facture). Interface web glisser-déposer
avec bascule **FR/EN** en un clic, plus un outil en ligne de commande.

<p align="center">
  <img src="exemples/apercu_facturo_fr.png" alt="Interface Facturo en français" width="49%">
  <img src="exemples/apercu_facturo_en.png" alt="Facturo interface in English" width="49%">
</p>

## Stack technique

Python 3.11 · FastAPI · LLM multimodal (API Anthropic ou OpenAI, sortie JSON
contrainte par schéma) · pdfplumber / PyMuPDF · Pillow (HEIC) · openpyxl ·
pytest · Docker · déploiement continu sur Render.

## Fonctionnalités

### Usage courant

- **Traitement par lot** — déposez plusieurs factures d'un coup ; elles sont lues
  en parallèle.
- **Photos ET PDF** — `.jpg .png .heic/.heif .webp` et PDF. Les photos et les
  PDF scannés passent par un modèle de vision ; les PDF texte par extraction
  directe. Aucune conversion manuelle nécessaire.
- **Extraction fiable** via LLM → JSON structuré validé.
- **Taxes canadiennes** — TPS et TVQ ventilées séparément ; TVH et TPS+TVP
  prises en compte dans le total des taxes.
- **Regroupement** des pièces d'un même achat (reçu détaillé + relevé de paiement).
- **Contrôles de cohérence** (HT + taxes ≈ TTC, somme des lignes ≈ HT)
  signalés dans une colonne *Alertes*.
- **Bilingue** — interface et pièces en français et en anglais ; l'Excel (noms
  d'onglets, titres, alertes) suit la langue de travail, et un fichier existant
  est reconnu dans l'une ou l'autre langue.

### Pour les comptables

Catégories et totaux sont toujours dans l'Excel, sur des onglets après le
résumé : un usager courant peut les ignorer. L'import d'un plan comptable,
l'ajout à un fichier existant et l'export QuickBooks se trouvent dans le
panneau replié *Options pour comptables* de l'app web.

- **Ajout à un Excel existant** — les nouvelles lignes s'ajoutent à la fin,
  sans écraser : pour tenir sa compta au fil de l'eau.
- **Plan comptable** — chaque facture est catégorisée selon votre plan comptable
  (import Excel/CSV) ou un plan standard de PME canadienne ; une colonne
  *Catégorie* avec liste déroulante permet au comptable de vérifier et filtrer,
  et le plan est enregistré dans le classeur pour les ajouts suivants. Chaque
  article est aussi catégorisé : un reçu mixte (fournitures + repas) est
  ventilé correctement.
- **Totaux par catégorie** — un onglet dédié additionne toutes les factures par
  compte et par devise (reçus mixtes répartis au prorata des articles),
  recalculé à chaque ajout.
- **Totaux cumulés** — une ligne TOTAL par devise en bas du résumé (nombre de
  factures, HT, TPS, TVQ, taxes, TTC), recalculée à chaque ajout ; l'app web
  affiche aussi le total cumulé du fichier.
- **Export QuickBooks Online** (sur demande) — cochez *Préparer aussi un
  fichier d'import QuickBooks Online* (ou `--qbo` en ligne de commande) pour
  obtenir, **en plus** de l'Excel inchangé, un CSV au format d'import de factures
  fournisseurs de QuickBooks Online Canada (*Paramètres › Importer des données ›
  Factures*) : une ligne par article, noms de comptes de votre plan comptable
  (importez le plan exporté de QuickBooks), montants hors taxes et code de taxe
  canadien correspondant (`GST`, `HST ON`, `GST/QST QC - 9.975`, `Exempt`,
  `Out of Scope`) — fini la ressaisie. Seules les pièces de l'envoi en cours y
  figurent : un ajout ne crée jamais de doublon. Trousse de test :
  [`exemples/quickbooks/`](exemples/quickbooks/).

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
python facture_vers_excel.py *.pdf --plan comptes_qbo.xlsx -o sortie.xlsx --qbo factures_qbo.csv
```

Sans clé API, le mode hors ligne fonctionne uniquement sur les PDF texte :
`--moteur tables`.

## Arborescence

```
facturo/
├── facture_vers_excel.py   # moteur : extraction (PDF/vision) + Excel + CLI
├── interface_web.py        # app web FastAPI (bilingue FR/EN)
├── requirements.txt
├── Dockerfile              # image de déploiement
├── LICENSE
├── tests/                  # suite de tests autonome (pytest)
└── exemples/               # PDF de démo, générateur, aperçus de l'UI
```

## Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

Les tests sont autonomes (ils génèrent leurs propres entrées et simulent le
LLM) : aucun appel réseau, aucune clé requise.

## Sécurité

- Le texte lu dans les pièces est traité comme non fiable : il est échappé dans
  la page web et ne peut jamais devenir une formule active dans l'Excel ou le
  fichier QuickBooks (injection de formules / CSV).
- Limites d'envoi (taille et nombre de fichiers par requête) et quota de démo
  quotidien.
- Les fichiers produits sont effacés au bout d'une heure ; les envois sont
  traités dans des dossiers temporaires et jamais conservés.
- En-têtes de sécurité (CSP, protection contre l'intégration en cadre et le
  « MIME sniffing ») ; les Excel envoyés sont lus avec `defusedxml`.
- Messages d'erreur génériques ; les détails techniques restent dans les
  journaux du serveur.

## Déploiement

L'application est conteneurisée (`Dockerfile`) et se déploie telle quelle sur
toute plateforme Docker (Render, Fly.io, Hugging Face Spaces…). Elle écoute sur
`$PORT` (défaut 7860). La clé API est fournie par variable d'environnement
(`ANTHROPIC_API_KEY` ou `OPENAI_API_KEY`), configurée comme secret de la
plateforme — jamais dans le code.

Pour protéger le crédit API d'une démo publique, un quota journalier est
appliqué : `QUOTA_JOUR_TOTAL` (défaut 30 fichiers/jour) et
`QUOTA_JOUR_VISITEUR` (défaut 10 fichiers/jour par adresse IP). Les fichiers sont lus en
parallèle (`FACTURO_PARALLELE`, défaut 4).

## Licence

© 2026 Mohamed Luc Aurel Degnon — tous droits réservés. Le code est publié
pour consultation ; toute réutilisation, copie ou redistribution nécessite une
autorisation écrite. Voir [LICENSE](LICENSE).
