# QuickBooks Online import — test kit

*Version française plus bas.*

Facturo can optionally produce, **in addition to** the Excel file, a CSV that
QuickBooks Online Canada imports as **bills** (supplier invoices). This kit lets
someone with a QuickBooks Online account check that the file imports cleanly.

`factures_qbo_exemple.csv` contains 6 fictional bills (8 rows) covering:

| Bill no. | Case tested | Tax code |
|---|---|---|
| PG-1042 | Québec purchase, 2 line items | `GST/QST QC - 9.975` |
| BR-778 | Québec restaurant, single line | `GST/QST QC - 9.975` |
| RH-2291 | Ontario purchase split across 2 accounts | `HST ON` |
| FPA-55 | GST only (5 %) | `GST` |
| LOY-2026-09 | No tax charged | `Exempt` |
| CN-88412 | US supplier, in USD | `Out of Scope` |

## Before importing (use a test company or the free trial)

1. **Suppliers** — create: Papeterie Gatineau, Bistro du Ruisseau,
   Rideau Hardware, Formation Pro Alberta, Gestion Immobilière Hull,
   CloudNote Inc. (CloudNote in **USD** if multicurrency is on; otherwise
   delete its row).
2. **Accounts** (expense) — make sure these exist, with these exact names:
   Office supplies, Meals and entertainment, Repairs and maintenance, Training,
   Rent, Software and subscriptions.
3. **Sales tax** must be set up (codes GST, HST ON, GST/QST QC - 9.975,
   Exempt, Out of Scope).

## Import

*Settings ⚙ › Import data › Bills* → upload `factures_qbo_exemple.csv`, then:

- date format: **dd/MM/yyyy**;
- sales tax: **Exclusive of tax** (amounts are before tax);
- map each column to the QuickBooks field of the same name. *Memo*, *Line Tax
  Amount* and *Currency Code* are optional: leave unmapped if QuickBooks does
  not offer them.

## What to check and report back

- [ ] Did the file import without errors? If not, the exact error message
  (a screenshot is perfect).
- [ ] Were all columns offered in the mapping step? Which ones were missing?
- [ ] PG-1042 shows **2 lines** on one bill; RH-2291 has one line in *Repairs
  and maintenance* and one in *Meals and entertainment*.
- [ ] Bill totals: PG-1042 = 137.97, BR-778 = 73.58, RH-2291 = 282.50,
  FPA-55 = 315.00, LOY-2026-09 = 900.00, CN-88412 = 29.00 USD.
  Off by a cent? Note which one.
- [ ] Dates read as September (not March, May, etc.).
- [ ] Are the tax code names identical in your QuickBooks? (They can vary by
  company setup.)

To regenerate the file: `python exemples/quickbooks/generer_exemple_qbo.py`.

---

# Import QuickBooks Online — trousse de test

Facturo peut produire, en option et **en plus** de l'Excel, un CSV que
QuickBooks Online Canada importe comme **factures fournisseurs** (*bills*).
Cette trousse permet à une personne qui a QuickBooks Online de vérifier que le
fichier s'importe correctement.

`factures_qbo_exemple.csv` contient 6 factures fictives (8 lignes) : achat au
Québec à 2 articles (TPS+TVQ), restaurant au Québec, achat en Ontario réparti
sur 2 comptes (TVH), TPS seule, achat exonéré, fournisseur américain en USD
(voir le tableau ci-dessus).

## Avant l'import (entreprise de test ou essai gratuit)

1. **Fournisseurs** à créer : Papeterie Gatineau, Bistro du Ruisseau,
   Rideau Hardware, Formation Pro Alberta, Gestion Immobilière Hull,
   CloudNote Inc. (en **USD** si le multidevise est activé ; sinon supprimer
   sa ligne).
2. **Comptes** de dépenses, avec ces noms exacts : Office supplies, Meals and
   entertainment, Repairs and maintenance, Training, Rent, Software and
   subscriptions.
3. **Taxes de vente** configurées (codes GST, HST ON, GST/QST QC - 9.975,
   Exempt, Out of Scope).

## Import

*Paramètres ⚙ › Importer des données › Factures* → téléverser
`factures_qbo_exemple.csv`, puis :

- format de date : **jj/MM/aaaa** ;
- taxes : **hors taxes / Exclusive** (montants avant taxes) ;
- associer chaque colonne au champ QuickBooks du même nom. *Memo*, *Line Tax
  Amount* et *Currency Code* sont facultatives : ne pas les associer si
  QuickBooks ne les propose pas.

## À vérifier et à me renvoyer

- [ ] L'import passe-t-il sans erreur ? Sinon, le message exact (une capture
  d'écran est idéale).
- [ ] Toutes les colonnes étaient-elles proposées à l'association ? Lesquelles
  manquaient ?
- [ ] PG-1042 a **2 lignes** sur une seule facture ; RH-2291 a une ligne dans
  *Repairs and maintenance* et une dans *Meals and entertainment*.
- [ ] Totaux : PG-1042 = 137,97 ; BR-778 = 73,58 ; RH-2291 = 282,50 ;
  FPA-55 = 315,00 ; LOY-2026-09 = 900,00 ; CN-88412 = 29,00 USD.
  Un écart d'un cent ? Noter laquelle.
- [ ] Les dates sont lues en septembre (pas en mars, mai, etc.).
- [ ] Les noms des codes de taxe sont-ils identiques dans votre QuickBooks ?
  (Ils peuvent varier selon la configuration.)

Pour régénérer le fichier : `python exemples/quickbooks/generer_exemple_qbo.py`.
