# Sample bank statement / Relevé bancaire d'exemple

`releve_demo.csv` matches the four demo invoices in `../factures_demo/`: drop
the invoices, open *Options for accountants*, import this statement, then look
at the **Bank reconciliation** sheet — 4 transactions matched, 3 expenses
without a receipt, each with a suggested account (fuel, meals, bank fees),
1 deposit.

`releve_demo.csv` correspond aux quatre factures de démonstration de
`../factures_demo/` : déposez les factures, ouvrez *Options pour comptables*,
importez ce relevé, puis consultez l'onglet **Rapprochement** — 4 transactions
rapprochées, 3 dépenses sans pièce, chacune avec un compte suggéré (carburant, repas,
frais bancaires), 1 dépôt.

```bash
python facture_vers_excel.py exemples/factures_demo/*.pdf --releve exemples/releves/releve_demo.csv -o demo.xlsx
```
