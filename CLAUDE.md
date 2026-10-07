# CLAUDE.md

Guidance for Claude Code (and any coding agent) working in this repository.

## Product

Facturo turns photos and PDFs of invoices and receipts into an accounting-ready
Excel workbook, for bilingual (English/French) small businesses in Canada.
Two levels of use, which must never get in each other's way:

- **Everyday use** — drop files, download the Excel. Nothing else on screen.
- **Accountants** (collapsed *Options for accountants* panel) — chart of
  accounts, appending to an existing workbook, bank statement reconciliation,
  QuickBooks Online export. These options never change the Excel produced by
  the everyday flow.

## Architecture

Three Python modules at the root, no heavy framework:

| Module | Role |
|---|---|
| `facture_vers_excel.py` | Engine: reading (PDF text, vision for photos and scanned PDFs), LLM extraction with a JSON schema, offline fallback, grouping, chart of accounts, Excel output, QuickBooks CSV, CLI. |
| `releves_bancaires.py` | Bank statements (OFX/QFX and CSV from Canadian banks), reconciliation with invoices, suggested accounts. |
| `interface_web.py` | FastAPI app serving a self-contained bilingual page; `GET /api/statut`, `POST /api/extraire`, `GET /telecharger/{id}`. |

The LLM returns `{fournisseur, numero, date, devise, total_ht, tps, tvq, tva,
total_ttc, lignes[]}`, plus an account (`compte`) per document and per line
when a chart of accounts is given.

## Commands

```bash
pip install -r requirements.txt -r requirements-dev.txt ruff
pytest -q          # self-contained: no network, no API key (LLM is mocked)
ruff check .       # lint (configuration in pyproject.toml)
python interface_web.py                      # http://localhost:7860
python facture_vers_excel.py invoice.pdf -o out.xlsx
```

## Conventions

- Code identifiers, comments and UI strings are in **French**; the UI is
  bilingual through the `I18N` dictionary. Public documentation is English
  first, with a French version (`README.md` + `README.fr.md`).
- Commit messages in English, imperative, one topic per commit.
- Keep dependencies minimal; justify any new one.
- Match the surrounding style; prefer extending an existing function over
  adding a parallel path.

## Quality bar

- Every behaviour change comes with a test; `pytest -q` and `ruff check .` pass
  before any push. A push to `main` redeploys the public demo.
- An unreadable file never breaks a batch: report it, process the rest.
- Every root module must be copied by the `Dockerfile` (a test checks it).
- Generated sample files (`exemples/quickbooks/`) stay in sync with the code
  (a test checks it).

## Security and privacy

- Never commit secrets; keys come from the environment (`ANTHROPIC_API_KEY`,
  `OPENAI_API_KEY`), see `.env.example`.
- Text read from documents or bank statements is untrusted: escape it in HTML
  (`esc()`), never let it become an Excel formula or a CSV formula.
- User-facing errors stay generic; details go to the server logs. The public
  interface never names the AI provider.
- Uploads are size- and count-limited, processed in temporary folders, and
  generated files expire after one hour. Customer data never goes into Git.
