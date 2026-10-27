# SVIET Question Paper Generation System

A Streamlit starter application for managing reusable question banks, examination templates, and generated Word papers.

## Features in this starter
- Import question banks from Excel (`.xlsx`, `.xls`) and Word (`.docx`).
- Persist banks, questions, and templates in SQLite.
- Manually classify questions by marks and eligible sections; approval is required before generation.
- Configure offered question counts, attempt counts, marks, and minimum attempts per section.
- Validate the declared maximum against the configured attempt rules.
- Generate random papers without repeating question records across sections.
- Preview the paper and reveal stored answers.
- Export a Word document.

## Requirements
- Python 3.10 or newer recommended.

## Run locally

```bash
python -m venv .venv
```

Windows PowerShell:
```powershell
.venv\Scripts\Activate.ps1
```

macOS/Linux:
```bash
source .venv/bin/activate
```

Install packages and start the app:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open the local URL shown by Streamlit, usually `http://localhost:8501`.

## Database
The app creates `data/app.db` automatically. Back up this file to preserve question banks and templates. Do not commit real examination question banks or answer keys to a public repository.

## Important scope notes
This is an initial implementation foundation, not yet a production-ready college system.
- The Word parser is best-effort; imported questions must be reviewed.
- The current data model has question records and answer text but does not yet provide a dedicated visual editor for nested subparts and OR groups.
- Section-level minimum attempt rules are validated individually. More complex cross-section rules (such as “attempt five total from B+C, at least two from each”) need an explicit combined rule in the next iteration.
- The initial Word export creates a clean structured paper, but exact pixel-level matching to the sample layout, page numbering, roll-number fields, and the official identity-disclosure note still need template refinement.
- SQLite is suitable for local development. For multiple faculty users in production, move to PostgreSQL and add authentication, roles, backups, and audit logging.

## Tests

```bash
python -m pip install pytest
pytest
```
