# Excel LLM Platform

An AI-powered tool for working with spreadsheets and commission reports. Upload Excel files, ask questions in plain English, analyze multiple files together, and process manufacturer commission reports for filing.

**Live repo:** [github.com/harnoorparmar-arch/excel-llm-platform](https://github.com/harnoorparmar-arch/excel-llm-platform)

---

## What does this do?

This platform has **three main tools**, each for a different job:

| Tool | URL | What it's for |
|------|-----|---------------|
| **Spreadsheet Intelligence** | http://localhost:8000/ | Upload a single spreadsheet, review AI-detected structure, and chat with your data |
| **Workspace** | http://localhost:8000/workspace | Upload multiple related files, find connections between them, and query them together |
| **Commission Filing** | http://localhost:8000/commission | Upload commission reports (PDF/Excel), review extracted POs, and export for filing |

All three use **Google Gemini** to understand messy, real-world spreadsheet layouts.

---

## Quick start

### 1. Prerequisites

- **Python 3.10+** (Python 3.14 works with the included `venv`)
- A **Google Gemini API key** — get one at [Google AI Studio](https://aistudio.google.com/apikey)

### 2. Install dependencies

Open a terminal in the project folder:

```powershell
cd "c:\Users\harno\OneDrive\Documents\excel-llm-platform"

# Activate the virtual environment
.\venv\Scripts\Activate.ps1

# Install packages (first time only)
pip install -r requirements.txt
```

On Mac/Linux:

```bash
cd excel-llm-platform
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure your API key

Create or edit the `.env` file in the project root:

```env
GEMINI_API_KEY=your_api_key_here
DATABASE_PATH=./data/platform.db
API_HOST=0.0.0.0
API_PORT=8000
```

> **Important:** Never commit `.env` to GitHub. It is already listed in `.gitignore`.

### 4. Start the server

**Windows (PowerShell):**

```powershell
uvicorn api.main:app --reload --reload-dir api --reload-dir parser --reload-dir frontend --reload-dir storage --host 0.0.0.0 --port 8000
```

**Mac/Linux (or Git Bash on Windows):**

```bash
./run.sh
```

### 5. Open the app

Go to **http://localhost:8000** in your browser.

API documentation (for developers): **http://localhost:8000/docs**

---

## How to use each tool

### Spreadsheet Intelligence (single file)

Best for: analyzing one workbook at a time.

1. Go to http://localhost:8000/
2. Drag and drop a spreadsheet (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`, `.tsv`, `.ods`)
3. The app parses the file and detects tables, headers, and data types
4. If anything is unclear, a **Review Required** section appears — confirm or skip AI suggestions
5. Use the **chat panel** to ask questions like:
   - *"What are total sales by region?"*
   - *"Which product had the highest revenue?"*
   - *"Forecast next quarter's sales"*

The AI reads your actual spreadsheet data to answer.

---

### Workspace (multiple files)

Best for: related spreadsheets that share columns or reference each other.

1. Go to http://localhost:8000/workspace
2. Create a workspace and upload multiple files
3. Click **Detect Relationships** — the app finds links between files (e.g. shared IDs, matching column names)
4. Review and confirm the suggested connections in the ER diagram
5. Click **Unify** to merge the data into one queryable dataset
6. Chat across all files at once

Example use case: a sales file, a product catalog, and a regional breakdown — all analyzed together.

---

### Commission Filing

Best for: manufacturer commission reports that need to be reviewed and filed.

1. Go to http://localhost:8000/commission
2. Upload one or more commission files:
   - PDF, Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`), CSV, TXT, or SLK
3. Click **Process Files** — the AI extracts purchase orders (POs), dealers, invoices, and commission amounts
4. Review each PO in the **HITL (Human-in-the-Loop)** panel:
   - **Approve** rows you're confident about
   - **Edit** fields that need correction
   - **Skip** rows you want to exclude
5. Export approved data to **CSV** for filing
6. Saved **manufacturer templates** are reused automatically on future uploads from the same vendor

Supported commission concepts include rebates, adjustments, prepaid freight, voucher rows, and multi-invoice POs.

---

## Project structure

```
excel-llm-platform/
├── api/                    # FastAPI backend
│   ├── main.py             # App entry point and page routes
│   └── routes/
│       ├── upload.py       # Single-file upload and review
│       ├── chat.py         # AI chat over workbook data
│       ├── workspace.py    # Multi-file workspaces
│       ├── commission.py   # Commission upload and export
│       └── quality.py      # Data quality reports
├── parser/                 # Core processing engines
│   ├── workbook_parser.py  # Reads Excel/CSV/ODS files
│   ├── schema_mapper.py    # AI column/table detection
│   ├── quality_engine.py   # Data quality checks
│   ├── hitl_engine.py      # Human review item detection
│   ├── relationship_engine.py  # Cross-file link detection
│   ├── unification_engine.py   # Merges related files
│   ├── forecast_engine.py  # Sales forecasting
│   ├── commission_extractor.py # Reads commission files
│   └── commission_mapper.py    # AI commission field mapping
├── frontend/               # Web UI (HTML + Tailwind CSS)
│   ├── index.html          # Single-file upload page
│   ├── workspace.html      # Multi-file workspace page
│   └── commission.html     # Commission review page
├── storage/                # SQLite database files
├── data/                   # Platform database (platform.db)
├── uploads/                # Temporary uploaded files
├── requirements.txt        # Python dependencies
├── run.sh                  # Dev server script (Mac/Linux)
└── .env                    # Your API key (not committed to git)
```

---

## API overview

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/workbooks/upload` | POST | Upload a spreadsheet |
| `/workbooks/{id}/map-schemas` | POST | Run AI schema detection |
| `/workbooks/{id}/review-items` | GET | Get items needing human review |
| `/workbooks/{id}/review-decisions` | POST | Submit review decisions |
| `/chat/query` | POST | Ask a question about a workbook |
| `/workspaces/create` | POST | Create a new workspace |
| `/workspaces/{id}/upload` | POST | Add a file to a workspace |
| `/workspaces/{id}/detect-relationships` | POST | Find links between files |
| `/workspaces/{id}/unify` | POST | Merge workspace files |
| `/workspaces/{id}/chat` | POST | Chat across workspace files |
| `/commission/upload` | POST | Process a commission file |
| `/commission/export-csv` | POST | Export approved POs to CSV |
| `/commission/templates` | GET | List saved manufacturer templates |

Full interactive docs: http://localhost:8000/docs

---

## Running tests

Make sure the server is running first, then in a second terminal:

```powershell
.\venv\Scripts\Activate.ps1
python run_integration_tests.py
python test_commission_extraction.py
```

---

## Pushing to GitHub

This project is already connected to:

```
https://github.com/harnoorparmar-arch/excel-llm-platform.git
```

To push your latest changes:

```powershell
git add .
git commit -m "Your commit message here"
git push origin main
```

Files that are **never pushed** (protected by `.gitignore`):
- `.env` (your API key)
- `venv/` (virtual environment)
- `*.db` (databases with your data)
- `*.xlsx`, `*.pdf`, `*.csv` (uploaded/test files)

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `GEMINI_API_KEY not set` | Add your key to `.env` and restart the server |
| Server won't start | Make sure `venv` is activated and run `pip install -r requirements.txt` |
| Upload fails | Check the file type is supported (see tables above) |
| Commission rows look wrong | Use the HITL review panel to edit fields, or delete the manufacturer template to force re-mapping |
| Port 8000 already in use | Stop the other process, or change `API_PORT` in `.env` and use `--port 8001` when starting |

---

## Tech stack

- **Backend:** Python, FastAPI, Uvicorn
- **AI:** Google Gemini (`google-genai`)
- **Data:** Pandas, OpenPyXL, DuckDB, SQLite
- **Frontend:** HTML, Tailwind CSS, D3.js (workspace ER diagram)
- **PDF parsing:** pdfplumber

---

## License

Private project. Contact the repository owner for usage terms.
