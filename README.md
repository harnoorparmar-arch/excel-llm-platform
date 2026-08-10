# Commission Filing

Upload manufacturer commission reports (PDF/Excel), review extracted POs with human-in-the-loop approval, and export approved data to CSV.

**Live repo:** [github.com/harnoorparmar-arch/excel-llm-platform](https://github.com/harnoorparmar-arch/excel-llm-platform)

---

## Deployment

This is a FastAPI app (`api.main:app`) that serves the commission UI and API. Set secrets in the host environment — do **not** commit `.env`.

### Required environment

| Variable | Required | Notes |
|----------|----------|--------|
| `GEMINI_API_KEY` | Yes | From [Google AI Studio](https://aistudio.google.com/apikey) |
| `PORT` | Often set by host | Use the platform’s port if provided |
| `API_HOST` / `API_PORT` | Optional | Local defaults; production usually uses `$PORT` |

The app also needs **writable** directories for:

- `storage/` — SQLite manufacturer templates (`storage/workbooks.db`)
- `uploads/commission/` — temporary upload files

### Production start command

```bash
pip install -r requirements.txt
uvicorn api.main:app --host 0.0.0.0 --port ${PORT:-8000}
```

Do **not** use `--reload` in production.

### Deploy on Render / Railway / similar

1. Connect the GitHub repo.
2. Set the start command to the production command above (or `python -m uvicorn api.main:app --host 0.0.0.0 --port $PORT`).
3. Add `GEMINI_API_KEY` in the platform’s environment variables.
4. Ensure the service has persistent disk (or accept that templates reset on redeploy if the filesystem is ephemeral).
5. Open the service URL — `/` redirects to `/commission`.

### Deploy with Docker (optional)

```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p storage uploads/commission
ENV PORT=8000
EXPOSE 8000
CMD uvicorn api.main:app --host 0.0.0.0 --port ${PORT}
```

Build and run:

```bash
docker build -t commission-filing .
docker run -p 8000:8000 -e GEMINI_API_KEY=your_key_here commission-filing
```

Mount volumes if you want templates/uploads to survive container restarts:

```bash
docker run -p 8000:8000 \
  -e GEMINI_API_KEY=your_key_here \
  -v "$(pwd)/storage:/app/storage" \
  -v "$(pwd)/uploads:/app/uploads" \
  commission-filing
```

### Deploy on a VPS

```bash
git clone https://github.com/harnoorparmar-arch/excel-llm-platform.git
cd excel-llm-platform
python -m venv venv
source venv/bin/activate   # Windows: .\venv\Scripts\Activate.ps1
pip install -r requirements.txt
export GEMINI_API_KEY=your_key_here
uvicorn api.main:app --host 0.0.0.0 --port 8000
```

Put Nginx (or Caddy) in front for HTTPS and reverse-proxy to `127.0.0.1:8000`. Use a process manager (`systemd`, Supervisor, or PM2) so the app restarts on reboot.

### After deploy checklist

- [ ] `GEMINI_API_KEY` is set on the host
- [ ] App opens at `/commission`
- [ ] Upload a small PDF/Excel and confirm extraction works
- [ ] Export CSV works
- [ ] `.env` / API key are not in the git repo

---

## Quick start (local)

### 1. Prerequisites

- **Python 3.10+**
- A **Google Gemini API key** — get one at [Google AI Studio](https://aistudio.google.com/apikey)

### 2. Install dependencies

**Windows (PowerShell):**

```powershell
cd "c:\Users\harno\OneDrive\Documents\excel-llm-platform"
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**Mac/Linux:**

```bash
cd excel-llm-platform
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure your API key

Create or edit `.env` in the project root:

```env
GEMINI_API_KEY=your_api_key_here
API_HOST=0.0.0.0
API_PORT=8000
```

> **Important:** Never commit `.env` to GitHub. It is already listed in `.gitignore`.

### 4. Start the server

**Windows (PowerShell):**

```powershell
python -m uvicorn api.main:app --reload --reload-dir api --reload-dir parser --reload-dir frontend --reload-dir storage --host 0.0.0.0 --port 8000
```

> If `uvicorn` is not recognized, use `python -m uvicorn` as shown above (recommended on Windows).

**Mac/Linux (or Git Bash on Windows):**

```bash
./run.sh
```

### 5. Open the app

Go to **http://localhost:8000** (redirects to `/commission`).

API docs: **http://localhost:8000/docs**

---

## How to use

1. Go to http://localhost:8000/commission
2. Upload one or more commission files:
   - PDF, Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`), CSV, TXT, or SLK
3. Click **Process Files** — the AI extracts purchase orders (POs), dealers, invoices, and commission amounts
4. Review each PO in the HITL panel:
   - **Approve** rows you're confident about
   - **Edit** fields that need correction
   - **Skip** rows you want to exclude
5. Watch **Total Commission** update as you approve, edit, or skip rows (and when you apply rebates/adjustments)
6. Export approved data to **CSV** for filing
7. Click **+ New Batch** (sidebar, Summary, or bottom bar) to clear the current review and upload another set — no page refresh needed
8. Saved **manufacturer templates** are reused automatically on future uploads from the same vendor

Supported commission concepts include rebates, adjustments, prepaid freight, voucher rows, and multi-invoice POs.

---

## Project structure

```
excel-llm-platform/
├── api/                    # FastAPI backend
│   ├── main.py             # App entry point
│   └── routes/
│       └── commission.py   # Upload, export, templates
├── parser/                 # Commission processing
│   ├── commission_extractor.py
│   └── commission_mapper.py
├── frontend/
│   └── commission.html     # HITL review UI
├── storage/                # SQLite (commission templates)
├── uploads/commission/     # Temporary uploads
├── requirements.txt
├── run.sh
└── .env                    # API key (not committed)
```

---

## API overview

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/commission/upload` | POST | Process a commission file |
| `/commission/export-csv` | POST | Export approved POs to CSV |
| `/commission/templates` | GET | List saved manufacturer templates |
| `/commission/templates/{manufacturer}` | DELETE | Delete a manufacturer template |

Full interactive docs: http://localhost:8000/docs

---

## Running tests

```powershell
.\venv\Scripts\Activate.ps1
python test_commission_extraction.py
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `GEMINI_API_KEY not set` | Add your key to `.env` (no spaces around `=`), then retry; restart the server if needed |
| `uvicorn` not recognized | Use `python -m uvicorn ...` with the venv activated |
| Server won't start | Activate `venv` and run `pip install -r requirements.txt` |
| Upload fails | Check the file type is supported |
| Commission rows look wrong | Use the HITL review panel to edit, or delete the manufacturer template to force re-mapping |
| Port already in use | Change `--port` when starting uvicorn (e.g. `--port 8001`) |

---

## Tech stack

- **Backend:** Python, FastAPI, Uvicorn
- **AI:** Google Gemini (`google-genai`)
- **Data:** Pandas, OpenPyXL, SQLite
- **Frontend:** HTML, Tailwind CSS
- **PDF parsing:** pdfplumber
