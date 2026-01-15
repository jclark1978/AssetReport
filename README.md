# Asset Reports

Web app for cleaning Excel asset reports using the AR2_Cleanup workflow. Upload a raw `.xlsx` file, the server processes it, and you download a formatted workbook with tables, colors, and summaries.

## Features
- FastAPI backend with Excel processing via `openpyxl`
- Static frontend for upload/download
- Preserves table styles and conditional formatting
- Automatic column width sizing

## Requirements
- Python 3.10+

## Install

```bash
git clone <your-repo-url>
cd asset_reports
```

## Run (recommended)

```bash
./run_server.sh
```

This will:
- Create a virtual environment in `backend/.venv`
- Install dependencies
- Start the server on `0.0.0.0:8080`

Open `http://<server-ip>:8080` in your browser.

## Manual run

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8080
```

## Files
- `backend/app/main.py`: FastAPI app and upload endpoint
- `backend/app/processor.py`: Excel cleanup logic
- `frontend/`: static UI
- `run_server.sh`: helper script to start the server

## Troubleshooting
- If the page loads but uploads fail, check `server.log` for errors.
- Ensure port `8080` is open on your firewall/security group.
