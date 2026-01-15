# AGENTS

## Project overview
This repo hosts a FastAPI backend and a static frontend to clean Excel asset reports using the AR2_Cleanup workflow. Users upload a `.xlsx`, the backend processes it, and returns a cleaned workbook.

## Key paths
- Backend app: `backend/app/main.py`
- Processing logic: `backend/app/processor.py`
- Frontend UI: `frontend/index.html`, `frontend/styles.css`, `frontend/app.js`
- Start script: `run_server.sh`
- Logs: `server.log`

## How to run
Use the provided script, which sets up a virtualenv, installs dependencies, and starts the server.

```bash
./run_server.sh
```

The server listens on `0.0.0.0:8080`.

## Notes for changes
- Keep date logic in `processor.py` consistent with AR2_Cleanup.
- Any table or style changes should be applied where the tables are created (e.g., `add_table`).
- `openpyxl` does not auto-fit columns; `autosize_columns` approximates widths.
- Conditional formatting ranges are shifted manually after inserting the header row.
