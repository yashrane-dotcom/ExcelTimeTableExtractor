# Timetable Extractor — Setup Guide

## Project Structure

```
timetable-extractor/
├── backend/
│   ├── main.py            ← FastAPI app + routes
│   ├── parser_engine.py   ← Core Excel parsing intelligence
│   ├── models.py          ← Pydantic request/response schemas
│   ├── utils.py           ← Shared helpers
│   └── requirements.txt
└── frontend/
    ├── index.html         ← (your existing file, unchanged)
    ├── style.css          ← (your existing file, unchanged)
    └── script.js          ← REPLACED — now uses API calls
```

---

## 1. Backend Setup

### Prerequisites
- Python 3.11+ installed
- pip

### Install dependencies

```bash
cd backend
pip install -r requirements.txt
```

### Run the server

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

The API will be live at: **http://localhost:8000**

Interactive docs (Swagger UI): **http://localhost:8000/docs**

---

## 2. Frontend Setup

The frontend is **static HTML** — no build step needed.

Just open `frontend/index.html` in your browser, **or** serve it with any static server:

```bash
# Option A: Python built-in server (from the frontend/ directory)
cd frontend
python -m http.server 5500

# Option B: VS Code Live Server (just open index.html and click "Go Live")

# Option C: npx serve
npx serve frontend/
```

> **Important:** The backend must be running on `http://localhost:8000`.
> If you change the port, update `API_BASE` at the top of `script.js`.

---

## 3. API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/upload` | Upload Excel file → returns `session_id` |
| `GET`  | `/teachers/{session_id}` | List all detected teachers |
| `GET`  | `/timetable/teacher/{session_id}/{code}` | Teacher's weekly schedule |
| `GET`  | `/timetable/division/{session_id}/{div}` | Division's weekly schedule |
| `GET`  | `/timetable/batch/{session_id}/{div}/{batch}` | Batch weekly schedule |
| `DELETE` | `/session/{session_id}` | Clean up session |

---

## 4. How the Parser Works

```
Excel File (.xlsx)
       │
       ▼
  openpyxl load
  ┌─────────────────────────────────────────────────┐
  │  STEP 1: Expand merged cells                    │
  │  Every merged region → top-left value fills all │
  └─────────────────────────────────────────────────┘
       │
       ▼
  ┌─────────────────────────────────────────────────┐
  │  STEP 2: Column role detection (ML heuristic)   │
  │  • Frequency vote → day_col, time_col           │
  │  • Pattern match → class/division header cols   │
  └─────────────────────────────────────────────────┘
       │
       ▼
  ┌─────────────────────────────────────────────────┐
  │  STEP 3: Teacher discovery                      │
  │  • Known FACULTY_MAP codes (always included)    │
  │  • Unknown 2-5 char CAPS tokens with ≥2 hits   │
  └─────────────────────────────────────────────────┘
       │
       ▼
  ┌─────────────────────────────────────────────────┐
  │  STEP 4: Cell parsing per query                 │
  │  • Split segments (newlines + commas)           │
  │  • Extract teacher/batch codes via regex        │
  │  • Clean subject name (remove codes, junk)      │
  │  • Detect lab vs lecture                        │
  └─────────────────────────────────────────────────┘
       │
       ▼
  ┌─────────────────────────────────────────────────┐
  │  STEP 5: Lab merging                            │
  │  Same subject, consecutive lab slots → span=2  │
  └─────────────────────────────────────────────────┘
       │
       ▼
     JSON response → frontend renders grid
```

---

## 5. Adding New Teachers

Edit `FACULTY_MAP` in **both** files:
- `backend/parser_engine.py`  (line ~30)
- `frontend/script.js`        (line ~10)

---

## 6. Troubleshooting

| Problem | Fix |
|---------|-----|
| CORS error in browser | Make sure backend is running and `API_BASE` in `script.js` matches |
| "No teachers found" | Your Excel may use non-standard codes; they need ≥2 occurrences |
| Wrong slot assignment | Check your Excel times match the `SLOT_CONFIG` keys (within ±20 min) |
| Merged cells missing | The backend uses openpyxl which handles this; check if file is `.xlsx` (not `.xls`) |
| Batch timetable empty | Verify batch codes like `A1`, `B2` appear literally in cells, OR that column headers clearly mark the division |

---

## 7. Production Deployment

For production, consider:
- Set `allow_origins` in CORS to your actual frontend domain
- Add a proper database instead of in-memory `_session_cache`
- Add session expiry (TTL cleanup)
- Deploy backend to Railway, Render, or a VPS
- Serve frontend via Netlify, Vercel, or Nginx
