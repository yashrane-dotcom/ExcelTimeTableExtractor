"""
main.py — FastAPI entry point for Timetable Extractor
Handles file upload, CORS, and routes.
"""

import uuid, os, time
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn

from parser_engine import TimetableParser
from models import UploadResponse, TeacherListResponse, TimetableResponse

# ─── App Setup ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="Timetable Extractor API",
    version="2.0.0",
    description="ML-assisted Excel timetable parsing backend",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # tighten in production
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Temporary storage ────────────────────────────────────────────────────────
UPLOAD_DIR = Path("/tmp/timetable_uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# In-memory cache: session_id → TimetableParser instance
_session_cache: dict[str, TimetableParser] = {}

# ─── Helpers ─────────────────────────────────────────────────────────────────
def get_parser(session_id: str) -> TimetableParser:
    if session_id not in _session_cache:
        raise HTTPException(404, detail=f"Session '{session_id}' not found. Please upload your file first.")
    return _session_cache[session_id]


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ok", "service": "Timetable Extractor API v2"}


@app.post("/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """
    Accept an .xlsx / .xls file, parse it, and return a session_id.
    All subsequent requests use this session_id.
    """
    ext = Path(file.filename or "").suffix.lower()
    if ext not in {".xlsx", ".xls"}:
        raise HTTPException(400, detail="Only .xlsx and .xls files are supported.")

    # Save to disk
    session_id = uuid.uuid4().hex
    dest = UPLOAD_DIR / f"{session_id}{ext}"
    contents = await file.read()
    dest.write_bytes(contents)

    # Parse
    try:
        parser = TimetableParser(str(dest))
        parser.parse()
        _session_cache[session_id] = parser
    except Exception as e:
        dest.unlink(missing_ok=True)
        raise HTTPException(422, detail=f"Failed to parse timetable: {e}")

    return UploadResponse(
        session_id=session_id,
        filename=file.filename or "",
        sheets=parser.sheet_names,
        rows=parser.row_count,
        cols=parser.col_count,
        teachers=sorted(parser.get_teacher_codes()),
        divisions=parser.get_divisions(),
    )


@app.get("/teachers/{session_id}", response_model=TeacherListResponse)
def get_teachers(session_id: str):
    """Return list of detected teacher codes + display names."""
    parser = get_parser(session_id)
    return TeacherListResponse(teachers=parser.get_teachers_with_names())


@app.get("/timetable/teacher/{session_id}/{code}", response_model=TimetableResponse)
def get_teacher_timetable(session_id: str, code: str):
    """Extract and return a teacher's weekly schedule."""
    parser = get_parser(session_id)
    code = code.upper().strip()
    if code not in parser.get_teacher_codes():
        raise HTTPException(404, detail=f"Teacher '{code}' not found.")
    schedule = parser.build_teacher_schedule(code)
    return TimetableResponse(
        label=f"{parser.faculty_display_name(code)} ({code})",
        schedule=schedule,
        slot_config=parser.slot_config_json(),
        fixed_days=parser.FIXED_DAYS,
    )


@app.get("/timetable/division/{session_id}/{div}", response_model=TimetableResponse)
def get_division_timetable(session_id: str, div: str):
    """Extract and return a full division's weekly schedule."""
    parser = get_parser(session_id)
    div = div.upper().strip()
    schedule = parser.build_division_schedule(div)
    return TimetableResponse(
        label=f"Division {div}",
        schedule=schedule,
        slot_config=parser.slot_config_json(),
        fixed_days=parser.FIXED_DAYS,
    )


@app.get("/timetable/batch/{session_id}/{div}/{batch}", response_model=TimetableResponse)
def get_batch_timetable(session_id: str, div: str, batch: str):
    """Extract and return a batch's weekly schedule."""
    parser = get_parser(session_id)
    div   = div.upper().strip()
    batch = batch.upper().strip()
    schedule = parser.build_batch_schedule(div, batch)
    return TimetableResponse(
        label=f"Division {div} · Batch {batch}",
        schedule=schedule,
        slot_config=parser.slot_config_json(),
        fixed_days=parser.FIXED_DAYS,
    )


@app.delete("/session/{session_id}")
def delete_session(session_id: str):
    """Clean up a session from memory and disk."""
    _session_cache.pop(session_id, None)
    for ext in (".xlsx", ".xls"):
        f = UPLOAD_DIR / f"{session_id}{ext}"
        f.unlink(missing_ok=True)
    return {"deleted": session_id}


# ─── Entrypoint ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
