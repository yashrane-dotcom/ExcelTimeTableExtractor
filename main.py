"""
main.py — FastAPI entry point for Timetable Extractor v4
Adds /extract/full endpoint: full per-teacher + per-division split JSON.
"""

import uuid
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn

from parser_engine import TimetableParser
from models import UploadResponse, TeacherListResponse, TimetableResponse, CompactScheduleResponse

UPLOAD_DIR = Path("/tmp/timetable_uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

_session_cache: dict[str, TimetableParser] = {}

app = FastAPI(
    title="Timetable Extractor API",
    version="4.0.0",
    description="Extracts structured (teacher, division, room) data from Excel timetables.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_parser(session_id: str) -> TimetableParser:
    if session_id in _session_cache:
        return _session_cache[session_id]
    for ext in (".xlsx", ".xlsm"):
        path = UPLOAD_DIR / f"{session_id}{ext}"
        if path.exists():
            parser = TimetableParser(str(path))
            parser.parse()
            _session_cache[session_id] = parser
            return parser
    raise HTTPException(404, detail=f"Session '{session_id}' not found.")


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "ok", "service": "Timetable Extractor API v3"}


@app.post("/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel timetable. Returns session_id for subsequent calls."""
    ext = Path(file.filename or "").suffix.lower()
    if ext not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, detail="Only .xlsx and .xlsm files are supported.")

    session_id = uuid.uuid4().hex
    dest = UPLOAD_DIR / f"{session_id}{ext}"
    dest.write_bytes(await file.read())

    try:
        parser = TimetableParser(str(dest))
        parser.parse()
        _session_cache[session_id] = parser
    except Exception as e:
        dest.unlink(missing_ok=True)
        raise HTTPException(422, detail=f"Failed to parse: {e}")

    return UploadResponse(
        session_id=session_id,
        filename=file.filename or "",
        sheets=parser.sheet_names,
        rows=parser.row_count,
        cols=parser.col_count,
        teachers=sorted(parser.get_teacher_codes()),
        divisions=parser.get_divisions(),
    )


@app.get("/extract/{session_id}", response_model=CompactScheduleResponse)
def extract_compact(session_id: str):
    """
    Extract full timetable as compact strings with lab merging.
    Output: {"MON": {"8:30": ["UPM SE-A 301"], "10:45": ["UPM SE-A Lab-204 [2hr]"]}}
    """
    parser = get_parser(session_id)
    schedule = parser.extract_compact_schedule(merge_labs=True)
    return CompactScheduleResponse(schedule=schedule)


@app.get("/extract/full/{session_id}")
def extract_full(session_id: str):
    """
    Split master timetable into per-teacher and per-division individual schedules.
    Returns: {"teacher": {"UPM": {...}}, "division": {"SE-A": {...}}}
    """
    parser = get_parser(session_id)
    return parser.extract_full_timetable_json()


@app.get("/extract/cell")
def extract_cell(text: str = Query(..., description="Raw cell text to parse")):
    """
    Debug endpoint: parse a single cell string and return extracted entities.
    Useful for testing extraction logic without a full file.
    Example: /extract/cell?text=DBMS (UPM) SE-A Room 301
    """
    from parser_engine import extract_entities
    entities = extract_entities(text)
    return {
        "input":    text,
        "entities": entities,
        "compact":  [e["formatted"] for e in entities],
    }


@app.get("/teachers/{session_id}", response_model=TeacherListResponse)
def get_teachers(session_id: str):
    parser = get_parser(session_id)
    return TeacherListResponse(teachers=parser.get_teachers_with_names())


@app.get("/timetable/teacher/{session_id}/{code}", response_model=TimetableResponse)
def get_teacher_timetable(session_id: str, code: str):
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
    parser = get_parser(session_id)
    schedule = parser.build_division_schedule(div.upper().strip())
    return TimetableResponse(
        label=f"Division {div.upper()}",
        schedule=schedule,
        slot_config=parser.slot_config_json(),
        fixed_days=parser.FIXED_DAYS,
    )


@app.get("/timetable/batch/{session_id}/{div}/{batch}", response_model=TimetableResponse)
def get_batch_timetable(session_id: str, div: str, batch: str):
    parser = get_parser(session_id)
    schedule = parser.build_batch_schedule(div.upper().strip(), batch.upper().strip())
    return TimetableResponse(
        label=f"Division {div.upper()} · Batch {batch.upper()}",
        schedule=schedule,
        slot_config=parser.slot_config_json(),
        fixed_days=parser.FIXED_DAYS,
    )


@app.delete("/session/{session_id}")
def delete_session(session_id: str):
    _session_cache.pop(session_id, None)
    for ext in (".xlsx", ".xls"):
        f = UPLOAD_DIR / f"{session_id}{ext}"
        f.unlink(missing_ok=True)
    return {"deleted": session_id}


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
