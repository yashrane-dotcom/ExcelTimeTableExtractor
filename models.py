"""
models.py — Pydantic schemas for request/response validation (v3).
"""

from pydantic import BaseModel
from typing import Any


class UploadResponse(BaseModel):
    session_id: str
    filename: str
    sheets: list[str]
    rows: int
    cols: int
    teachers: list[str]
    divisions: list[str]


class TeacherInfo(BaseModel):
    code: str
    display_name: str


class TeacherListResponse(BaseModel):
    teachers: list[TeacherInfo]


class SlotEntry(BaseModel):
    subject: str
    type: str                     # "lec" | "lab"
    faculty_codes: list[str]
    class_div: str
    room: str = ""                # NEW: extracted room number
    compact: str = ""             # NEW: "TEACHER DIV ROOM" string
    raw: str


class DaySchedule(BaseModel):
    slots: dict[str, list[SlotEntry]]


class TimetableResponse(BaseModel):
    label: str
    fixed_days: list[str]
    slot_config: list[dict[str, Any]]
    schedule: dict[str, DaySchedule]


# ── New compact response ──────────────────────────────────────────────────────

class CompactScheduleResponse(BaseModel):
    """
    Compact format: only teacher/division/room strings.
    {"MON": {"8:30": ["UPM SE-A 301", "DKP TE-B 204"]}}
    """
    schedule: dict[str, dict[str, list[str]]]
