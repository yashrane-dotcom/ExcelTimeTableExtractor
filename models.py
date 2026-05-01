"""
models.py — Pydantic schemas for request/response validation.
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
    type: str                    # "lec" | "lab"
    faculty_codes: list[str]
    class_div: str
    raw: str


class DaySchedule(BaseModel):
    # key = slot key e.g. "8:30", value = list of entries
    slots: dict[str, list[SlotEntry]]


class TimetableResponse(BaseModel):
    label: str
    fixed_days: list[str]
    slot_config: list[dict[str, Any]]
    # key = day abbrev e.g. "MON"
    schedule: dict[str, DaySchedule]
