"""
parser_engine.py — Core Excel parsing intelligence.

Pipeline:
  1. Load workbook with openpyxl (preserves merged cells).
  2. Expand merged cells → full flat grid (no empty holes).
  3. Detect structure: which col = Day, which col = Time, which cols = Classes.
  4. For each cell → extract teacher codes, subject, type (lec/lab).
  5. Build schedules keyed by teacher / division / batch.

ML/heuristic layer:
  - Fuzzy column-role detection using frequency voting.
  - Teacher-code extraction with regex + known-faculty whitelist.
  - Subject cleaning via token filtering (SKIP_TOKENS set).
  - Lab-merge: consecutive same-subject lab slots → one span.
"""

from __future__ import annotations

import re, logging
from collections import defaultdict
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)

# ─── Constants ────────────────────────────────────────────────────────────────

FACULTY_MAP: dict[str, str] = {
    "UPM":  "Dr. Umesh Moharil",      "MRY":  "Dr. Meghna Yashwante",
    "MDB":  "Ms. Manisha Bhise",      "AGP":  "Dr. Amita Pal",
    "AGD":  "Dr. Anil Darekar",       "BDP":  "Dr. B D Patil",
    "PSD":  "Dr. Pratibha Desai",     "PMN":  "Dr. Poonam Nakhate",
    "MS":   "Mr. Mukesh Sharma",      "RBM":  "Mr. Rahul Mali",
    "HDV":  "Mr. Harshal Vaidya",     "PG":   "Mr. Pankaj Gaur",
    "VVK":  "Mr. Vishal Kulkarni",    "RPD":  "Mr. R P Dharmale",
    "SIB":  "Mr. Sanket Barde",       "SG":   "Dr. Sandhya Gadge",
    "NG":   "Mr. Nikhil Gurav",       "PVM":  "Mrs. Pallavi Munde",
    "SK":   "Ms. Sheetal Khande",     "CJ":   "Dr. Chhaya Joshi",
    "TP":   "Mr. Tukaram Patil",      "ST":   "Ms. Shilpa Tambe",
    "NV":   "Mrs. Neha Verma",        "SM":   "Mrs. Sonali Murumkar",
    "SB":   "Ms. Swati Bagade",       "SD":   "Mr. Shankar Deshmukh",
    "MPP":  "Mr. Martand Pandagale",
}

# Tokens that look like codes but are NOT teacher codes
SKIP_TOKENS: frozenset[str] = frozenset({
    "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN",
    "THE", "AND", "FOR", "NOT", "ARE", "WAS", "HAS",
    "LAB", "LEC", "TH", "FE", "SE", "TE", "BE",
    "DAY", "TIME", "AM", "PM", "NO", "ID", "PR", "TD", "LE",
    "A", "B", "C", "D", "E", "F", "G", "H",
    "CLASS", "DIVISION", "SLOT", "BREAK", "LUNCH", "FREE",
    "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY",
})

DAYS_SHORT     = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
DAYS_FULL_LIST = [
    "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"
]

# Canonical slot ordering — must match frontend SLOT_CONFIG keys exactly
SLOT_CONFIG = [
    {"key": "8:30",  "label": "8:30",  "type": "slot"},
    {"key": "9:30",  "label": "9:30",  "type": "slot"},
    {"key": "BRK",   "label": "Break", "type": "break"},
    {"key": "10:45", "label": "10:45", "type": "slot"},
    {"key": "11:45", "label": "11:45", "type": "slot"},
    {"key": "12:45", "label": "12:45", "type": "slot"},
    {"key": "LCH",   "label": "Lunch", "type": "lunch"},
    {"key": "1:30",  "label": "1:30",  "type": "slot"},
    {"key": "2:30",  "label": "2:30",  "type": "slot"},
    {"key": "3:30",  "label": "3:30",  "type": "slot"},
]
LECTURE_SLOTS = [s["key"] for s in SLOT_CONFIG if s["type"] == "slot"]

DIVISION_CONFIG: dict[str, dict] = {
    "A": {"name": "FE-A (Comp)",  "batches": ["A1", "A2", "A3"]},
    "B": {"name": "FE-B (Civil)", "batches": ["B1", "B2", "B3"]},
    "C": {"name": "FE-C (Comp)",  "batches": ["C1", "C2", "C3"]},
    "D": {"name": "FE-D (Mech)",  "batches": ["D1", "D2", "D3"]},
    "E": {"name": "FE-E (AI&DS)", "batches": ["E1", "E2", "E3"]},
    "F": {"name": "FE-F (Ro&AI)", "batches": ["F1", "F2", "F3"]},
    "G": {"name": "FE-G (AI&DS)", "batches": ["G1", "G2", "G3"]},
    "H": {"name": "FE-H (MTRX)",  "batches": ["H1", "H2", "H3"]},
}

FIXED_DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT"]

LAB_RE    = re.compile(r'\b(lab|pr|practical)\b', re.IGNORECASE)
CODE_RE   = re.compile(r'\b([A-Z]{2,5})\b')
TIME_RE   = re.compile(r'(\d{1,2})[:\.](\d{2})(\s*(AM|PM))?', re.IGNORECASE)
BATCH_RE  = re.compile(r'\b([A-H]\d)\b', re.IGNORECASE)


# ─── Utilities ────────────────────────────────────────────────────────────────

def normalise_time(raw: str) -> str:
    """Convert various time formats → canonical 'H:MM' (24h-lite)."""
    raw = raw.strip()
    m = TIME_RE.search(raw)
    if not m:
        return ""
    h, mn = int(m.group(1)), m.group(2)
    ampm = (m.group(4) or "").upper()
    if ampm == "PM" and h < 12:
        h += 12
    if ampm == "AM" and h == 12:
        h = 0
    return f"{h}:{mn}"


def time_to_mins(t: str) -> int:
    """'H:MM' → total minutes. Used for nearest-slot matching."""
    m = re.match(r'(\d{1,2}):(\d{2})', t)
    if not m:
        return 9999
    h, mn = int(m.group(1)), int(m.group(2))
    if h < 7:
        h += 12         # treat 1:30 as 13:30 etc.
    return h * 60 + mn


def match_day(val: str) -> str | None:
    """Return canonical 3-letter day key or None."""
    v = val.strip().upper()
    for i, d in enumerate(DAYS_SHORT):
        if v == d:
            return d
    for i, d in enumerate(DAYS_FULL_LIST):
        if v == d or v.startswith(d):
            return DAYS_SHORT[i]
    return None


def resolve_slot_key(time_norm: str) -> str | None:
    """Map a normalised time string to the nearest lecture slot key (±20 min)."""
    mins = time_to_mins(time_norm)
    best, best_diff = None, 99999
    for key in LECTURE_SLOTS:
        diff = abs(time_to_mins(key) - mins)
        if diff < best_diff:
            best_diff, best = diff, key
    return best if best_diff <= 20 else None


def extract_teacher_codes(raw: str) -> list[str]:
    """
    Extract candidate teacher codes from a cell string.
    Priority: known FACULTY_MAP codes > high-frequency unknown caps tokens.
    """
    return [m.group(1) for m in CODE_RE.finditer(raw.upper())
            if m.group(1) not in SKIP_TOKENS]


def split_cell_segments(raw: str) -> list[str]:
    """
    Split one cell's text into individual schedule segments.
    Handles newline-delimited entries and comma-splits outside parentheses.
    """
    lines = [l.strip() for l in re.split(r'[\n\r]+', raw) if l.strip()]
    result = []
    for line in lines:
        result.extend(_split_comma_outside_parens(line))
    return [s.strip() for s in result if s.strip()] or [raw.strip()]


def _split_comma_outside_parens(s: str) -> list[str]:
    parts, depth, start = [], 0, 0
    for i, ch in enumerate(s):
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ',' and depth == 0:
            parts.append(s[start:i])
            start = i + 1
    parts.append(s[start:])
    return parts


def contains_teacher(text: str, code: str) -> bool:
    """Whole-word case-insensitive check for teacher code in cell text."""
    pattern = re.compile(
        rf'(?:^|[^A-Za-z]){re.escape(code)}(?:[^A-Za-z]|$)', re.IGNORECASE
    )
    return bool(pattern.search(text))


def clean_subject(raw_seg: str, remove_codes: list[str]) -> str:
    """
    Strip teacher codes, batch codes, lab keywords, and junk from a
    cell segment to leave only the subject name.
    """
    s = raw_seg
    # Remove batch codes like A1, B2
    s = BATCH_RE.sub('', s)
    # Remove known/found teacher code mentions
    for code in remove_codes:
        s = re.sub(rf'(?:^|\s){re.escape(code)}(?:\s|$)', ' ', s, flags=re.IGNORECASE)
        s = re.sub(rf'\(\s*{re.escape(code)}\s*\)', '', s, flags=re.IGNORECASE)
    # Remove Lab/Pr/Lec keywords
    s = re.sub(r'\b(Lab|Pr|Lec|Practical|TD|LE)\b', '', s, flags=re.IGNORECASE)
    # Remove dangling numbers like room/batch numbers
    s = re.sub(r'\b\d{2,}\b', '', s)
    # Remove parentheses containing only codes
    def clean_paren(m):
        inner = m.group(1).strip()
        toks = re.split(r'[\s,·]+', inner)
        if all(re.match(r'^[A-Z]{2,5}$', t) or re.match(r'^[A-H]\d$', t) or not t
               for t in toks):
            return ''
        return m.group(0)
    s = re.sub(r'\(([^)]*)\)', clean_paren, s)
    # Clean up leftover punctuation
    s = re.sub(r'[-/,·•;]+$', '', s)
    s = re.sub(r'^[-/,·•;]+', '', s)
    s = re.sub(r'\s{2,}', ' ', s).strip()
    return s


# ─── Main Parser ──────────────────────────────────────────────────────────────

class TimetableParser:
    """
    Stateful parser: load once, query many times.

    Attributes set after parse():
      grid          : list[list[str]]  — flat 2-D grid, merged cells expanded
      row_count     : int
      col_count     : int
      sheet_names   : list[str]
      day_col       : int              — column index of day labels
      time_col      : int              — column index of time labels
      class_col_map : dict[int, str]   — col index → class/division header
      day_row_map   : dict[str, list]  — day key → list of (row_index, time_key)
      div_col_map   : dict[str, set]   — div letter → set of col indices
      _teacher_occ  : dict[str, int]   — occurrence counter for unknown codes
      _known_teachers: set[str]        — confirmed teacher codes found in file
    """

    FIXED_DAYS = FIXED_DAYS

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.grid: list[list[str]] = []
        self.row_count = 0
        self.col_count = 0
        self.sheet_names: list[str] = []
        self.day_col  = 0
        self.time_col = 1
        self.class_col_map: dict[int, str] = {}
        self.day_row_map: dict[str, list[tuple[int, str]]] = {}
        self.div_col_map: dict[str, set[int]] = {d: set() for d in DIVISION_CONFIG}
        self._teacher_occ: dict[str, int] = {}
        self._known_teachers: set[str] = set()

    # ── Step 1: Load & expand merged cells ────────────────────────────────────

    def _load_grid(self) -> list[list[str]]:
        """
        Use openpyxl to read the first sheet.
        Merged cell regions are expanded: every cell in the merge gets
        the value of the top-left cell of that merge.
        """
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        self.sheet_names = wb.sheetnames
        ws = wb.active

        # Build a dict of (row, col) → merged_value for merged ranges
        merge_map: dict[tuple[int, int], str] = {}
        for merge_range in ws.merged_cells.ranges:
            # Get value from the top-left cell of the merged region
            tl_cell = ws.cell(merge_range.min_row, merge_range.min_col)
            val = str(tl_cell.value or "").strip()
            for row in range(merge_range.min_row, merge_range.max_row + 1):
                for col in range(merge_range.min_col, merge_range.max_col + 1):
                    merge_map[(row, col)] = val

        # Build flat 2-D list
        grid: list[list[str]] = []
        for r_idx, row in enumerate(ws.iter_rows(), start=1):
            row_data: list[str] = []
            for c_idx, cell in enumerate(row, start=1):
                key = (r_idx, c_idx)
                if key in merge_map:
                    row_data.append(merge_map[key])
                else:
                    row_data.append(str(cell.value or "").strip())
            grid.append(row_data)

        return grid

    # ── Step 2: Detect column roles ───────────────────────────────────────────

    def _detect_columns(self):
        """
        Heuristic: vote for which column most often contains:
          • Day labels  → day_col
          • Time labels → time_col
        Uses frequency voting across all rows.
        """
        day_votes: dict[int, int]  = defaultdict(int)
        time_votes: dict[int, int] = defaultdict(int)

        for row in self.grid:
            for c, val in enumerate(row):
                v = val.strip().upper()
                # Day vote
                if (v in DAYS_SHORT or any(v == d or v.startswith(d) for d in DAYS_FULL_LIST)):
                    day_votes[c] += 1
                # Time vote
                if TIME_RE.match(v):
                    time_votes[c] += 1

        self.day_col  = max(day_votes,  key=day_votes.get,  default=0)
        self.time_col = max(time_votes, key=time_votes.get, default=1)
        log.info(f"[columns] day_col={self.day_col}  time_col={self.time_col}")

    # ── Step 3: Detect class/division column headers ───────────────────────────

    def _detect_class_cols(self):
        """
        Scan the first 6 rows for column headers that look like class/division labels.
        Populate self.class_col_map  and  self.div_col_map.
        """
        header_rows_to_check = min(6, len(self.grid))
        for r in range(header_rows_to_check):
            row = self.grid[r]
            row_val = row[self.day_col].strip().upper() if self.day_col < len(row) else ""
            # Likely a header row if day col is empty or says CLASS/DIVISION
            if row_val in {"", "CLASS", "DIVISION", "DAY"}:
                for c, val in enumerate(row):
                    if c in (self.day_col, self.time_col):
                        continue
                    v = val.strip()
                    if v and not v.isdigit():
                        self.class_col_map[c] = v
                        # Map to division
                        vu = v.upper()
                        for div in DIVISION_CONFIG:
                            if (vu == div or vu == f"FE-{div}" or vu == f"FE {div}"
                                    or vu.startswith(f"FE-{div} ")
                                    or f"({div})" in vu or vu.startswith(f"{div} (")):
                                self.div_col_map[div].add(c)

        # Fallback: use first row as headers
        if not self.class_col_map and self.grid:
            for c, val in enumerate(self.grid[0]):
                if c not in (self.day_col, self.time_col):
                    v = val.strip()
                    if v:
                        self.class_col_map[c] = v
        log.info(f"[class_cols] {self.class_col_map}")
        log.info(f"[div_cols]   { {k: list(v) for k, v in self.div_col_map.items()} }")

    # ── Step 4: Map days to rows ───────────────────────────────────────────────

    def _build_day_row_map(self):
        """
        Walk rows; track current day. For each time-bearing row,
        store (row_index, normalised_time_key).
        """
        cur_day = None
        self.day_row_map = {d: [] for d in FIXED_DAYS}

        for r, row in enumerate(self.grid):
            day_val  = row[self.day_col].strip()  if self.day_col  < len(row) else ""
            time_raw = row[self.time_col].strip()  if self.time_col < len(row) else ""

            dk = match_day(day_val.upper())
            if dk:
                cur_day = dk

            if cur_day and cur_day in FIXED_DAYS:
                t_norm = normalise_time(time_raw)
                if t_norm:
                    slot_key = resolve_slot_key(t_norm)
                    if slot_key:
                        self.day_row_map[cur_day].append((r, slot_key))

        log.info(f"[day_rows]  { {d: len(v) for d, v in self.day_row_map.items()} }")

    # ── Step 5: Discover all teacher codes ────────────────────────────────────

    def _discover_teachers(self):
        """
        Two-pass teacher discovery:
          Pass 1 → known FACULTY_MAP codes (always include if present).
          Pass 2 → unknown ALL-CAPS 2-5 char tokens that appear ≥2 times
                   and are not in SKIP_TOKENS.
        """
        known   = set(FACULTY_MAP.keys())
        occ: dict[str, int] = defaultdict(int)

        for row in self.grid:
            for cell in row:
                if not cell:
                    continue
                codes = extract_teacher_codes(cell)
                for code in codes:
                    if code in known:
                        self._known_teachers.add(code)
                    elif code not in SKIP_TOKENS and len(code) >= 2:
                        occ[code] += 1

        # Unknown codes appearing ≥ 2 times are likely real teacher codes
        for code, count in occ.items():
            if count >= 2:
                self._known_teachers.add(code)
        self._teacher_occ = dict(occ)
        log.info(f"[teachers] found {len(self._known_teachers)}: {sorted(self._known_teachers)}")

    # ── Public: parse() orchestrates all steps ────────────────────────────────

    def parse(self):
        log.info(f"[parse] loading {self.filepath}")
        self.grid = self._load_grid()
        if not self.grid:
            raise ValueError("Workbook appears to be empty.")
        self.row_count = len(self.grid)
        self.col_count = max(len(r) for r in self.grid)
        self._detect_columns()
        self._detect_class_cols()
        self._build_day_row_map()
        self._discover_teachers()
        log.info("[parse] complete")

    # ── Public accessors ──────────────────────────────────────────────────────

    def get_teacher_codes(self) -> set[str]:
        return self._known_teachers

    def get_teachers_with_names(self) -> list[dict]:
        return [
            {"code": c, "display_name": self.faculty_display_name(c)}
            for c in sorted(self._known_teachers)
        ]

    def faculty_display_name(self, code: str) -> str:
        return FACULTY_MAP.get(code.upper(), code.upper())

    def get_divisions(self) -> list[str]:
        return sorted(DIVISION_CONFIG.keys())

    def slot_config_json(self) -> list[dict]:
        return SLOT_CONFIG

    # ── Schedule builders ────────────────────────────────────────────────────

    def _empty_schedule(self) -> dict:
        """Return an empty day→slot→[] schedule dict."""
        return {
            day: {key: [] for key in LECTURE_SLOTS}
            for day in FIXED_DAYS
        }

    def _merge_labs(self, schedule: dict) -> dict[str, dict]:
        """
        For each day, scan consecutive slots. If the same subject appears
        as 'lab' in two consecutive slots, mark the second as skipped (span merge).
        Returns: mergeInfo[day][slot_key] = {"span": N} or {"skip": True}
        """
        merge_info: dict[str, dict] = {d: {} for d in FIXED_DAYS}
        for day in FIXED_DAYS:
            i = 0
            while i < len(LECTURE_SLOTS):
                key     = LECTURE_SLOTS[i]
                entries = schedule[day].get(key, [])
                if not entries:
                    i += 1
                    continue
                is_lab = any(e["type"] == "lab" for e in entries)
                if is_lab:
                    subj0 = entries[0].get("subject", "")
                    span  = 1
                    while i + span < len(LECTURE_SLOTS):
                        k2   = LECTURE_SLOTS[i + span]
                        ent2 = schedule[day].get(k2, [])
                        if (ent2
                                and any(e["type"] == "lab" for e in ent2)
                                and ent2[0].get("subject", "") == subj0):
                            merge_info[day][k2] = {"skip": True}
                            span += 1
                        else:
                            break
                    merge_info[day][key] = {"span": span}
                i += 1
        return merge_info

    def _schedule_to_response(self, schedule: dict, merge_info: dict) -> dict:
        """
        Convert raw schedule + merge_info into the DaySchedule model structure
        expected by the frontend.
        Slots marked skip=True in merge_info are omitted (frontend handles span).
        """
        result = {}
        for day in FIXED_DAYS:
            slots = {}
            for key in LECTURE_SLOTS:
                mi = merge_info.get(day, {}).get(key, {})
                if mi.get("skip"):
                    continue
                entries = schedule[day].get(key, [])
                slots[key] = entries
                # Attach span info to first entry if this is a lab merge
                if mi.get("span", 1) > 1 and entries:
                    for e in entries:
                        e["span"] = mi["span"]
            result[day] = {"slots": slots}
        return result

    # ─── Teacher timetable ───────────────────────────────────────────────────

    def build_teacher_schedule(self, teacher_code: str) -> dict:
        """
        Build the full weekly schedule for one teacher.
        For each row keyed by (day, slot), scan all data columns.
        If a cell contains the teacher code, parse the cell for subject + type.
        """
        schedule = self._empty_schedule()
        code_up  = teacher_code.upper()

        for day, row_entries in self.day_row_map.items():
            for row_idx, slot_key in row_entries:
                row = self.grid[row_idx]
                for c, cell_val in enumerate(row):
                    if c in (self.day_col, self.time_col) or not cell_val:
                        continue
                    if not contains_teacher(cell_val, code_up):
                        continue
                    # Parse each segment of the cell
                    for seg in split_cell_segments(cell_val):
                        if not contains_teacher(seg, code_up):
                            # Could be same cell, different segment
                            continue
                        entry = self._parse_segment_for_teacher(seg, code_up)
                        if not entry:
                            continue
                        entry["class_div"] = self.class_col_map.get(c, "")
                        # Dedup
                        if not any(e["subject"] == entry["subject"]
                                   and e["type"] == entry["type"]
                                   for e in schedule[day][slot_key]):
                            schedule[day][slot_key].append(entry)

        merge_info = self._merge_labs(schedule)
        return self._schedule_to_response(schedule, merge_info)

    def _parse_segment_for_teacher(self, seg: str, teacher_code: str) -> dict | None:
        """
        Given one text segment and the target teacher code,
        return a structured entry dict or None.
        """
        is_lab = bool(LAB_RE.search(seg))
        all_codes = [m.group(1) for m in CODE_RE.finditer(seg.upper())
                     if m.group(1) not in SKIP_TOKENS]
        subject = clean_subject(seg, all_codes)
        if not subject or len(subject) < 2 or subject.upper() in SKIP_TOKENS:
            return None
        faculty_codes = [c for c in all_codes
                         if c != teacher_code
                         and not BATCH_RE.match(c)
                         and c not in SKIP_TOKENS]
        return {
            "subject":      subject,
            "type":         "lab" if is_lab else "lec",
            "faculty_codes": faculty_codes,
            "class_div":    "",
            "raw":          seg,
        }

    # ─── Division timetable ──────────────────────────────────────────────────

    def build_division_schedule(self, div: str) -> dict:
        """
        Build the weekly schedule for an entire division.
        Includes all entries from columns assigned to that division.
        """
        schedule = self._empty_schedule()
        div_cols = self.div_col_map.get(div, set())

        for day, row_entries in self.day_row_map.items():
            for row_idx, slot_key in row_entries:
                row = self.grid[row_idx]
                for c, cell_val in enumerate(row):
                    if c in (self.day_col, self.time_col) or not cell_val:
                        continue
                    # Must belong to this division's columns
                    if div_cols and c not in div_cols:
                        continue
                    col_label = self.class_col_map.get(c, "")
                    if not div_cols:
                        # No explicit col mapping — fall back to checking header
                        vu = col_label.upper()
                        if not (vu == div or f"FE-{div}" in vu or f"FE {div}" in vu):
                            continue

                    for seg in split_cell_segments(cell_val):
                        entry = self._parse_generic_segment(seg)
                        if not entry:
                            continue
                        entry["class_div"] = col_label
                        if not any(e["subject"] == entry["subject"]
                                   and e["type"] == entry["type"]
                                   for e in schedule[day][slot_key]):
                            schedule[day][slot_key].append(entry)

        merge_info = self._merge_labs(schedule)
        return self._schedule_to_response(schedule, merge_info)

    # ─── Batch timetable ─────────────────────────────────────────────────────

    def build_batch_schedule(self, div: str, batch: str) -> dict:
        """
        Build the weekly schedule for one batch within a division.
        Logic:
          • If cell explicitly names the batch (e.g. 'A1') → include.
          • If cell has NO batch codes AND its column belongs to this div → whole-div lecture → include.
        """
        schedule  = self._empty_schedule()
        div_cols  = self.div_col_map.get(div, set())
        batch_up  = batch.upper()

        for day, row_entries in self.day_row_map.items():
            for row_idx, slot_key in row_entries:
                row = self.grid[row_idx]
                for c, cell_val in enumerate(row):
                    if c in (self.day_col, self.time_col) or not cell_val:
                        continue
                    if not self._cell_belongs_to_batch(cell_val, div, batch_up, c, div_cols):
                        continue

                    for seg in split_cell_segments(cell_val):
                        entry = self._parse_segment_for_batch(seg, div, batch_up)
                        if not entry:
                            continue
                        entry["class_div"] = self.class_col_map.get(c, "")
                        if not any(e["subject"] == entry["subject"]
                                   and e["type"] == entry["type"]
                                   for e in schedule[day][slot_key]):
                            schedule[day][slot_key].append(entry)

        merge_info = self._merge_labs(schedule)
        return self._schedule_to_response(schedule, merge_info)

    def _cell_belongs_to_batch(
        self, cell: str, div: str, batch: str, col: int, div_cols: set[int]
    ) -> bool:
        """
        Determine whether a cell's content should be included for a given batch.
        """
        # Case A: batch code explicitly present
        if re.search(rf'(?:^|[^A-Za-z0-9]){re.escape(batch)}(?:[^A-Za-z0-9]|$)',
                     cell, re.IGNORECASE):
            return True
        # Case B: cell has no explicit batch code → whole-division lecture
        if not BATCH_RE.search(cell):
            if div_cols and col in div_cols:
                return True
            col_label = self.class_col_map.get(col, "").upper()
            if col_label and (col_label == div
                              or col_label.startswith(f"FE-{div}")
                              or col_label.startswith(f"FE {div}")):
                return True
        return False

    def _parse_segment_for_batch(self, seg: str, div: str, batch: str) -> dict | None:
        """Parse a cell segment in the context of a batch."""
        s = seg.strip()
        if not s or len(s) < 2:
            return None
        is_lab = bool(LAB_RE.search(s))

        all_caps = [m.group(1) for m in CODE_RE.finditer(s.upper())]
        batch_codes_in_seg = [c for c in all_caps if BATCH_RE.match(c)]

        # If this segment explicitly names other batches but NOT ours → skip
        if batch_codes_in_seg:
            if batch.upper() not in [c.upper() for c in batch_codes_in_seg]:
                return None

        faculty_codes = [c for c in all_caps
                         if not BATCH_RE.match(c)
                         and c not in SKIP_TOKENS
                         and (c in FACULTY_MAP or (len(c) >= 2 and c in self._known_teachers))]

        subject = clean_subject(s, all_caps)
        if not subject or len(subject) < 2 or subject.upper() in SKIP_TOKENS:
            return None

        return {
            "subject":      subject,
            "type":         "lab" if is_lab else "lec",
            "faculty_codes": faculty_codes,
            "class_div":    "",
            "raw":          s,
        }

    def _parse_generic_segment(self, seg: str) -> dict | None:
        """Parse a segment without a specific teacher/batch filter."""
        s = seg.strip()
        if not s or len(s) < 2:
            return None
        is_lab     = bool(LAB_RE.search(s))
        all_caps   = [m.group(1) for m in CODE_RE.finditer(s.upper())]
        faculty    = [c for c in all_caps
                      if c not in SKIP_TOKENS
                      and not BATCH_RE.match(c)
                      and (c in FACULTY_MAP or c in self._known_teachers)]
        subject    = clean_subject(s, all_caps)
        if not subject or len(subject) < 2 or subject.upper() in SKIP_TOKENS:
            return None
        return {
            "subject":      subject,
            "type":         "lab" if is_lab else "lec",
            "faculty_codes": faculty,
            "class_div":    "",
            "raw":          s,
        }
