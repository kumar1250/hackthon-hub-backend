"""
All reads/writes to the Problem Statements Excel "database"
(data/problems.xlsx) live here. Mirrors excel_utils.py's pattern
but uses its own workbook so it never collides with team
registrations.
"""

import os
import threading
from datetime import datetime

import openpyxl
from django.conf import settings

PROBLEMS_FILE = str(settings.PROBLEMS_FILE)

HEADERS = [
    "ID", "Title", "Description", "Track / Domain",
    "Difficulty", "Created At", "Updated At",
]

PROBLEM_FIELDS = [
    "id", "title", "description", "track",
    "difficulty", "created_at", "updated_at",
]

_lock = threading.Lock()


def _ensure_file():
    os.makedirs(os.path.dirname(PROBLEMS_FILE), exist_ok=True)
    if not os.path.exists(PROBLEMS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Problem Statements"
        ws.append(HEADERS)
        wb.save(PROBLEMS_FILE)


def _row_to_problem(row):
    return dict(zip(PROBLEM_FIELDS, row))


def _next_id(ws):
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            try:
                max_id = max(max_id, int(row[0]))
            except (ValueError, TypeError):
                pass
    return max_id + 1


# ---------- Read ----------

def get_all_problems(search=None, track=None, difficulty=None):
    _ensure_file()
    wb = openpyxl.load_workbook(PROBLEMS_FILE, data_only=True)
    ws = wb.active
    problems = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        p = _row_to_problem(row)
        if track and track.strip().lower() != (p["track"] or "").strip().lower():
            continue
        if difficulty and difficulty.strip().lower() != (p["difficulty"] or "").strip().lower():
            continue
        if search:
            s = search.strip().lower()
            haystack = f"{p['title']} {p['description']} {p['track']}".lower()
            if s not in haystack:
                continue
        problems.append(p)
    problems.sort(key=lambda p: p["id"], reverse=True)
    return problems


def get_problem(problem_id):
    _ensure_file()
    wb = openpyxl.load_workbook(PROBLEMS_FILE, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] and int(row[0]) == int(problem_id):
            return _row_to_problem(row)
    return None


# ---------- Create ----------

def create_problem(data):
    """Append a new problem statement. Raises PermissionError if the
    workbook is open elsewhere (e.g. locked open in Excel)."""
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(PROBLEMS_FILE)
        ws = wb.active

        problem_id = _next_id(ws)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        ws.append([
            problem_id,
            data["title"].strip(),
            data.get("description", "").strip(),
            data.get("track", "").strip(),
            data.get("difficulty", "").strip(),
            now,
            now,
        ])

        try:
            wb.save(PROBLEMS_FILE)
        except PermissionError:
            raise PermissionError(
                "problems.xlsx is currently open in Excel/OneDrive elsewhere. "
                "Please close it and try again."
            )
        return problem_id


# ---------- Update ----------

def update_problem(problem_id, data):
    """Updates a row in place. Returns the updated dict, or None if
    no matching row was found."""
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(PROBLEMS_FILE)
        ws = wb.active

        for row_cells in ws.iter_rows(min_row=2):
            if row_cells[0].value and int(row_cells[0].value) == int(problem_id):
                if "title" in data and data["title"]:
                    row_cells[1].value = data["title"].strip()
                if "description" in data:
                    row_cells[2].value = (data.get("description") or "").strip()
                if "track" in data:
                    row_cells[3].value = (data.get("track") or "").strip()
                if "difficulty" in data:
                    row_cells[4].value = (data.get("difficulty") or "").strip()
                row_cells[6].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                try:
                    wb.save(PROBLEMS_FILE)
                except PermissionError:
                    raise PermissionError(
                        "problems.xlsx is currently open in Excel/OneDrive elsewhere. "
                        "Please close it and try again."
                    )
                return _row_to_problem(tuple(c.value for c in row_cells))
        return None


# ---------- Delete ----------

def delete_problem(problem_id):
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(PROBLEMS_FILE)
        ws = wb.active
        for row_cells in list(ws.iter_rows(min_row=2)):
            if row_cells[0].value and int(row_cells[0].value) == int(problem_id):
                ws.delete_rows(row_cells[0].row, 1)
                try:
                    wb.save(PROBLEMS_FILE)
                except PermissionError:
                    raise PermissionError(
                        "problems.xlsx is currently open in Excel/OneDrive elsewhere. "
                        "Please close it and try again."
                    )
                return True
        return False
    
