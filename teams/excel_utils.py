"""
All reads/writes to the Excel "database" (data/teams.xlsx) live here.
Each row = one team registration. Column order below must stay in sync
with HEADERS and TEAM_FIELDS.
"""

import json
import os
import threading
from datetime import datetime

import openpyxl
from django.conf import settings

TEAMS_FILE = str(settings.TEAMS_FILE)

HEADERS = [
    "Team ID", "Team Name", "Track", "College / Organization",
    "Leader Name", "Leader Roll No", "Leader Email", "Leader Phone",
    "Members Count", "Members JSON", "Idea / Abstract", "Registered At",
    "Problem Statement ID", "Problem Statement Title",
]

TEAM_FIELDS = [
    "team_id", "team_name", "track", "college",
    "leader_name", "leader_roll_no", "leader_email", "leader_phone",
    "members_count", "members_json", "idea", "registered_at",
    "problem_id", "problem_title",
]

_lock = threading.Lock()


def _ensure_file():
    """Create the workbook with headers if it doesn't exist yet. If an
    older workbook (from before the problem-statement columns existed)
    is found, migrate it in place by appending the missing headers."""
    os.makedirs(os.path.dirname(TEAMS_FILE), exist_ok=True)
    if not os.path.exists(TEAMS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Registrations"
        ws.append(HEADERS)
        wb.save(TEAMS_FILE)
        return

    # Migration: add any headers (columns) that don't exist yet.
    wb = openpyxl.load_workbook(TEAMS_FILE)
    ws = wb.active
    existing = [c.value for c in ws[1]] if ws.max_row >= 1 else []
    missing = [h for h in HEADERS if h not in existing]
    if missing:
        start_col = len(existing) + 1
        for i, h in enumerate(missing):
            ws.cell(row=1, column=start_col + i, value=h)
        wb.save(TEAMS_FILE)


def _row_to_team(row):
    team = dict(zip(TEAM_FIELDS, row))
    try:
        team["members"] = json.loads(team.get("members_json") or "[]")
    except (json.JSONDecodeError, TypeError):
        team["members"] = []
    return team


def get_all_teams(search=None, track=None):
    _ensure_file()
    wb = openpyxl.load_workbook(TEAMS_FILE, data_only=True)
    ws = wb.active
    teams = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        team = _row_to_team(row)
        if track and track.strip().lower() != (team["track"] or "").strip().lower():
            continue
        if search:
            s = search.strip().lower()
            member_names = " ".join(m.get("name", "") for m in team["members"])
            haystack = (
                f"{team['team_name']} {team['leader_name']} "
                f"{team['leader_roll_no']} {team['college']} {member_names}"
            ).lower()
            if s not in haystack:
                continue
        teams.append(team)
    teams.sort(key=lambda t: t["team_id"], reverse=True)
    return teams


def find_team_by_name(team_name):
    _ensure_file()
    wb = openpyxl.load_workbook(TEAMS_FILE, data_only=True)
    ws = wb.active
    target = team_name.strip().lower()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[1] and str(row[1]).strip().lower() == target:
            return _row_to_team(row)
    return None


# def find_team_by_roll_no(roll_no):
#     """Look up a team by its leader's roll number (case-insensitive).
#     Used for the team-leader login."""
#     _ensure_file()
#     wb = openpyxl.load_workbook(TEAMS_FILE, data_only=True)
#     ws = wb.active
#     target = (roll_no or "").strip().lower()
#     if not target:
#         return None
#     for row in ws.iter_rows(min_row=2, values_only=True):
#         if row and row[0] and str(row[5] or "").strip().lower() == target:
#             return _row_to_team(row)
#     return None


def get_team_by_id(team_id):
    _ensure_file()
    wb = openpyxl.load_workbook(TEAMS_FILE, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] and int(row[0]) == int(team_id):
            return _row_to_team(row)
    return None


def set_team_problem(team_id, problem_id, problem_title):
    """Set (or change) the single problem statement a team has chosen.
    Passing problem_id=None clears the selection."""
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(TEAMS_FILE)
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        pid_col = header_row.index("Problem Statement ID") + 1
        ptitle_col = header_row.index("Problem Statement Title") + 1

        for row_cells in ws.iter_rows(min_row=2):
            if row_cells[0].value and int(row_cells[0].value) == int(team_id):
                ws.cell(row=row_cells[0].row, column=pid_col, value=problem_id)
                ws.cell(row=row_cells[0].row, column=ptitle_col, value=problem_title)
                try:
                    wb.save(TEAMS_FILE)
                except PermissionError:
                    raise PermissionError(
                        "teams.xlsx is currently open in Excel/OneDrive elsewhere. "
                        "Please close it and try again."
                    )
                return _row_to_team(tuple(
                    ws.cell(row=row_cells[0].row, column=c + 1).value
                    for c in range(len(TEAM_FIELDS))
                ))
        return None


def set_team_idea(team_id, idea):
    """Persist the team's idea / abstract text."""
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(TEAMS_FILE)
        ws = wb.active
        header_row = [c.value for c in ws[1]]
        idea_col = header_row.index("Idea / Abstract") + 1

        for row_cells in ws.iter_rows(min_row=2):
            if row_cells[0].value and int(row_cells[0].value) == int(team_id):
                ws.cell(row=row_cells[0].row, column=idea_col, value=idea.strip())
                try:
                    wb.save(TEAMS_FILE)
                except PermissionError:
                    raise PermissionError(
                        "teams.xlsx is currently open in Excel/OneDrive elsewhere. "
                        "Please close it and try again."
                    )
                return _row_to_team(tuple(
                    ws.cell(row=row_cells[0].row, column=c + 1).value
                    for c in range(len(TEAM_FIELDS))
                ))
        return None


def _next_team_id(ws):
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            try:
                max_id = max(max_id, int(row[0]))
            except (ValueError, TypeError):
                pass
    return max_id + 1


def save_team(data):
    """Append a new team registration. Raises ValueError on duplicate
    team name, PermissionError if the workbook is locked open elsewhere."""
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(TEAMS_FILE)
        ws = wb.active

        target = data["team_name"].strip().lower()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[1] and str(row[1]).strip().lower() == target:
                raise ValueError("A team with this name has already registered.")

        members = data.get("members", [])
        team_id = _next_team_id(ws)

        ws.append([
            team_id,
            data["team_name"].strip(),
            data.get("track", "").strip(),
            data.get("college", "").strip(),
            data["leader_name"].strip(),
            data.get("leader_roll_no", "").strip(),
            data["leader_email"].strip(),
            data["leader_phone"].strip(),
            len(members),
            json.dumps(members),
            data.get("idea", "").strip(),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            None,   # Problem Statement ID — chosen later, via team login
            None,   # Problem Statement Title
        ])

        try:
            wb.save(TEAMS_FILE)
        except PermissionError:
            raise PermissionError(
                "teams.xlsx is currently open in Excel/OneDrive elsewhere. "
                "Please close it and try again."
            )
        return team_id


def delete_team(team_id):
    _ensure_file()
    with _lock:
        wb = openpyxl.load_workbook(TEAMS_FILE)
        ws = wb.active
        for row_cells in list(ws.iter_rows(min_row=2)):
            if row_cells and row_cells[0].value and int(row_cells[0].value) == int(team_id):
                ws.delete_rows(row_cells[0].row, 1)
                try:
                    wb.save(TEAMS_FILE)
                except PermissionError:
                    raise PermissionError(
                        "teams.xlsx is currently open in Excel/OneDrive elsewhere. "
                        "Please close it and try again."
                    )
                return True
        return False


def get_dashboard_stats():
    teams = get_all_teams()
    total_teams = len(teams)
    total_participants = sum(1 + t["members_count"] for t in teams)
    by_track = {}
    for t in teams:
        key = t["track"] or "Unspecified"
        by_track[key] = by_track.get(key, 0) + 1
    return {
        "total_teams": total_teams,
        "total_participants": total_participants,
        "by_track": by_track,
    }

def find_team_by_roll_no(roll_no):
    """Look up a team by its leader's roll number (case-insensitive).
    Used for the team-leader login."""
    _ensure_file()
    wb = openpyxl.load_workbook(TEAMS_FILE, data_only=True)
    ws = wb.active
    target = (roll_no or "").strip().lower()
    if not target:
        return None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] and str(row[5] or "").strip().lower() == target:
            return _row_to_team(row)
    return None