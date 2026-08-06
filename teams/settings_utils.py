"""
Tiny persisted app settings — currently just whether the problem
statements are visible to participants (public site + team dashboard),
plus the editable public-site content.

Stored as key/value rows in the "Settings" tab of the same Google
Sheet the rest of the app uses, so it survives Render restarts without
needing a Django model/migration or local disk.

The very first time this runs against a brand-new Google Sheet, the
"Settings" tab will be empty, so _DEFAULTS below (seeded from your old
data/app_settings.json) is used until the admin edits something from
the dashboard, at which point real rows get written and take over.
"""

import json
import threading

from . import gsheet_utils

SHEET_NAME = "Settings"
HEADERS = ["Key", "Value"]

_lock = threading.Lock()

_DEFAULTS = {
    "problems_visible": True,

    "site_content": {
        "event": {
            "name": "CODE ODALREVU",
            "tagline": "6-Hour Internal Hackathon",
            "subtitle": "build something real, in one sitting, with your crew. Team registration is open.",
            "college": "Bonam Venkata Chalamayya Engineering College",
            "place": "Odalrevu",
            "dateLabel": "Hackathon Day 2026 • 10 August 2026",
            "venueLabel": "Bonam Venkata Chalamayya Engineering College Beach Road, Odalrevu — Andhra Pradesh, India",
            "status": "registrations_open",
            "minTeamSize": 1,
            "maxTeamSize": 5,
        },
        "tracks": [
            "AI / ML",
            "Web Development",
            "App Development",
            "Cybersecurity",
            "IoT & Hardware",
            "Open Innovation",
        ],
        "rules": [
            "Team size: 3–5 members",
            "Use any technology",
            "Internet allowed",
            "AI tools allowed (ChatGPT, GitHub Copilot, Claude, Gemini, etc.)",
            "Must explain the code during evaluation",
            "Work must be developed during the hackathon",
        ],
        "judging_criteria": [
            {"label": "Idea", "value": 20},
            {"label": "User Interface", "value": 15},
            {"label": "Functionality", "value": 25},
            {"label": "Innovation", "value": 15},
            {"label": "Code Quality", "value": 10},
            {"label": "Presentation", "value": 10},
            {"label": "Q&A", "value": 5},
        ],
        "deliverables": [
            "Source Code (ZIP)",
            "README",
            "PPT (3–5 slides)",
            "Live Demo",
        ],
        "presentation_format": [
            "Team introduction — 30 sec",
            "Problem — 30 sec",
            "Solution — 2 min",
            "Live demo — 3 min",
            "Future scope — 1 min",
            "Questions — 1 min",
        ],
        "how_to_register": [
            {"title": "Form a team", "desc": "1 to 5 members, one leader."},
            {"title": "Fill the form", "desc": "Team, track and contact details — takes about two minutes."},
            {"title": "Get your ticket", "desc": "Instant confirmation with a downloadable PDF ticket."},
        ],
        "coordinators": [
            {
                "name": "y.kumar",
                "role": "coordinator",
                "phone": "8121863879",
                "email": "kumaryedida339@gamil.com",
            }
        ],
    },
}


def _ws():
    return gsheet_utils.get_or_create_worksheet(SHEET_NAME, HEADERS)


def _read():
    ws = _ws()
    values = ws.get_all_values()[1:]
    data = {}
    for row in values:
        if len(row) >= 2 and row[0]:
            try:
                data[row[0]] = json.loads(row[1])
            except (json.JSONDecodeError, TypeError):
                data[row[0]] = row[1]

    merged = json.loads(json.dumps(_DEFAULTS))  # deep copy of defaults
    for key, value in data.items():
        if key == "site_content" and isinstance(value, dict):
            merged["site_content"].update(value)
        else:
            merged[key] = value
    return merged


def _write_key(key, value):
    """Upserts a single key/value row in the Settings sheet."""
    ws = _ws()
    values = ws.get_all_values()
    for i, row in enumerate(values[1:], start=2):
        if row and row[0] == key:
            ws.update_cell(i, 2, json.dumps(value))
            return
    ws.append_row([key, json.dumps(value)], value_input_option="RAW")


def get_problems_visible():
    with _lock:
        return _read()["problems_visible"]


def set_problems_visible(value):
    with _lock:
        v = bool(value)
        _write_key("problems_visible", v)
        return v


# ---- Site content (everything the admin can edit on the public site) ----

def get_site_content():
    with _lock:
        return _read()["site_content"]


def set_site_content(new_content):
    """Replaces the whole site_content object. `new_content` should be a
    dict shaped like _DEFAULTS['site_content']; any missing top-level keys
    fall back to current values so a partial payload doesn't wipe things."""
    with _lock:
        current = _read()["site_content"]
        if isinstance(new_content, dict):
            current.update(new_content)
        _write_key("site_content", current)
        return current
