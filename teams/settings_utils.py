"""
Tiny persisted app settings — currently just whether the problem
statements are visible to participants (public site + team dashboard).

Stored as a small JSON file next to the Excel "databases" so it
survives restarts without needing a Django model/migration.
"""

import json
import os
import threading

from django.conf import settings

SETTINGS_FILE = os.path.join(str(settings.DATA_DIR), "app_settings.json")

_lock = threading.Lock()

_DEFAULTS = {
    # Off by default — admin has to explicitly publish the problem
    # statements before participants can see them.
    "problems_visible": False,

    # ---- Everything shown on the public site is editable by the admin.
    # This is the single source of truth for the Home page: event info,
    # the terminal card, rules/judging/deliverables/presentation-format
    # sections, tracks, "how to register" steps, and coordinators.
    "site_content": {
        "event": {
            "name": "CODE ODALREVU",
            "tagline": "24-Hour Internal Hackathon",
            "subtitle": "build something real, in one sitting, with your crew. Team registration is open.",
            "college": "Bonam Venkata Chalamayya Engineering College",
            "place": "Odalrevu",
            "dateLabel": "Hackathon Day 2026 • 22nd August 2026",
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
        "coordinators": [],
    },
}


def _read():
    if not os.path.exists(SETTINGS_FILE):
        return json.loads(json.dumps(_DEFAULTS))  # deep copy
    try:
        with open(SETTINGS_FILE, "r") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return json.loads(json.dumps(_DEFAULTS))

    merged = json.loads(json.dumps(_DEFAULTS))  # deep copy of defaults
    for key, value in data.items():
        if key == "site_content" and isinstance(value, dict):
            merged["site_content"].update(value)
        else:
            merged[key] = value
    return merged


def _write(data):
    os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
    tmp_path = SETTINGS_FILE + ".tmp"
    with open(tmp_path, "w") as f:
        json.dump(data, f)
    os.replace(tmp_path, SETTINGS_FILE)


def get_problems_visible():
    with _lock:
        return _read()["problems_visible"]


def set_problems_visible(value):
    with _lock:
        data = _read()
        data["problems_visible"] = bool(value)
        _write(data)
        return data["problems_visible"]


# ---- Site content (everything the admin can edit on the public site) ----

def get_site_content():
    with _lock:
        return _read()["site_content"]


def set_site_content(new_content):
    """Replaces the whole site_content object. `new_content` should be a
    dict shaped like _DEFAULTS['site_content']; any missing top-level keys
    fall back to current values so a partial payload doesn't wipe things."""
    with _lock:
        data = _read()
        current = data["site_content"]
        if isinstance(new_content, dict):
            current.update(new_content)
        data["site_content"] = current
        _write(data)
        return data["site_content"]