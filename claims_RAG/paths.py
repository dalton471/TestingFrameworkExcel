"""
paths.py
--------
Determines every important project directory relative to this file's
own location, so the application works identically no matter which
drive or folder the project has been copied to.

Never hardcode an absolute path anywhere else in this project.
Always import the directory constants from this module instead.
"""

from pathlib import Path

# This file lives at: Claims_RAG/app/paths.py
# The project root is exactly one level above the app/ folder.
APP_DIR = Path(__file__).resolve().parent
BASE_DIR = APP_DIR.parent

DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
KB_DIR = DATA_DIR / "kb"
OUTPUT_DIR = DATA_DIR / "output"

CONFIG_DIR = BASE_DIR / "config"
INDEXES_DIR = BASE_DIR / "indexes"
LOGS_DIR = BASE_DIR / "logs"

ALL_PROJECT_DIRS = [
    DATA_DIR, INPUT_DIR, KB_DIR, OUTPUT_DIR,
    CONFIG_DIR, INDEXES_DIR, LOGS_DIR,
]


def ensure_project_directories() -> None:
    """Create every required project directory if it doesn't already
    exist (useful right after copying the project to a new drive)."""
    for directory in ALL_PROJECT_DIRS:
        directory.mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    ensure_project_directories()
    print(f"APP_DIR     = {APP_DIR}")
    print(f"BASE_DIR    = {BASE_DIR}")
    print(f"DATA_DIR    = {DATA_DIR}")
    print(f"INPUT_DIR   = {INPUT_DIR}")
    print(f"KB_DIR      = {KB_DIR}")
    print(f"OUTPUT_DIR  = {OUTPUT_DIR}")
    print(f"CONFIG_DIR  = {CONFIG_DIR}")
    print(f"INDEXES_DIR = {INDEXES_DIR}")
    print(f"LOGS_DIR    = {LOGS_DIR}")
