"""
helper.py
---------
Small, generic, reusable utility functions that don't belong to
Excel reading, report writing, or validation logic specifically.
"""

import json


def load_json(json_file):
    """Load and parse a JSON configuration file."""
    with open(json_file, "r") as file:
        config = json.load(file)
    return config
