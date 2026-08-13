"""
helper.py
---------
Small, generic, reusable utility functions that don't belong to
Excel reading, report writing, or validation logic specifically.

Currently contains:
    - load_json(): loads and parses the JSON configuration file.
"""

import json


def load_json(json_file):
    with open(json_file, "r") as file:
        config = json.load(file)
    return config
