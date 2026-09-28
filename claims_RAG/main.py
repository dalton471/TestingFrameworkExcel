"""
main.py
-------
Project entry point. Run this from the Claims_RAG project root:

    python main.py

It adds the app/ folder to Python's import path (so the files inside
app/ can keep importing each other the simple way they already do),
then triggers rag_claims_kb.py's full startup sequence - project
directories, Ollama/model checks, KB discovery, KB loading, BM25
index building - and starts the interactive question loop.
"""

import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent / "app"
sys.path.insert(0, str(APP_DIR))

from rag_claims_kb import run_cli  # noqa: E402 (import must follow sys.path setup)

if __name__ == "__main__":
    run_cli()
