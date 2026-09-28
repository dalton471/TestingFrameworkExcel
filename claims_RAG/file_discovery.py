"""
file_discovery.py
------------------
Automatically finds supported data files inside the project's own
data directories, so the application never needs a manually typed
or hardcoded absolute file path.
"""

from pathlib import Path
from typing import List, Optional

from paths import KB_DIR, INPUT_DIR

# .txt is intentionally not included: there is no loader for it in
# rag_claims_kb.py's LOADER_REGISTRY, and advertising support for a
# format the application cannot actually read would be misleading.
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".pdf", ".docx", ".html", ".htm", ".md"}


class NoSupportedFilesFoundError(FileNotFoundError):
    """Raised when a data directory contains no file with a supported extension."""


class MultipleFilesFoundError(FileNotFoundError):
    """Raised when more than one supported file is found and the
    application cannot safely guess which one to use."""


def list_supported_files(directory: Path) -> List[Path]:
    """Every file directly inside `directory` with a supported
    extension, sorted by name. Does not look inside subfolders."""
    if not directory.exists():
        return []
    return sorted(
        p for p in directory.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def find_single_file(directory: Path, purpose: str, explicit_path: Optional[str] = None) -> Path:
    """
    Resolve exactly one file to use for `purpose` (e.g. "knowledge base").

    - If `explicit_path` is given, it is resolved and returned directly.
    - Otherwise `directory` is scanned:
        - exactly one supported file found -> used automatically
        - zero found  -> a clear error names the exact folder to place a file in
        - more than one found -> a clear error lists every candidate;
          nothing is ever silently guessed
    """
    if explicit_path:
        resolved = Path(explicit_path).expanduser().resolve()
        if not resolved.exists():
            raise FileNotFoundError(f"The specified {purpose} file does not exist: {resolved}")
        return resolved

    candidates = list_supported_files(directory)

    if len(candidates) == 0:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise NoSupportedFilesFoundError(
            f"No supported {purpose} file was found.\n"
            f"Please place one file in: {directory}\n"
            f"Supported file types: {supported}"
        )

    if len(candidates) > 1:
        listed = "\n".join(f"  - {p.name}" for p in candidates)
        raise MultipleFilesFoundError(
            f"Multiple possible {purpose} files were found in:\n{directory}\n\n"
            f"{listed}\n\n"
            f"Please keep only one file in this folder, or pass the exact "
            f"file path explicitly to choose which one to use."
        )

    return candidates[0]


def find_kb_file(explicit_path: Optional[str] = None) -> Path:
    """Locate the single Claims Knowledge Base file, from data/kb/."""
    return find_single_file(KB_DIR, purpose="knowledge base", explicit_path=explicit_path)


def find_input_file(explicit_path: Optional[str] = None) -> Path:
    """Locate a single input file, from data/input/. Not consumed by
    rag_claims_kb.py yet (it takes its input as a typed question, not
    a file) - provided here for later use."""
    return find_single_file(INPUT_DIR, purpose="input", explicit_path=explicit_path)


if __name__ == "__main__":
    try:
        kb_path = find_kb_file()
        print(f"Discovered KB file: {kb_path}")
    except (NoSupportedFilesFoundError, MultipleFilesFoundError) as exc:
        print(str(exc))
