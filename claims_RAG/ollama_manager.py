"""
ollama_manager.py
------------------
The single place in this project that talks to Ollama's HTTP API.
Handles availability/model checks and wraps every generate/chat call,
so connection, timeout, and error-response problems each produce one
clear, beginner-friendly error instead of a raw traceback scattered
across the codebase.
"""

from __future__ import annotations

import requests


class OllamaUnavailableError(RuntimeError):
    """Raised when the Ollama server cannot be reached, or returns an
    error while handling a request."""


class OllamaModelNotFoundError(RuntimeError):
    """Raised when the configured model is not installed in Ollama."""


def is_reachable(base_url: str, timeout: float = 5.0) -> bool:
    try:
        resp = requests.get(f"{base_url}/api/tags", timeout=timeout)
        return resp.status_code == 200
    except requests.exceptions.RequestException:
        return False


def list_installed_models(base_url: str, timeout: float = 5.0) -> list[str]:
    resp = requests.get(f"{base_url}/api/tags", timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return [m.get("name", "") for m in data.get("models", [])]


def ensure_ready(base_url: str, model: str, timeout: float = 5.0) -> None:
    """
    Verify Ollama is reachable and the configured model is installed.
    Call this once at startup, before anything else touches Ollama.
    """
    if not is_reachable(base_url, timeout=timeout):
        raise OllamaUnavailableError(
            f"Could not reach Ollama at {base_url}\n\n"
            "Ollama does not appear to be running. To fix this:\n"
            "  1. Open a new terminal window\n"
            "  2. Run:  ollama serve\n"
            "  3. Leave that window open, then re-run this application\n\n"
            f"If Ollama is already running elsewhere, check that {base_url} "
            "is the correct address in config\\config.json."
        )

    try:
        installed = list_installed_models(base_url, timeout=timeout)
    except requests.exceptions.RequestException as exc:
        raise OllamaUnavailableError(
            f"Connected to Ollama at {base_url} but failed to list "
            f"installed models: {exc}"
        ) from exc

    if model not in installed:
        installed_list = ", ".join(installed) if installed else "(none installed)"
        raise OllamaModelNotFoundError(
            f"The model '{model}' is not installed in Ollama.\n\n"
            "To fix this, run:\n"
            f"  ollama pull {model}\n\n"
            f"Currently installed models: {installed_list}"
        )


def _post_to_ollama(url: str, payload: dict, timeout: float, action: str) -> dict:
    """Shared request/error handling for both generate() and chat()."""
    try:
        resp = requests.post(url, json=payload, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.ConnectionError as exc:
        raise OllamaUnavailableError(
            f"Lost connection to Ollama while {action}. "
            f"Is 'ollama serve' still running?\n({exc})"
        ) from exc
    except requests.exceptions.Timeout as exc:
        raise OllamaUnavailableError(
            f"Ollama did not respond within {timeout} seconds while "
            f"{action}.\n({exc})"
        ) from exc
    except requests.exceptions.HTTPError as exc:
        raise OllamaUnavailableError(
            f"Ollama returned an error while {action}: {exc}"
        ) from exc
    except requests.exceptions.RequestException as exc:
        raise OllamaUnavailableError(
            f"Unexpected error communicating with Ollama while {action}: {exc}"
        ) from exc


def generate(base_url: str, model: str, system: str, prompt: str,
             options: dict | None = None, timeout: float = 600) -> dict:
    payload = {
        "model": model,
        "system": system,
        "prompt": prompt,
        "stream": False,
        "options": options or {},
    }
    return _post_to_ollama(f"{base_url}/api/generate", payload, timeout, "generating a response")


def chat(base_url: str, model: str, messages: list, tools: list | None = None,
         options: dict | None = None, timeout: float = 600) -> dict:
    payload = {
        "model": model,
        "messages": messages,
        "stream": False,
        "options": options or {},
    }
    if tools:
        payload["tools"] = tools
    return _post_to_ollama(f"{base_url}/api/chat", payload, timeout, "running a chat call")


if __name__ == "__main__":
    from config import OLLAMA_BASE_URL, OLLAMA_MODEL

    try:
        ensure_ready(OLLAMA_BASE_URL, OLLAMA_MODEL)
        print(f"Ollama is reachable at {OLLAMA_BASE_URL}")
        print(f"Model '{OLLAMA_MODEL}' is installed and ready.")
    except (OllamaUnavailableError, OllamaModelNotFoundError) as exc:
        print(f"\n{exc}\n")
