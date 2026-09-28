"""
validation_tools.py
--------------------
*** PLACEHOLDER - NOT THE REAL IMPLEMENTATION ***

This file exists ONLY so the application can start and the rest of
the RAG pipeline (extraction, BM25 retrieval, matching) can run and
be tested. It contains NO real Claims Denial business logic.

The original Claims Denial RAG sample's actual validation_tools.py -
which is supposed to implement:

    lookup_denial_code(code)
    lookup_allowable(cpt_code, plan_name=None)
    lookup_cpt_code(cpt_code)

- was never provided in any part of this project's source material.
Those three functions require real denial-code definitions, real
CPT/E&M categorization, and real allowable-amount reference data that
this project has no legitimate source for. Inventing plausible-looking
versions of them would silently produce wrong validation results that
look correct, which is worse than the application not running at all.

WHAT THIS PLACEHOLDER DOES INSTEAD:

TOOL_DEFINITIONS is set to an empty list. Ollama's /api/chat only
offers tools to the model when the "tools" field is present and
non-empty (see app/ollama_manager.py's chat() function) - with an
empty list, no tools are offered, so validate_instruction_against_input()
in rag_claims_kb.py runs normally, using only the CPT "Codes" sheet
reference data that is already passed to it directly in the prompt,
and the model's own general medical-billing knowledge as a fallback
(exactly as the _VALIDATE_SYSTEM prompt already instructs it to do
when reference data isn't available) - it simply never gets to call
a tool for additional lookups.

run_tool_call() is defined defensively in case something ever does
attempt a tool call despite the empty TOOL_DEFINITIONS list; it fails
loudly with a clear explanation rather than returning invented data.

TO MAKE VALIDATION TOOL LOOKUPS FULLY FUNCTIONAL:

Replace this entire file with the original project's real
validation_tools.py, implementing the three functions above against
your actual denial-code, CPT, and allowable-amount reference data.
Once that real file is in place, TOOL_DEFINITIONS should describe
those three functions in Ollama's tool-calling schema, and
run_tool_call() should dispatch to their real implementations.
"""

from __future__ import annotations

# Empty on purpose - see module docstring. This means the "TOOLS
# AVAILABLE" section described in rag_claims_kb.py's _VALIDATE_SYSTEM
# prompt has no tools actually available yet.
TOOL_DEFINITIONS: list = []


class ValidationToolNotImplementedError(RuntimeError):
    """Raised if a tool call is attempted despite none being offered
    to the model (TOOL_DEFINITIONS is empty), or if this placeholder
    file has not yet been replaced with the real implementation."""


def run_tool_call(name: str, args: dict) -> dict:
    raise ValidationToolNotImplementedError(
        f"Tool '{name}' was called with arguments {args}, but "
        f"validation_tools.py is still the placeholder from this "
        f"delivery - it does not contain real lookup_denial_code, "
        f"lookup_allowable, or lookup_cpt_code implementations. "
        f"Replace app/validation_tools.py with the original project's "
        f"real implementation to enable tool-based validation lookups."
    )
