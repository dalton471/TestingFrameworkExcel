"""
rag_claims_kb.py
----------------
Claims Denial KB - RAG: LLM Extraction -> Structured BM -> LLM Matching.

Pipeline:
  1. LLM field extraction   -> parses NL question into structured JSON fields FIRST.
  2. Structured BM25 query  -> builds a rich text string from extracted fields
                                and queries the BM25 index over KB rules.
  3. LLM field matching     -> strict rule-based match against BM25 candidates;
                                returns instruction VERBATIM.
  4. Validation              -> LLM validates the matched instruction against
                                the original input, optionally calling
                                validation tools for reference lookups.

Run build_index.py whenever the KB changes (if used).
"""

import json
import os
import re
import time
import numpy as np
import pandas as pd
from pathlib import Path
from rank_bm25 import BM25Okapi
from validation_tools import TOOL_DEFINITIONS, run_tool_call

from paths import ensure_project_directories
from file_discovery import find_kb_file, NoSupportedFilesFoundError, MultipleFilesFoundError
from config import OLLAMA_BASE_URL, OLLAMA_MODEL, OLLAMA_TIMEOUT, TOP_K
from ollama_manager import (
    ensure_ready, generate, chat,
    OllamaUnavailableError, OllamaModelNotFoundError,
)


# -- CONFIG -------------------------------------------------------
# Ollama's URL/model/timeout and the BM25 top_k all come from
# config/config.json via config.py - nothing is hardcoded here.
OLLAMA_BASE = OLLAMA_BASE_URL
LLM_MODEL   = OLLAMA_MODEL

ensure_project_directories()

try:
    ensure_ready(OLLAMA_BASE, LLM_MODEL, timeout=5.0)
except (OllamaUnavailableError, OllamaModelNotFoundError) as exc:
    raise SystemExit(f"\n{exc}\n")

try:
    FILEPATH = str(find_kb_file())
except (NoSupportedFilesFoundError, MultipleFilesFoundError) as exc:
    raise SystemExit(str(exc))

codepath = FILEPATH  # same workbook also contains the "Codes" sheet


# -- SHARED HELPERS -------------------------------------------------
REQUIRED_HEADERS = ("Rule ID", "Description", "Validation")


def _build_kb_entry(rule_id: str, description: str, validation: str) -> dict:
    return {"id": rule_id, "description": description, "validation": validation}


def _kb_from_tables(tables: list) -> list:
    """
    Shared row-processor used by every format plugin.
    `tables` is a list of (header, data_rows) pairs, where:
      - header    : list[str]        (column names as found in the source)
      - data_rows : list[list[str]]  (raw cell values, already stringified)
    This is the ONE place that knows how to turn raw table cells into
    KB entries: validating headers, padding short rows, and skipping
    blank/duplicate-header rows. Format plugins never do this themselves.
    """
    kb = []
    for header, data_rows in tables:
        try:
            idx_id   = header.index("Rule ID")
            idx_desc = header.index("Description")
            idx_val  = header.index("Validation")
        except ValueError:
            raise ValueError(
                f"Table headers must include: {' | '.join(REQUIRED_HEADERS)}\n"
                f"Found: {header}"
            )
        pad_to = max(idx_id, idx_desc, idx_val) + 1
        for row in data_rows:
            row = list(row) + [""] * max(0, pad_to - len(row))
            rule_id     = str(row[idx_id]   or "").strip()
            description = str(row[idx_desc] or "").strip()
            validation  = str(row[idx_val]  or "").strip()
            if not rule_id or rule_id.lower() == "rule id":
                continue
            kb.append(_build_kb_entry(rule_id, description, validation))
    return kb


# ===================================================================
# FORMAT PLUGINS
# Each plugin's only job: return raw tables as [(header, data_rows), ...]
# All validation/row-building logic lives in _kb_from_tables above.
# To support a new file type: write one function like these and add
# it to LOADER_REGISTRY at the bottom.
# ===================================================================
def _extract_tables_excel(filepath: str) -> list:
    df = pd.read_excel(filepath)
    header = [str(c) for c in df.columns]
    rows   = df.astype(str).values.tolist()
    print(f"[Loader] Excel headers found: {header}")
    return [(header, rows)]


def _extract_tables_docx(filepath: str) -> list:
    """First table in the .docx, header row + data rows."""
    try:
        from docx import Document
    except ImportError:
        raise ImportError(
            "python-docx is not installed.\n"
            "Run: pip install python-docx --break-system-packages"
        )
    print(f"[Loader] Reading Word document: {filepath}")
    doc = Document(filepath)
    if not doc.tables:
        raise ValueError(
            "No tables found in the Word document.\n"
            f"Rules must be in a table with columns: {' | '.join(REQUIRED_HEADERS)}"
        )
    rows = doc.tables[0].rows
    if len(rows) < 2:
        raise ValueError("Table has no data rows (only header or empty).")
    header = [c.text.strip() for c in rows[0].cells]
    data   = [[c.text.strip() for c in r.cells] for r in rows[1:]]
    print(f"  Word table headers found: {header}")
    return [(header, data)]


def _extract_tables_pdf(filepath: str) -> list:
    """
    Tables across all pages via pdfplumber. The header found on the first
    table is reused for any continuation tables on later pages that don't
    repeat it.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "pdfplumber is not installed.\n"
            "Run: pip install pdfplumber --break-system-packages"
        )
    print(f"[Loader] Reading PDF: {filepath}")
    tables_out = []
    header = None
    with pdfplumber.open(filepath) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            page_tables = page.extract_tables()
            if not page_tables:
                print(f"  Page {page_num}: no tables found, skipping")
                continue
            for table in page_tables:
                if not table:
                    continue
                clean = [[str(c).strip() if c else "" for c in row] for row in table]
                if header is None:
                    header = clean[0]
                    print(f"  PDF table headers found: {header}")
                    data_rows = clean[1:]
                else:
                    data_rows = clean[1:] if "Rule ID" in clean[0] else clean
                tables_out.append((header, data_rows))
    if header is None:
        raise ValueError("No tables found in the PDF.")
    return tables_out


def _extract_tables_html(filepath: str) -> list:
    """Every <table> in the file, each with its own header row."""
    from bs4 import BeautifulSoup
    print(f"[Loader] Reading HTML: {filepath}")
    with open(filepath, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
    html_tables = soup.find_all("table")
    if not html_tables:
        raise ValueError(
            "No tables found in HTML file.\n"
            f"Expected columns: {' | '.join(REQUIRED_HEADERS)}"
        )
    tables_out = []
    for table in html_tables:
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
        data   = [[c.get_text(strip=True) for c in r.find_all(["th", "td"])] for r in rows[1:]]
        print(f"  HTML table headers found: {header}")
        tables_out.append((header, data))
    return tables_out


def _extract_tables_markdown(filepath: str) -> list:
    """Every pipe-table in the markdown file."""
    print(f"[Loader] Reading Markdown: {filepath}")
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
    table_pattern = re.compile(r"(\|.+\|)\n\|[-| :]+\|\n((?:\|.+\|\n?)+)", re.MULTILINE)
    tables_out = []
    for match in table_pattern.finditer(content):
        header = [h.strip() for h in match.group(1).split("|") if h.strip()]
        data = [
            [c.strip() for c in line.split("|") if c.strip()]
            for line in match.group(2).strip().split("\n")
        ]
        print(f"  Markdown headers found: {header}")
        tables_out.append((header, data))
    if not tables_out:
        raise ValueError(
            "No valid markdown table found.\n"
            f"Expected columns: {' | '.join(REQUIRED_HEADERS)}"
        )
    return tables_out


# ===================================================================
# REGISTRY + PUBLIC LOADERS
# ===================================================================
LOADER_REGISTRY = {
    ".xlsx": _extract_tables_excel,
    ".xls":  _extract_tables_excel,
    ".docx": _extract_tables_docx,
    ".pdf":  _extract_tables_pdf,
    ".html": _extract_tables_html,
    ".htm":  _extract_tables_html,
    ".md":   _extract_tables_markdown,
}


def load_rules_from_excel(filepath: str) -> list:
    """Kept for backward compatibility; delegates to the registry."""
    return load_rules_from_file(filepath)


def load_codes_from_excel(codepath: str) -> list:
    df = pd.read_excel(codepath, sheet_name="Codes")
    return [
        {"category": str(row["Category"]).strip(), "code": str(row["Code"]).strip()}
        for _, row in df.iterrows()
    ]


def load_rules_from_file(filepath: str) -> list:
    """
    Universal knowledge base loader.
    Looks up a format plugin in LOADER_REGISTRY by file extension, gets
    raw (header, rows) tables from it, then builds KB entries via the
    single shared row-processor. To add a new format: write a plugin
    that returns [(header, data_rows), ...] and register its extension.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Knowledge base file not found: {filepath}")
    ext = Path(filepath).suffix.lower()
    extractor = LOADER_REGISTRY.get(ext)
    if extractor is None:
        raise ValueError(
            f"Unsupported file type: {ext}\n"
            f"Supported formats: {', '.join(sorted(LOADER_REGISTRY))}"
        )
    tables = extractor(filepath)
    kb = _kb_from_tables(tables)
    if not kb:
        raise ValueError(f"No valid rules found in {filepath}")
    print(f"Loaded {len(kb)} rules from {filepath}")
    return kb

# -- BUILD INDEXES ---------------------------------------------------
try:
    KB = load_rules_from_file(FILEPATH)
except (FileNotFoundError, ValueError, ImportError) as exc:
    raise SystemExit(f"\nFailed to load the knowledge base rules:\n{exc}\n")

try:
    CODES_KB = load_codes_from_excel(codepath)
except ValueError as exc:
    raise SystemExit(
        f"\nFailed to load the 'Codes' sheet from the knowledge base "
        f"workbook:\n{exc}\n\n"
        f"The workbook at {codepath} must contain a sheet named exactly "
        f"'Codes' with 'Category' and 'Code' columns.\n"
    )
except KeyError as exc:
    raise SystemExit(
        f"\nThe 'Codes' sheet is missing an expected column: {exc}\n"
        f"It must contain 'Category' and 'Code' columns.\n"
    )

bm25       = BM25Okapi([f"{e['description']} {e['validation']}".lower().split() for e in KB])
bm25_codes = BM25Okapi([f"{e['category']} {e['code']}".lower().split() for e in CODES_KB])

print(f"Loaded {len(KB)} KB rules | {len(CODES_KB)} code categories")


# -- STEP 1: LLM FIELD EXTRACTION ------------------------------------
_EXTRACT_SYSTEM = """You are a structured data extractor for a medical billing system.

Extract the following fields from the user's question and return ONLY a valid JSON object -- no markdown, no explanation.

Fields:
  - group             : e.g. "Group 1", "Group 2"          (null if not mentioned)
  - practice          : practice code e.g. "MICLAI"         (null if not mentioned)
  - insurance_company : e.g. "Medicare", "Aetna"            (null if not mentioned)
  - plan_name         : e.g. "MED", "HMO"                   (null if not mentioned)
  - cpt_codes         : list of CPT/J-codes e.g. ["99454"]  ([] if not mentioned)
  - denial_code       : numeric string e.g. "151"           (null if not mentioned)
  - remark_code       : e.g. "M25"                          (null if not mentioned)

Return ONLY the JSON object."""


def extract_fields(question: str) -> dict:
    """Use LLM to parse a natural-language question into structured billing fields."""
    t_start = time.time()
    full_json = generate(
        OLLAMA_BASE, LLM_MODEL,
        system=_EXTRACT_SYSTEM,
        prompt=f"Extract fields from:\n{question}",
        options={"temperature": 0.0},
        timeout=OLLAMA_TIMEOUT,
    )

    raw        = full_json["response"].strip()
    elapsed_ms = int((time.time() - t_start) * 1000)

    json_match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not json_match:
        raise ValueError(f"LLM did not return valid JSON.\nRaw: {raw}")

    extracted = json.loads(json_match.group())
    print(f"Extracted fields ({elapsed_ms} ms): {json.dumps(extracted)}")
    return extracted


# -- STEP 2: BM25 RETRIEVAL (rules) -----------------------------------
def retrieve_candidates(extracted: dict, k: int = TOP_K) -> list:
    query = (
        f"{extracted.get('group','')} "
        f"{extracted.get('practice','')} "
        f"{extracted.get('insurance_company','')} "
        f"{extracted.get('plan_name','')} "
        f"{' '.join(extracted.get('cpt_codes',[]))} "
        f"{extracted.get('denial_code','')} "
        f"{extracted.get('remark_code','')}"
    )
    tokens = query.lower().split()
    scores = bm25.get_scores(tokens)
    ranked = np.argsort(scores)[::-1][:k]
    results = [(KB[idx], float(scores[idx])) for idx in ranked]

    print(f"\nBM25 top-{k} candidates:")
    for i, (entry, score) in enumerate(results, 1):
        print(f"  [{i}] score={score:.4f} id={entry['id']}")
    return results


# -- STEP 2.5: BM25 RETRIEVAL (codes) ---------------------------------
def retrieve_code_candidates_from_instruction(instruction_text: str, k: int = 2) -> list:
    if not instruction_text:
        return []
    tokens = instruction_text.lower().split()
    scores = bm25_codes.get_scores(tokens)
    ranked = np.argsort(scores)[::-1][:k]
    results = [(CODES_KB[idx], float(scores[idx])) for idx in ranked]

    print(f"\nCodes BM25 from instruction top-{k} candidates:")
    if results:
        for i, (entry, score) in enumerate(results, 1):
            print(f"  [{i}] score={score:.4f} category={entry['category']} code={entry['code']}")
    else:
        print("  No relevant code categories found")
    return results


# -- STEP 3: LLM FIELD MATCHING ---------------------------------------
_MATCH_SYSTEM = """You are a medical billing rules matcher. You will be given:
  1. EXTRACTED FIELDS -- a JSON object parsed from the user question.
  2. KB CANDIDATES    -- a numbered list of Knowledge Base entries.
                        Each candidate has:
                          - 'description': rule context in natural language ending with Keywords.
                          - 'validation': "Action: <action>. <instruction text>"


FIELD ACCESS RULE:
  - If a field key is missing from EXTRACTED FIELDS, treat it exactly as if its value were null.

HOW TO READ THE DESCRIPTION:
  - The description follows this pattern:
      "This rule applies to <group>, practice <practice>, insurance: <insurance>,
       plan: <plan>. It covers CPT code(s) <cpts> with denial code(s) <denial>
       and remark code(s) <remark>."
  - Parse each value from the description before applying matching rules.
  - The INSTRUCTION to return is the Validation column in KB.

MATCHING RULES (apply strictly, field by field):
  - group             : If KB value is "All" -> always matches (regardless of extracted value).
                        If KB value is a specific value AND extracted value is null -> NO MATCH.
                        If KB value is a specific value AND extracted value is not null -> must equal KB value (case-insensitive).
                        Split the KB value on "|" and check if any segment equals the extracted value (case-insensitive)
  - practice          : Same rule as group.
  - insurance_company : Same rule as group.
  - plan_name         : Same rule as group.
  - denial_code       : If KB value is "All" -> always matches.
                        If KB value is a specific value AND extracted value is null -> NO MATCH.
                        If KB value is a specific value AND extracted value is not null -> must equal KB value exactly.
                        Split the KB value on "|" and check if any segment equals the extracted value exactly
  - remark_code       : Same rule as denial_code.
  - cpt_codes         : If KB cpt_codes contains "All" -> always matches.
                        If extracted cpt_codes is [] or missing -> always matches (user did not specify).
                        Otherwise -> at least one extracted CPT code must appear in the KB cpt_codes list.

DECISION:
  - Check every candidate in order.
  - The FIRST candidate where ALL fields satisfy the rules above is the match.
  - If a match is found, respond with ONLY the INSTRUCTION from the matched description.
  - Do NOT return the full description.
  - Do NOT return the rule scope sentence.
  - Do NOT return the matching criteria sentence.
  - Do NOT include group, practice, insurance, plan, denial code, or remark code text unless that text is part of the action sentence.
  - If NO candidate matches all rules: respond with exactly:
    No matching rule found in the KB.

STRICT OUTPUT RULES:
  - Output ONLY the instruction text or the no-match phrase. Nothing else.
  - Do not explain your reasoning.
  - Do not combine instructions from multiple entries.
  - Do not hallucinate field values or rules."""


def _format_candidates(candidates: list) -> str:
    lines = []
    for i, (entry, sim) in enumerate(candidates, 1):
        lines.append(
            f"[{i}] id={entry['id']}  score={sim:.4f}\n"
            f"    description: {entry['description']!r}"
            f"    validation: {entry['validation']!r}"
        )
    return "\n\n".join(lines)


def llm_match(extracted: dict, candidates: list) -> tuple:
    prompt = (
        f"EXTRACTED FIELDS:\n{json.dumps(extracted, indent=2)}\n\n"
        f"KB CANDIDATES:\n{_format_candidates(candidates)}"
    )

    t_start = time.time()
    full_json = generate(
        OLLAMA_BASE, LLM_MODEL,
        system=_MATCH_SYSTEM,
        prompt=prompt,
        options={"temperature": 0.0, "top_p": 0.9},
        timeout=OLLAMA_TIMEOUT,
    )
    elapsed_ms = int((time.time() - t_start) * 1000)

    answer = full_json.get("response", "")

    print(f"\n{'='*60}")
    print(f"[DEBUG] LLM MATCH RAW RESPONSE ({elapsed_ms} ms)")
    print(f"{'='*60}")
    print(f"  response length : {len(answer)}")
    print(f"  response repr   : {repr(answer[:2000])}")
    print(f"  done            : {full_json.get('done')}")
    print(f"  done_reason     : {full_json.get('done_reason')}")
    print(f"  eval_count      : {full_json.get('eval_count')}")
    print(f"  prompt_eval_count: {full_json.get('prompt_eval_count')}")

    if "<think>" in answer:
        think_end = answer.find("</think>")
        if think_end != -1:
            after_think = answer[think_end + len("</think>"):]
            print(f"  <think> block   : YES (ends at char {think_end})")
            print(f"  after </think>  : {repr(after_think[:500])}")
        else:
            print(f"  <think> block   : UNCLOSED (no </think> found)")
    else:
        print(f"  <think> block   : NONE")
    print(f"{'='*60}\n")

    return answer, elapsed_ms


# -- STEP 4: VALIDATE INSTRUCTION AGAINST INPUT -----------------------
_VALIDATE_SYSTEM = """You are a medical billing rule validator.

You will be given:
1. MATCHED RULE INSTRUCTION
2. ORIGINAL INPUT
3. CPT CODES REFERENCE (top matches from Codes sheet) - top relevant CPT categories
   retrieved by BM25 search. Use this to identify whether a CPT code belongs to
   a category like E&M. If a CPT is not covered by these results, use your own
   medical billing knowledge. Always state the source -- "from sheet" or
   "from LLM knowledge" -- in your response.

Task:
- Understand the matched rule instruction
- Compare it directly against the ORIGINAL INPUT
- If the rule mentions CPT code types (like E&M), check the CPT CODES REFERENCE first.
  If not found there, use your own medical billing knowledge as fallback.
- Decide whether the rule actually matches the data in the input

TOOLS AVAILABLE:
  - lookup_denial_code(code): returns the official description/category for
    a denial code. Call it whenever you need to confirm what a denial code
    actually means instead of guessing.
  - lookup_allowable(cpt_code, plan_name=None): returns the allowable
    reimbursement amount(s) for a CPT code, optionally narrowed to a plan.
    Call it whenever the rule requires comparing paid/allowed amounts
    (e.g. PPMAX, secondary-vs-primary allowed amount).
  - lookup_cpt_code(cpt_code): returns category, description, procedure
    name, and whether the code is flagged as E&M (is_em_code). Call it
    whenever a rule's condition depends on E&M status or on what a CPT
    code actually covers.
  Call a tool only when the matched rule instruction or original input
  actually requires that reference data. Do not call a tool speculatively.

Rules:
- The ORIGINAL INPUT may contain natural language, JSON, or both
- Use only the information explicitly present in the input
- Do not assume missing values
- If a required condition is not clearly supported, mark it as failed
- If critical information is missing, use INSUFFICIENT_DATA
- Be strict. Do not force a match.
- If the rule contains multiple branches (such as Medicare vs non-Medicare), determine the applicable branch only if the main rule conditions match
- If the rule does not match, applicable_action must be null

Return ONLY valid JSON in this exact structure:
{
  "match_status": "MATCH | NO_MATCH | INSUFFICIENT_DATA",
  "is_match": true,
  "reason": "short clear summary",
  "rule_understanding": [
    "condition 1",
    "condition 2"
  ],
  "conditions_met": [
    "..."
  ],
  "conditions_failed": [
    "..."
  ],
  "conditions_unverifiable": [
    "..."
  ],
  "applicable_action": "text or null",
  "cpt_source": "from sheet | from LLM knowledge | null",
  "cpt_source_detail": "which category/CPTs were checked and against what, or null"
}"""


def validate_instruction_against_input(
    instruction_text: str,
    question: str,
    code_candidates: list,
) -> tuple:
    """Step 4: Validate the matched instruction against the original input.

    Tool-enabled: the LLM can call lookup_denial_code / lookup_allowable /
    lookup_cpt_code (see validation_tools.py). Every tool call made is
    recorded in the returned dict's 'tools_called' list, so callers can
    confirm from the result itself whether reference data was actually
    looked up for a given query."""
    clean_instruction = (instruction_text or "").strip()

    _base = {
        "rule_understanding": [], "conditions_met": [], "conditions_failed": [],
        "conditions_unverifiable": [], "applicable_action": None,
        "cpt_source": None, "cpt_source_detail": None, "validation_performed": True,
        "tools_called": [],
    }

    if not clean_instruction:
        return {**_base, "match_status": "NO_MATCH", "is_match": False,
                "reason": "Matched instruction is empty.",
                "conditions_failed": ["Matched instruction is empty"]}, 0

    if clean_instruction == "No matching rule found in the KB.":
        return {**_base, "match_status": "NO_MATCH", "is_match": False,
                "reason": "Step 3 returned no matching KB instruction.",
                "conditions_failed": ["No matched KB instruction available"]}, 0

    codes_text = (
        json.dumps(
            [{"category": e["category"], "code": e["code"]} for e, _ in code_candidates],
            indent=2,
        )
        if code_candidates
        else "No relevant code categories found for the extracted CPT codes."
    )

    prompt = (
        f"MATCHED RULE INSTRUCTION:\n{clean_instruction}\n\n"
        f"ORIGINAL INPUT:\n{question}\n\n"
        f"CPT CODES REFERENCE (top matches from Codes sheet):\n{codes_text}\n"
        f"Use this to identify CPT code categories (e.g. E&M). "
        f"If a CPT is not covered by these results, use your own medical billing knowledge. "
        f"In your response, state whether each category was identified "
        f"'from sheet' or 'from LLM knowledge'."
    )

    messages = [
        {"role": "system", "content": _VALIDATE_SYSTEM},
        {"role": "user", "content": prompt},
    ]

    MAX_TOOL_ROUNDS = 3
    t_start = time.time()
    final_message = {}
    tools_called = []  # records every tool call this validation made

    for round_num in range(MAX_TOOL_ROUNDS):
        full_json = chat(
            OLLAMA_BASE, LLM_MODEL,
            messages=messages,
            tools=TOOL_DEFINITIONS,
            options={"temperature": 0.0, "top_p": 0.9},
            timeout=OLLAMA_TIMEOUT,
        )
        message = full_json.get("message", {}) or {}
        tool_calls = message.get("tool_calls") or []

        if not tool_calls:
            final_message = message
            break

        messages.append(message)
        print(f"[Validate] round {round_num + 1}: {len(tool_calls)} tool call(s)")
        for call in tool_calls:
            fn = call.get("function", {}) or {}
            name = fn.get("name")
            raw_args = fn.get("arguments", {})
            args = json.loads(raw_args) if isinstance(raw_args, str) else (raw_args or {})
            print(f"    -> {name}({args})")
            result = run_tool_call(name, args)
            tools_called.append({"name": name, "arguments": args, "result": result})
            messages.append({"role": "tool", "name": name, "content": json.dumps(result)})
    else:
        # Ran out of tool rounds -- force a final answer with no more tools offered.
        full_json = chat(
            OLLAMA_BASE, LLM_MODEL,
            messages=messages,
            options={"temperature": 0.0, "top_p": 0.9},
            timeout=OLLAMA_TIMEOUT,
        )
        final_message = full_json.get("message", {}) or {}

    elapsed_ms = int((time.time() - t_start) * 1000)
    raw = (final_message.get("content") or "").strip()

    json_match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not json_match:
        return {
            **_base,
            "tools_called": tools_called,
            "match_status": "INSUFFICIENT_DATA",
            "is_match": False,
            "reason": f"Validator did not return valid JSON. Raw: {raw[:1000]}",
            "conditions_unverifiable": ["Validator output was not valid JSON"],
        }, elapsed_ms

    try:
        validation = json.loads(json_match.group())
    except Exception as exc:
        validation = {
            **_base,
            "match_status": "INSUFFICIENT_DATA",
            "is_match": False,
            "reason": f"Failed to parse validator JSON: {exc}",
            "conditions_unverifiable": ["Validator JSON parsing failed"],
        }

    validation["validation_performed"] = True
    validation["tools_called"] = tools_called
    return validation, elapsed_ms

# -- MAIN ASK FUNCTION -------------------------------------------------
def ask(question: str) -> dict:
    try:
        extracted = extract_fields(question)
    except Exception as exc:
        return {"answer": f"Field extraction failed: {exc}", "extracted": None, "candidates": []}

    candidates      = retrieve_candidates(extracted)
    answer, elapsed = llm_match(extracted, candidates)
    code_candidates = retrieve_code_candidates_from_instruction(answer)
    validation, _   = validate_instruction_against_input(
        instruction_text=answer,
        question=question,
        code_candidates=code_candidates,
    )

    return {
        "answer":            answer,
        "extracted":         extracted,
        "candidates":        [e["id"] for e, _ in candidates],
        "response_time_ms":  elapsed,
        "validation":        validation,
        "cpt_source":        validation.get("cpt_source"),
        "cpt_source_detail": validation.get("cpt_source_detail"),
    }


# -- INTERACTIVE CLI -----------------------------------------------------
def run_cli() -> None:
    """Entry point for the interactive question loop. Called directly
    when this file is run as a script, and also called from the
    project-root main.py."""
    sample_questions = [
        "Get the comments for group 1 MICLAI practice with Medicare insurance and MED plan with CPT 99454 and denial code 151",
        "What should I do for a claim denied with code 29 and no proof of timely filing?",
        "J2356 denied with code 96 for AAASC practice Group 1 all insurance all plan",
        "CPT 95251 denied as inclusive by Medicare for MICLAI Group 1 MED plan with denial 97",
        "Lab code 83036 denied with denial 50 remark M25 for AZEN Group 2 all insurance",
        "get the comments for the denial code 29",
    ]

    print("\n" + "=" * 60)
    print("  Claims Denial KB - RAG_EXP  (LLM-first pipeline)")
    print("  Step 1 : LLM field extraction   (NL -> structured JSON)")
    print("  Step 2 : Structured BM25 query (JSON -> chunk -> BM25)")
    print("  Step 3 : LLM field matching     (structured JSON vs candidates)")
    print("  Type 'quit' to exit")
    print("=" * 60)
    print("\nSample questions (type 1-6):")
    for i, q in enumerate(sample_questions, 1):
        print(f"  {i}. {q}")

    while True:
        question = input("\nYour question: ").strip()
        if not question:
            continue
        if question.lower() in ("quit", "exit"):
            print("Goodbye!")
            break
        if question.isdigit() and 1 <= int(question) <= len(sample_questions):
            question = sample_questions[int(question) - 1]
            print(f"  -> {question}")

        try:
            result = ask(question)
        except (OllamaUnavailableError, OllamaModelNotFoundError) as exc:
            print(f"\n{exc}\n")
            continue

        print(f"\n{'-' * 60}")
        print(f"BM25 candidates : {result['candidates']}")
        print(f"Extracted fields : {json.dumps(result.get('extracted'), indent=2)}")
        print(f"Time             : {result.get('response_time_ms', 'N/A')} ms")
        print(f"\nInstruction:\n{result['answer']}")
        print("-" * 60)


if __name__ == "__main__":
    run_cli()
