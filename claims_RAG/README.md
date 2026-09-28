# Claims RAG

A portable, local Retrieval-Augmented Generation (RAG) application for
matching medical billing claim denial questions against a rules-based
Knowledge Base, using a locally-run LLM (Ollama + `qwen3.5:4b`) and
BM25 keyword retrieval.

## 1. What This Project Does

You type a natural-language question about a claim denial (for
example: *"CPT 95251 denied as inclusive by Medicare for MICLAI
Group 1 MED plan with denial 97"*), and the application:

1. Uses the LLM to extract structured billing fields from your
   question (group, practice, insurance company, plan, CPT codes,
   denial code, remark code).
2. Uses BM25 keyword search to retrieve the most relevant rules from
   your Knowledge Base.
3. Uses the LLM again to strictly match the extracted fields against
   the retrieved rule candidates and pick the correct one.
4. Uses the LLM a third time to validate that matched rule against
   your original question, optionally looking up denial codes, CPT
   codes, or allowable amounts as reference data.
5. Prints the final instruction.

## 2. Architecture

```
User question (typed into the terminal)
        |
        v
LLM field extraction        (extract_fields)
        |
        v
Structured fields (JSON)
        |
        v
BM25 candidate retrieval     (retrieve_candidates - rank_bm25.BM25Okapi)
        |
        v
LLM field matching           (llm_match)
        |
        v
Code/denial candidate retrieval  (retrieve_code_candidates_from_instruction)
        |
        v
Validation (with optional tool calls)  (validate_instruction_against_input)
        |
        v
Final result                 (ask)
```

Every LLM call goes through Ollama's local HTTP API
(`http://localhost:11434`) to the `qwen3.5:4b` model - there is no
cloud API involved, and you never need to type questions directly
into an `ollama run` chat session.

## 3. Folder Structure

```
Claims_RAG/
├── app/
│   ├── paths.py              Dynamic project-path resolution
│   ├── file_discovery.py     Automatic KB file discovery
│   ├── config.py             Loads config/config.json
│   ├── ollama_manager.py     All Ollama HTTP API communication
│   ├── rag_claims_kb.py      The RAG pipeline itself
│   └── validation_tools.py   Tool functions used during validation
│                              (lookup_denial_code, lookup_allowable,
│                              lookup_cpt_code) - see note below
├── data/
│   ├── input/                 reserved for future use, not yet read
│   ├── kb/                    put your Knowledge Base file here
│   └── output/                 reserved for future use, not yet written
├── config/
│   └── config.json            Ollama URL/model, BM25 top_k
├── indexes/                    reserved for future use, not currently written to
├── logs/                       reserved for future use, not currently written to
├── main.py                     entry point - run this
├── requirements.txt
├── run.bat
└── README.md
```

> **Note on `validation_tools.py` - read this before relying on
> validation results:** the file included in this package is a
> **placeholder, not the real implementation**. The original Claims
> Denial RAG sample's `lookup_denial_code()`, `lookup_allowable()`,
> and `lookup_cpt_code()` functions require real denial-code
> definitions, real CPT/E&M categorization, and real allowable-amount
> reference data that were never supplied anywhere in this project's
> source material. Rather than invent plausible-looking versions of
> that business logic (which would silently produce wrong validation
> results that *look* correct), the included placeholder sets
> `TOOL_DEFINITIONS = []`. This lets the full pipeline - extraction,
> BM25 retrieval, matching, and validation - run end-to-end using only
> the CPT reference data already passed inline from the `Codes` sheet
> plus the model's own general knowledge, but it means the LLM is
> **never able to call a tool for an additional lookup** during
> validation. **To enable real tool-based validation lookups, replace
> `app\validation_tools.py` with your original project's actual
> implementation** of the three functions above and a matching
> `TOOL_DEFINITIONS` schema. Until then, treat the validation step's
> output as based only on inline reference data and the model's own
> knowledge, not on verified lookups.

> **Note on `indexes/` and `logs/`:** these folders are created
> automatically (see `paths.py`) and reserved for future use, but
> nothing in the current code writes to them yet. They exist so the
> project structure is ready for that without further path changes.

## 4. Requirements

- **Python**: 3.10 or later (uses modern type-hint syntax such as
  `list[str]` and `X | None`)
- **Ollama**: version 0.14.2 or compatible, installed and runnable
  locally
- **Model**: `qwen3.5:4b` pulled into Ollama

## 5. Installing Python Dependencies

```powershell
cd Claims_RAG
python -m pip install -r requirements.txt
```

`requirements.txt` contains only the packages the code actually
imports: `pandas`, `numpy`, `rank_bm25`, `requests`, and `openpyxl`
(required by pandas for Excel) are always needed. `python-docx`,
`pdfplumber`, and `beautifulsoup4` are only exercised if you place a
`.docx`, `.pdf`, or `.html`/`.htm` Knowledge Base file, respectively,
in `data\kb\` instead of an `.xlsx` file.

`validation_tools.py`'s own dependencies (if any) are not covered
here, since its contents are not part of this documentation set -
check its imports yourself and add anything missing.

## 6. Ollama Setup

Install Ollama separately (not part of this Python project) from
[ollama.com](https://ollama.com). Then, in a terminal:

```powershell
ollama serve
```

Leave that window running. In another terminal, pull the model if you
haven't already:

```powershell
ollama pull qwen3.5:4b
```

The application checks both of these automatically at startup - see
sections 13 and 14.

## 7. Where to Place the Knowledge Base

Put your Knowledge Base Excel workbook in:

```
Claims_RAG\data\kb\
```

The workbook's **first sheet** is read as the rules table and must
contain columns named exactly `Rule ID`, `Description`, and
`Validation`. A second sheet named exactly `Codes` must exist, with
columns `Category` and `Code`.

## 8. Automatic Knowledge Base Discovery

You never type a KB filename anywhere. At startup, `file_discovery.py`
scans `data\kb\`:

- **Exactly one** supported file found -> used automatically.
- **Zero** files found -> the application exits with a message naming
  the exact folder to place a file in.
- **More than one** file found -> the application exits, listing every
  candidate by name, and asks you to keep only one.

## 9. Starting the Application

**Option A - double-click `run.bat`** (Windows Explorer), or run it
from PowerShell:

```powershell
.\run.bat
```

**Option B - run directly with Python:**

```powershell
cd Claims_RAG
python main.py
```

Both start the identical application; `run.bat` just adds a Python
availability check and keeps the window open afterward so you can
read the output.

## 10. How `run.bat` Works

It reads its own location using `%~dp0` (a batch-file variable that
always resolves to the drive and folder the `.bat` file itself is
running from), changes into that directory, checks that `python` is
on your PATH, and then runs `main.py` from that same location. It
never contains a written-out drive letter or folder name, which is
what makes it work unchanged after copying the project elsewhere.

## 11. What Happens If Ollama Is Unavailable

The application checks Ollama's availability before loading anything
else. If it can't connect, it prints an explanation and the exact
command to fix it (`ollama serve`), then exits cleanly - no Python
traceback.

## 12. What Happens If `qwen3.5:4b` Is Missing

If Ollama is running but the model isn't installed, the application
exits with the exact command to fix it (`ollama pull qwen3.5:4b`) and
lists whatever models *are* currently installed.

## 13. What Happens If Multiple KB Files Exist

The application refuses to guess. It lists every supported file found
in `data\kb\` and asks you to remove all but one.

## 14. Supported Knowledge Base File Formats

| Extension | Requires |
|---|---|
| `.xlsx`, `.xls` | (built in, via pandas/openpyxl) |
| `.docx` | `python-docx` |
| `.pdf` | `pdfplumber` |
| `.html`, `.htm` | `beautifulsoup4` |
| `.md` | (built in) |

`.txt` is intentionally **not** supported - there is no reader for it
in the code, so it is not advertised as a supported format.

## 15. How Portability Works

Every path in the project is derived, at runtime, from
`Path(__file__).resolve()` inside `app\paths.py` - never written out
as a fixed string. Copying the entire `Claims_RAG` folder from one
drive or folder to another (e.g. `C:\Claims_RAG` to `D:\Claims_RAG`)
requires no source code changes; every module recalculates its
location and all dependent paths (`data\`, `config\`, `indexes\`,
`logs\`) from wherever the project currently sits.

## 16. Example Questions

These match the style of question the field-extraction step expects
(they are the built-in sample questions the CLI offers as menu
options 1-6):

- *"Get the comments for group 1 MICLAI practice with Medicare
  insurance and MED plan with CPT 99454 and denial code 151"*
- *"What should I do for a claim denied with code 29 and no proof of
  timely filing?"*
- *"get the comments for the denial code 29"*

The actual result for any question depends entirely on the rules
present in **your** Knowledge Base file - this application does not
ship with pre-defined answers.

## 17. BM25 Retrieval Architecture

Retrieval uses `rank_bm25.BM25Okapi` exclusively - a keyword/term-
frequency ranking algorithm, not a vector or embedding-based search.
Two separate BM25 indexes are built at startup: one over the KB rules
(`description` + `validation` text per rule), and one over the
`Codes` sheet (`category` + `code` text per entry). No vector
database, embedding model, or semantic search library (FAISS, Chroma,
Pinecone, etc.) is used anywhere in this project.

## 18. Validation Tools

See the note in section 3 above. In short: `app\validation_tools.py`
as shipped is a placeholder with `TOOL_DEFINITIONS = []` - real
denial-code, CPT, and allowable-amount lookup logic must be supplied
by you from the original project before validation results reflect
anything beyond inline reference data and the model's own knowledge.

## 19. Testing Instructions

See the Testing Checklist provided alongside this README for the
full list of tests and their status (verified in a Linux sandbox
using mocked Ollama responses, vs. tests that require your actual
Windows + Ollama + qwen3.5:4b environment and are marked
**NOT EXECUTED - USER MUST RUN**). At minimum, before relying on this
application:

1. Run `python app\ollama_manager.py` to confirm Ollama and
   `qwen3.5:4b` are both detected on your machine.
2. Place your real KB in `data\kb\` and run `python app\file_discovery.py`
   to confirm it's found.
3. Replace `app\validation_tools.py` with your real implementation.
4. Run `python main.py` and try at least 3 questions based on your
   actual KB's rules.

## 20. Known Limitations

- `validation_tools.py` ships as a placeholder only (see sections 3
  and 18) - its real business logic was never available in this
  project's source material and has not been fabricated.
- `.xls` (the legacy pre-2007 Excel format) is listed in the loader
  registry but has not been exercised in testing; older `.xls` files
  may require an additional `xlrd` package that is not currently in
  `requirements.txt`. `.xlsx` is the tested and recommended format.
- The first sheet of the KB workbook is read as the rules table
  regardless of its name - it is not required to be literally named
  "Rules". Only the `Codes` sheet is matched by exact name.
- `data\input\` and `data\output\`, and the `indexes\` and `logs\`
  folders, are created automatically but nothing in the current code
  reads from or writes to them yet.
- LLM output quality (correct field extraction, correct rule
  matching, correct validation) depends on `qwen3.5:4b`'s actual
  responses on your machine and cannot be guaranteed by the
  surrounding Python code - the code's job is to reliably route
  requests and handle malformed responses, not to guarantee any
  specific answer is correct.
