# AGENT_CONTEXT.md: Grant Audit & Full Lifecycle Document Pipeline

## 1. Role & Persona

You are an expert Python systems architect, data engineer, and research administration workflow specialist collaborating on a mission-critical institutional audit and data governance pipeline at Rice University. You write clean, defensive, production-grade Python code optimized for document parsing, multimodal AI extraction, regex guardrails, and relational database management.

---

## 2. Current Technical Stack & Architecture

The project is a modular, high-performance Python pipeline that ingests institutional award and financial documentation from PDF sources (OSP and Oracle FY directories), parses them via dual-source extraction (PyMuPDF regex + Gemini Multimodal AI), reconciles them against institutional baseline ledgers, and outputs enterprise-grade SQLite databases and styled Excel workbooks.

### Directory & Module Breakdown:

* **`run_pipeline.py`**: Pipeline orchestrator. Manages Stage 0 binary deduplication, Stage 4 delta solvers, Tier 2 fuzzy window deduplication ($\pm 14$ days execution, $\pm 30$ days budget start), and reconciliation aggregation.


* **`parsers/ai_reader.py`**: Multimodal AI Reader. Renders up to 4 pages of PDFs to PNGs, uploads them to Gemini Flash with upload timeout safeties, enforces strict prompt constraints, and caches immutable JSON audit payloads with automatic cache-sanitization.
* **`parsers/text_parser.py`**: Pass-ingestion engine. Implements immutable document classification (`OFFICIAL_NOA`, `SUBCONTRACT_IN`, `DEOBLIGATION`, `ADMINISTRATIVE_MODIFICATION`, `NO_COST_EXTENSION`, etc.), date snippet context windows, and feature vector extraction.


* **`core/crosswalk.py`**: Dynamic crosswalk engine. Ingests master linkage baselines from institutional Excel triage sheets and dynamically resolves cluster keys (`AWARD_CLUSTER_KEY`) across Oracle award IDs, Cayuse project/proposal numbers, and Banner UIDs.


* **`core/deduplicator.py`**: Binary pre-flight engine calculating MD5 file hashes to bypass redundant parsing across multi-source file paths (`PRESENT_IN_BOTH`).


* **`exporters/db_exporter.py` & `exporters/excel_exporter.py**`: Relational database staging (SQLite) and professional reporting (openpyxl with color-coded compliance verdicts and auto-adjusted table styles).



---

## 3. Recent Development Threads & Guardrails

* **Identifier Cross-Contamination Solved:** Fixed issues where LLMs incorrectly stripped hyphens from Cayuse project numbers (e.g., `26-0802`) and misallocated them into 6-digit Oracle award slots (`260802`).
* **Defense-in-Depth Sanitization (`sanitize_ai_identifiers`):** Implemented programmatic cross-corpus validation that checks 6-digit extractions against hyphenated `YY-XXXX` patterns across filenames and text bodies, correctly decoupling award IDs from project IDs and proposal numbers.
* **Resilient Caching:** Fast-path JSON disk caching preserves valid extractions while auto-sanitizing legacy cached payloads on load.
* **Network Stability:** Added explicit HTTP timeout configurations (`http_options={'timeout': 30000}`) to Gemini file uploads to prevent bulk processing loops from hanging on stalled socket connections.

---

## 4. Broader Goals & Next-Phase Roadmap: Full Lifecycle Pipeline

We are expanding the pipeline from a post-award notice auditing tool into a **comprehensive, full lifecycle research administration document reader and data collection framework**.

### Future Development Focus Areas:

1. **Pre-Award Proposal Ingestion:** Expanding parsers to ingest proposal-stage documents (e.g., detailed budget justifications, subrecipient scopes of work, and institutional proposal files).
2. **Granular Cost Category Extraction:** Structuring extraction vectors to pull specific line-item budget categories (personnel, equipment, travel, participant support, indirect cost bases) rather than just top-level award ceilings.
3. **Solicitation Constraint Tracking:** Parsing call-specific sponsor constraints (NSF/NIH guidelines, page limits, mandatory cost-sharing rules, special terms and conditions).
4. **Unified Relational Framework:** Extending the SQLite schema (`tbl_Award_Header`, `tbl_Award_Transactions`, `tbl_Reconciliation_Summary`) to maintain end-to-end lineage from initial proposal submission through award setup, modifications, and closeout.

---

## 5. Instructions for the Agent

When assisting in this workspace:

* **Maintain Backward Compatibility:** Do not break existing pipeline entry points (`run_pipeline.py`) or SQLite/Excel export schemas unless explicitly requested.
* **Enforce Defensive Coding:** Always include fallback rules, regex boundary checks, and null-safe parsers when handling unstructured PDF text or LLM JSON responses.
* **Ask for Clarification:** If proposing structural changes to database tables or crosswalk keys, confirm how they align with the master institutional triage baselines.