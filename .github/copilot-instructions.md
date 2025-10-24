## Quick goal
Help maintainers extract structured PNR data from DOCX and render PNR DOCX files using LLMs (OpenRouter/Gemini). Be conservative: prefer cached outputs when editing logic, preserve schema and FERP canonical data.

## Where to look (high value files)
- `main10_openroute.py` — primary runner; contains pipeline (read DOCX → extract JSON via OpenRouter → compute FERP counts → generate sections → render DOCX).
- `main_grok_9_doc.py` — alternative/older runner showing similar flows and helpful comments.
- `pnr_extractor/schema.py` — authoritative JSON Schema (JSON_SCHEMA) the LLM must conform to. Use this when building structured-output prompts and validators.
- `pnr_extractor/ferp_reference.py` — single source of truth for FERP codes, titles, units, notes and canonical display order (FERP_ORDER).
- `pnr_extractor/extractor.py` — extraction, normalization, and assembly helpers (equipment detection, attach_ferp, build_pnr_from_ferp). Good examples of heuristics to follow.
- `pnr_extractor/*.txt` — prompt templates (PROMPT_*). They are read with pkg_resources and formatted with PROJECT_BLOCK, EQUIPMENT_BLOCK, COUNTS_BLOCK.

## Runtime environment and quick run
- Required env vars for normal runs:
  - `OPENROUTER_API_KEY` — API key used by OpenAI/OpenRouter client in code.
  - `OPENROUTER_MODEL` — (optional) model id used by OpenRouter calls; code sets a default.
  - `GEMINI_API_KEY` / `GEMINI_MODEL` — present in some files for Gemini compatibility; not mandatory unless you call Gemini-specific paths.

Example (run locally with cached data enabled):
```bash
export OPENROUTER_API_KEY="<your-key>"
# optionally: export OPENROUTER_MODEL="qwen/qwen3-235b-a22b:free"
python main10_openroute.py
```
Notes: the scripts read `DOCX_PATH_DEFAULT` and `DOCX_TEMPLATE_PATH_DEFAULT` constants at the top of the runner; modify those constants when running ad-hoc or edit `main()` to accept CLI args.

## Caching & debugging tips
- Cache location: `_pnr_cache/<safe_basename>/` adjacent to the source DOCX (safe basename created by `_safe_basename`).
  - `gemini.json` (or `gemini.json`-style files) store structured JSON returned by LLMs.
  - Per-section caches: `tasks.txt`, `scope.txt`, `procedure.txt`, `tests.txt`, `methodology.txt`.
- When LLM output fails schema validation, check the cached raw JSON in the above folder for diagnostics.
- Use `use_cache=True`/`force_refresh=False` defaults in generator helpers to avoid hitting rate limits while iterating.

## LLM & prompt patterns to preserve
- The pipeline strictly requests a single JSON object conforming to `JSON_SCHEMA`. Look at `call_structured_extractor` / `call_gemini_structured` for how prompts are composed:
  - system messages include the JSON schema (stringified) and a strict instruction: "Return ONLY a JSON object. No markdown/fences."
  - The code strips fences and `json:` prefixes (`_strip_md_fences`, `_safe` wrappers) before parsing.
- Retry/backoff behavior: MAX_LLM_RETRIES (default 5) and LLM_RETRY_DELAY_SEC (default 120s). Preserve these safeguards.

## Project conventions and invariants
- Canonical equipment types and allowed fields are defined in `schema.py`. Any change to extraction or to prompt must preserve schema compatibility.
- `pnr` fields: `main_tasks_5` must be exactly 5 items; `checklist_10_15` 10–15 items. These are relied on by rendering and JSON schema.
- `ferp_reference.py` is the single source of truth for FERP text/units/notes. Update it if you need to change labels or order; other code derives `FERP_TASKS`, `UNIT_BY_CODE`, `NOTE_BY_CODE` from it.
- Text prompts live in `pnr_extractor/*.txt`. They are filled using `.format(PROJECT_BLOCK=..., EQUIPMENT_BLOCK=..., COUNTS_BLOCK=...)` — the placeholders are important and must remain.

## Common edits an agent may be asked to do (how to do them safely)
- Add a new equipment_type: update `schema.py` enum, add mapping in `ferp_reference.EQUIPMENT_FERP_MAP`, and add heuristics in `extractor.py` (and tests/manual examples). Keep changes minimal and run a local extraction on a small DOCX.
- Change prompt wording: update the appropriate `prompt_*.txt` and preserve the `{PROJECT_BLOCK}`/`{EQUIPMENT_BLOCK}`/`{COUNTS_BLOCK}` placeholders and the escaping strategy used by `build_prompt()`.
- Change output DOCX templates: edit template file at `DOCX_TEMPLATE_PATH_DEFAULT`. The renderer expects placeholders like `{{tasks_block_full}}` etc.; keep those exact names or update `insert_text_with_style` calls accordingly.

## Fast checks before PR
- Run a small local extraction on one sample DOCX (set `DOCX_PATH_DEFAULT` to that file) and confirm:
  - `gemini.json` is created under `_pnr_cache/<safe>/` and parses as JSON.
  - `*_pnr.docx` renders and contains expected inserted text (open in Word/LibreOffice).
- Optional: `pip install jsonschema` to enable schema validation path and catch mismatches early.

## Edge cases to watch for
- Empty DOCX or DOCX with only tables → `read_docx_text` may return empty string; code logs a warning and skips.
- LLM returns fenced markdown or a prefixed `json:` — code strips that but if you change prompt style remove additional wrappers.
- Rate limits: generators have built-in sleeps and retries; avoid running non-cached heavy experiments on CI.

If anything here is unclear or you'd like more examples (sample DOCX, example gemini.json snippet to use in tests, or a CI job to validate schema conformity), tell me which part to expand and I will iterate.
