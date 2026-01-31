# PaperChecker Deep Code Check Worklist

## Highest priority (correctness + crash prevention)
- [ ] **Harden JSON pointer helpers**: guard list index parsing and bounds; return `None` or raise a controlled error for malformed pointers to avoid verifier crashes.
- [ ] **Normalize PMID comparisons**: coerce both Excel cell values and `pmid` to a common type (string or int) before row matching to prevent duplicate rows.
- [ ] **Add LLM call error handling**: wrap OpenAI/Gemini calls with retries/backoff and clear error reporting for schema or network failures.

## Medium priority (data quality + verification)
- [ ] **Decision dedupe strategy**: when multiple tasks write the same path, prefer the most recent decision (or merge) instead of first-write-wins.
- [ ] **Verifier context targeting**: optionally provide task-specific text windows per decision chunk to improve evidence precision.
- [ ] **Excel row emptiness detection**: consider formulas/formatting when detecting empty rows to avoid overwriting template content.

## Low priority (cleanup + maintainability)
- [ ] **Remove or wire unused helpers**: `_normalize_string`, `_values_match`, `_extract_page_from_evidence` are unused—either integrate or remove.
- [ ] **Add small unit tests**: cover JSON pointer handling, PMID matching, and decision dedupe behavior.

