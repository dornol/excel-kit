# excel-kit — Improvement Backlog

Analyzed 2026-04-19. Updated 2026-04-19.

---

## Declined

### R5. ColumnStyleConfig — 59 fields

Intentionally flat for fluent API. Splitting into FontConfig/BorderConfig would
force nested builders on callers (`cfg.font(f -> f.bold(true))`) without reducing
actual complexity. The R1 refactoring already solved the pain point (ExcelColumn
constructor breakage on new fields).

### A4. skipRowsWhere(predicate) — read-side row filtering

`readAsStream()` already returns a Stream, so `.filter()` covers this use case.
No need for a dedicated API.
