---
template: analysis
feature: api-cleanup-v011
project: excel-kit
project-version: 0.11.0
date: 2026-04-12
status: Final
---

# api-cleanup-v011 Gap Analysis

**Plan**: [../01-plan/features/api-cleanup-v011.plan.md](../01-plan/features/api-cleanup-v011.plan.md)
**Design**: [../02-design/features/api-cleanup-v011.design.md](../02-design/features/api-cleanup-v011.design.md)
**Implementation**: commits `a32620a` → `feb1e58` (5 commits on `main`)
**Scope**: T1, T2, T4, T5, T7

---

## Overall Match Rate: **99%**

The implementation matches the design tightly. Every in-scope deliverable landed. Three deviations from the design are all pre-approved or accepted. Only one **Low** severity polish item remains (CHANGELOG missing a note about the FileHandler sealed→interface deviation).

---

## Task Status Summary

| Task | Name | Status | Notes |
|------|------|:------:|-------|
| **T1** | ExcelWriter Builder | ✅ Done | 5 public constructors removed; `builder()` + `Builder<T>` in place with validation |
| **T2** | FileHandler + `write()` rename | ✅ Done | Interface created; both handlers `final` and implement it; no stale `consumeOutputStream(` raw method |
| **T4** | Reader column unify | ✅ Done (with deviation D2) | Legacy methods removed; `column`/`columnAt`/`skipColumn(s)` return Reader; configurer overload intentionally dropped |
| **T5** | Map Writer delete + `forMap()` | ✅ Done | Map Writer files deleted; 3 `forMap` factories present; Map Readers preserved for v0.12.0 |
| **T7** | Documentation | ✅ Done | Version bump, CHANGELOG migration guide, README clean, META-INF + llms.txt updated |

All five tasks are fully done. Nothing partial.

---

## Per-Task Verification

### T1 — ExcelWriter Builder

- `ExcelWriter.java:30` — class has **no** public constructors; only package-private `ExcelWriter(Builder<T> builder)` at `:148`
- `ExcelWriter.java:72` — `public static <T> Builder<T> builder()` exists
- `ExcelWriter.java:167-231` — `public static final class Builder<T>` with:
  - `color(ExcelColor)` at `:184` — null check (`IllegalArgumentException`)
  - `maxRows(int)` at `:199` — validates `<= 0`
  - `rowAccessWindowSize(int)` at `:215` — validates `<= 0`
  - `build()` at `:228` — returns `new ExcelWriter<>(this)`
- Defaults via named constants (`DEFAULT_MAX_ROWS = 1_000_000`, `DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000`)
- Javadoc example matches Design §3.1

### T2 — FileHandler interface + write() rename

- `shared/FileHandler.java:29` — `public interface FileHandler` (regular interface — see deviation D1)
  - `write(OutputStream out) throws IOException` at `:40`
- `ExcelHandler.java:43` — `public final class ExcelHandler implements FileHandler`
  - `@Override public void write(OutputStream outputStream) throws IOException` at `:84-91`
  - Both `consumeOutputStreamWithPassword` overloads preserved (Excel-only, as specified)
- `CsvHandler.java:22` — `public final class CsvHandler extends TempResourceContainer implements FileHandler`
  - `@Override public void write(OutputStream outputStream)` at `:50-51` — narrower `throws` clause (see D3)
  - Javadoc documents that it never actually throws IOException
- Grep confirms zero raw `consumeOutputStream(` callsites anywhere (all matches are `consumeOutputStreamWithPassword`)

### T4 — Reader column unify

- `ExcelReader.java`:
  - No public `addColumn`, no `columnAtBuilder`
  - `column(BiConsumer)` at `:182` returns `ExcelReader<T>`
  - `column(String, BiConsumer)` at `:195` returns `ExcelReader<T>`
  - `columnAt(int, BiConsumer)` at `:208` returns `ExcelReader<T>`
  - `skipColumn()` at `:219`, `skipColumns(int)` at `:231` preserved
- `CsvReader.java` — symmetric to Excel
- `ExcelReadColumn.ExcelReadColumnBuilder` inner class **deleted**
- `CsvReadColumn.CsvReadColumnBuilder` inner class **deleted**
- See deviation D2 about configurer overloads

### T5 — Map Writer delete + forMap()

- `ExcelMapWriter.java` and `CsvMapWriter.java` — **deleted**
- `ExcelWriter.forMap(String...)` at `ExcelWriter.java:98`
- `ExcelWriter.forMap(String[], Consumer<...>...)` at `ExcelWriter.java:129-146` with length validation
- `CsvWriter.forMap(String...)` at `CsvWriter.java:70`
- `ExcelMapReader` / `CsvMapReader` preserved — correct for v0.11.0 scope per Design §4.2

### T7 — Documentation

- `build.gradle.kts:7` — `version = "0.11.0"` ✓
- `CHANGELOG.md:5` — `[0.11.0]` entry with full Breaking Changed section + Migration Guide
- `README.md` Installation snippets show `0.11.0`
- `README.md` has zero legacy API references
- `META-INF/AI.md`, `META-INF/excel-kit/*.md`, `docs/llms.txt` updated
- `example/` fully migrated

---

## Gaps

| # | Severity | Location | Description |
|---|:--------:|----------|-------------|
| G1 | **Low** | `CHANGELOG.md` FileHandler bullet | Doesn't mention that `FileHandler` is a plain `interface` rather than `sealed`. Design §2.3 explicitly called out this deviation; Javadoc on `FileHandler.java:20-25` documents it, but not the release notes. Optional polish — add one line. |

That's it. One Low polish item.

---

## Deviations (pre-approved, not counted as gaps)

| # | Location | Deviation | Justification |
|---|----------|-----------|---------------|
| **D1** | `shared/FileHandler.java:29` | `public interface FileHandler` — not `sealed` | Design §2.3 pre-approved this: excel-kit ships as an automatic module (no `module-info.java`), which breaks `sealed`'s same-module requirement. Handlers are `final` + package-private constructors, so the closed-hierarchy intent is preserved. Javadoc documents this. |
| **D2** | `ExcelReader` / `CsvReader` column methods | Only 3 column methods each (no `Consumer<Builder>` overloads) instead of 8 specified in Plan FR-05 / Design §3.4 | During implementation, `ExcelReadColumnBuilder` / `CsvReadColumnBuilder` were found to have no actual configuration methods worth exposing — they were just chain continuations. So the configurer overload would be a no-op surface expansion. The inner Builder classes are deleted, so the configurer overload has nothing to configure anyway. |
| **D3** | `CsvHandler.write(OutputStream)` | Does not declare `throws IOException` (narrowed throws on override) | Legal Java. CsvHandler wraps IOException as `CsvWriteException` internally, so declaring `throws IOException` would force callers into try/catch for an exception that never actually escapes. Design §3.3 option matched. |

---

## Recommendations

1. **(Low — optional)** Add one line to `CHANGELOG.md` under the `FileHandler` bullet: "Note: `FileHandler` is a plain interface rather than `sealed` because excel-kit ships as an automatic module; third-party implementations remain unsupported." Closes G1.
2. Proceed to `/pdca report api-cleanup-v011` — Match Rate ≥ 90%, Definition of Done satisfied.
3. Then the v0.11.0 release checklist in `CLAUDE.md`: tag, push, Maven Central publish.

---

## Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 1.0 | 2026-04-12 | Initial gap analysis — 99% match, 1 Low gap, 3 accepted deviations | gap-detector |
