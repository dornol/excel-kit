---
template: analysis
feature: map-reader-absorption
project: excel-kit
project-version: 0.12.0
date: 2026-04-12
status: Final
---

# map-reader-absorption Gap Analysis

**Plan**: [../01-plan/features/map-reader-absorption.plan.md](../01-plan/features/map-reader-absorption.plan.md)
**Design**: [../02-design/features/map-reader-absorption.design.md](../02-design/features/map-reader-absorption.design.md)
**Implementation commits**:
- `ee4a3c3` Absorb Map readers into Reader.forMap() factories for v0.12.0
- `6abd8bc` Strengthen map-mode tests with edge cases and behavioral pins
- *(post-analysis)* `HEAD` Close map-reader-absorption G1: broken javadoc link in ExcelReadSupport

---

## Overall Match Rate: **100%** (after G1 close)

Pre-close: 98%. Single medium-severity gap was a broken `@link ExcelMapReader` in a library-internal javadoc comment (`ExcelReadSupport.java:4`) that would have tripped the Plan's "javadoc 경고 0" NFR. Closed by dropping the stale reference from the `@link` list.

---

## Task Status

| Task | Status | Notes |
|------|:------:|-------|
| **M1** `ExcelReader.forMap()` | ✅ | `ExcelReader.java:58` flag, `:167-187` factory |
| **M2** `CsvReader.forMap()` | ✅ | `CsvReader.java:43` flag, `:122-140` factory |
| **M3/M4** Handler changes | ✅ *(no-op)* | Zero changes to `ExcelReadHandler` / `CsvReadHandler` — Design §2.1 predicted this correctly |
| **M5** File deletion | ✅ | `ExcelMapReader.java`, `CsvMapReader.java` removed |
| **M6** Test migration | ✅ | 0 `new ExcelMapReader()` / `new CsvMapReader()` in tests |
| **M7** Mixed-mode guards | ✅ | 5 guards × 2 readers = 10 call sites |
| **M8** New tests | ✅ | `ExcelReaderMapModeTest` (35 annotations), `CsvReaderMapModeTest` (33 annotations) |
| **M9** Documentation | ✅ | build.gradle.kts `0.12.0`, CHANGELOG `[0.12.0]`, README/META-INF/llms.txt migrated |
| **M10** Example migration | ✅ | `ReadShowcaseController`, `CsvShowcaseController` use `forMap()` |

All 10 tasks fully done. Nothing partial.

---

## Gaps

### Closed during analysis

| # | Severity | Location | Fix |
|---|:--------:|----------|-----|
| **G1** | Medium | `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelReadSupport.java:4` | Class-level javadoc had `{@link ExcelMapReader}` pointing at the deleted class. Tripped the `./gradlew javadoc` zero-warning NFR. Dropped `ExcelMapReader` from the `@link` list; now references only the surviving `ExcelReadHandler` and `ExcelReader`. |

### Out-of-scope finding (fixed opportunistically)

| # | Location | Description |
|---|---|---|
| **O1** | `kit/src/main/resources/META-INF/excel-kit/csv.md:98` | `CsvMapWriter writer = CsvWriter.forMap(...)` — leftover from v0.11.0 documentation migration where the type annotation wasn't updated when `CsvMapWriter` was deleted. Not a gap for this feature, but fixed while reviewing: `CsvWriter<Map<String, Object>> writer = ...`. |

### Cosmetic, not counted as gaps

| # | Location | Description |
|---|---|---|
| C1 | `CoverageBoostTest.java` class javadoc + nested class names (`CsvMapReaderStreamErrors`, `ExcelMapReaderStreamCoverage`, `CsvMapReaderConfigCoverage`) | Test-file-only references to deleted class names. Tests themselves use `forMap()`. Test javadoc is not emitted as library javadoc, so no NFR impact. Worth renaming in a future cosmetic cleanup pass. |
| C2 | `MapReaderStreamTest.java:16` javadoc | `{@link ExcelMapReader.ExcelMapReadHandler#readAsStream()}` — dangling link in a test file. Same rationale as C1. |

---

## Deviations from Plan (Design-approved)

All six deviations are documented and accepted.

| # | Plan said | Reality |
|---|---|---|
| **D1** | M3/M4: extend `ExcelReadHandler` / `CsvReadHandler` with a 3rd mode, graft `MapSheetHandler` as an inner class | Zero handler changes. Design §2.1 discovered that `RowData` already exposes `headerNames()` + `get(name)`, so the absorption can be a synthetic `Function<RowData, Map<String, String>>` that reuses the existing mapping-mode path. |
| **D2** | Plan §5 rated "SAX state-machine rewrite" as **High/Medium** risk | Risk eliminated. Nothing was rewritten. |
| **D3** | Design §3.5 initially suggested "named lookup via `RowData.get(headerName)` is more robust" | Implementation uses **positional pairing** truncated at `min(headers.size(), row.size())` via `row.get(i)`, to match the deleted `ExcelMapReader` / `CsvMapReader` behavior bit-for-bit. Pinned by `BehavioralEquivalence` tests in M8. |
| **D4** | Plan §2.1 M7 listed 3 methods to guard | Implementation guards 5 methods (`column×2`, `columnAt`, `skipColumn`, `skipColumns`). Design §3.3 already listed all 5. Plan undercounted. |
| **D5** | Plan didn't mention it | `ExcelReader.forMap()` gains `onProgress` support that `ExcelMapReader` never had. Celebrated in CHANGELOG as an incidental improvement. |
| **D6** | Plan didn't mention it | `readAsStream()` on a non-existent sheet now throws `ExcelReadException` (was silent empty stream). More correct — surfaces caller bugs instead of hiding them. Documented in CHANGELOG. |

---

## Test Coverage Delta

| Test file | Test/Display methods | Status |
|---|:---:|:---:|
| `ExcelReaderMapModeTest.java` | **35** | New — Factory (2) + HeaderAutoDiscover (2) + FluentApi (4) + MixedModeGuards (7) + BehavioralEquivalence (4) + ReadPathErrors (1) + MultiSheetSelection (1) + nested DisplayNames (14) |
| `CsvReaderMapModeTest.java` | **33** | New — Factory (2) + HeaderAutoDiscover (2) + FluentApi (6) + MixedModeGuards (7) + BehavioralEquivalence (5) + nested DisplayNames (11) |

Key behavior pins added:
- Excel **fewer cells than headers** → trailing keys absent
- Excel **more cells than headers** → extras ignored
- Excel **blank header cell** → becomes `""` map key (pins `CellData` null→"" coercion)
- Excel **present-but-empty cell** → `""` value (never null)
- Excel **multi-sheet `sheetIndex(1)`** → actually selects sheet 2, no leakage
- Excel **`read()` on non-existent sheet** → throws `ExcelReadException` (symmetric with `readAsStream()`)
- CSV **more cells than headers**
- CSV **blank header cell** → `""` key (documented against Excel's equivalent path)
- CSV **present-but-empty cell**
- CSV **duplicate headers** → `putIfAbsent` + second-put wins
- CSV **BOM prefix** → stripped from first header
- CSV **guard message format** → names method + mentions `forMap()`

---

## Recommendations

1. **Ready for `/pdca report map-reader-absorption`** — Match Rate 100% after G1 close. All Plan tasks done, all Deviations pre-approved, all critical behaviors pinned by tests.
2. Cosmetic cleanup (C1, C2) can happen any time — not release-blocking.
3. Then `v0.12.0` release checklist per `CLAUDE.md`.

---

## Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 1.0 | 2026-04-12 | Initial gap analysis — 98% pre-close, 100% after G1 fix. 6 pre-approved Deviations. 68 new M8 test methods across Excel and CSV sides. | gap-detector |
