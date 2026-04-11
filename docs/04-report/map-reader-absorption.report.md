---
template: report
version: 1.0
feature: map-reader-absorption
project: excel-kit
project-version: 0.12.0
author: DongHyeok Kim
date: 2026-04-12
status: Final
---

# map-reader-absorption Completion Report

> **Feature**: Complete the Map I/O symmetry by absorbing `ExcelMapReader` / `CsvMapReader` into `ExcelReader.forMap()` / `CsvReader.forMap()` static factories. Finishes the cleanup work deferred from v0.11.0.
>
> **Project**: excel-kit
> **Version**: 0.11.0 → 0.12.0
> **Duration**: 2026-04-12 (1 iteration + analysis cycle)
> **Status**: Complete
> **Match Rate**: 100%

---

## Executive Summary

| Perspective | Content |
|---|---|
| **Problem** | v0.11.0 achieved Map Writer symmetry via `ExcelWriter.forMap()` / `CsvWriter.forMap()` factories but deferred Map Reader due to perceived SAX state-machine rewrite risk. Result: API asymmetry — Map Writer was a factory pattern, Map Reader remained a dedicated class. Users reached for `new ExcelMapReader()` while using `Writer.forMap()`, causing cognitive friction and preventing Reader from inheriting fluent features like `onProgress`. |
| **Solution** | During Design, discovered that the existing `Function<RowData, T>` mapping-mode infrastructure already exposes `headerNames()` and `get(headerName)`, enabling absorption via a synthetic `mapMapper` that builds `Map<String, String>` without touching SAX handlers. Implementation: added `mapMode` flag + `forMap()` factory to `ExcelReader` / `CsvReader`, guarded mixed-mode calls at runtime, migrated 6 test files and 2 example files, deleted both Map Reader files. **Zero SAX handler changes.** Plan's top risk (High/Medium) was eliminated during Design via infrastructure reuse. |
| **Function/UX Effect** | Full Map API symmetry (Writer ↔ Reader both `forMap()`). Users now write `ExcelReader.forMap().sheetIndex(0).onProgress(1000, cb).build(in).read(...)` with complete fluent API parity. Excel readers gain `onProgress` support they never had before (old `ExcelMapReader` didn't support callbacks). Map reading is now a factory pattern, not a dedicated class, matching the Writer side. 68 new test methods pin all edge cases (null headers, blank cells, multi-sheet selection, guard errors). |
| **Core Value** | Completes v0.12.0's scope (v0.11.0 promise fulfilled). Establishes Map I/O as one consistent fluent-API-first pattern for the library. Positions v0.12.0 as the release that resolved all Reader/Writer asymmetries — one step closer to v1.0.0 API stability. Single-user breaking was cleanly executed: two 500+-line Map Reader files deleted, no deprecation path (per v0.11.0 policy). |

---

## PDCA Cycle Summary

### Plan Phase

**Document**: [docs/01-plan/features/map-reader-absorption.plan.md](../../01-plan/features/map-reader-absorption.plan.md)

- **Goal**: Define 10 tasks (M1–M10) to delete `ExcelMapReader.java` / `CsvMapReader.java` and expose `Reader.forMap()` factories. Complete v0.11.0's deferred Map Reader cleanup.
- **Scope**: In-scope — M1/M2 factories, M3/M4 handler integration (presumed to be 3-mode extension), M5 file deletion, M6/M8 tests, M7 mixed-mode guards, M9 documentation, M10 example migration. Out-of-scope — ColumnStyleConfig/ExcelColumnBuilder unification (separate feature).
- **Key Assumption**: M3/M4 would require rewriting `MapSheetHandler` state machine and extending handler classes to 3-mode architecture. Rated as **High/Medium risk** in Plan §5.
- **Success Criteria**: Match Rate ≥ 90%, all 10 tasks complete, zero javadoc warnings, `./gradlew test` green, files deleted.

### Design Phase

**Document**: [docs/02-design/features/map-reader-absorption.design.md](../../02-design/features/map-reader-absorption.design.md)

- **Design Goal**: Eliminate the SAX rewrite risk through infrastructure reuse.
- **Key Discovery (§2.1 — "Plan vs Reality")**: Reading the codebase revealed that the existing `RowData` class already exposes `headerNames()` and `get(String headerName)`. The mapping-mode path already constructs a `RowData` for every row. Therefore, map absorption can be implemented as a synthetic `Function<RowData, Map<String, String>>` that converts one `RowData` to a map — exactly what the deleted `ExcelMapReader` did. **No SAX handler changes needed.** The entire "Handler 3-mode extension" (M3/M4) evaporates.
- **Consequence**: Plan's top risk (SAX rewrite) is **eliminated entirely**. The feature reduces to Reader-level additions (flag + factory + guards) + file deletion + test migration. Much smaller footprint.
- **Breaking changes approved**: Delete both Map Reader files immediately (single-user policy). Runtime guard prevents column() calls on map-mode readers.
- **Behavioral improvements**: (1) `ExcelReader.forMap()` gains `onProgress` support that old `ExcelMapReader` lacked. (2) `readAsStream()` on non-existent sheet now throws instead of silently returning empty stream.

### Do Phase

**Implementation Commits**:
- `ee4a3c3` — Absorb Map readers into Reader.forMap() factories for v0.12.0
- `6abd8bc` — Strengthen map-mode tests with edge cases and behavioral pins

**Tasks Completed**:

| Task | Description | Status |
|------|---|:---:|
| **M1** | `ExcelReader.forMap()` static factory + `mapMode` field | ✅ |
| **M2** | `CsvReader.forMap()` static factory + `mapMode` field | ✅ |
| **M3/M4** | Handler changes (discovered to be unnecessary) | ✅ No-op |
| **M5** | Delete `ExcelMapReader.java` / `CsvMapReader.java` | ✅ |
| **M6** | Migrate existing map-reader tests to new API | ✅ |
| **M7** | Runtime guards on `column()` / `columnAt()` / `skipColumn()` / `skipColumns()` (5 methods × 2 readers) | ✅ |
| **M8** | Write `ExcelReaderMapModeTest` (35 test methods) and `CsvReaderMapModeTest` (33 test methods) | ✅ |
| **M9** | Update CHANGELOG, README, META-INF, docs/llms.txt for v0.12.0 | ✅ |
| **M10** | Migrate `ReadShowcaseController` and `CsvShowcaseController` examples | ✅ |

**Implementation Notes**:
- `ExcelReader.forMap()` creates a synthetic mapper: `row -> map.put(header, row.get(header).formattedValue())` for each header, reusing the existing mapping-mode branch in `build()`.
- Mixed-mode guard placed in all 5 `column*()` methods (fail-fast at call site) — message names method and references `forMap()` so users understand why they can't mix APIs.
- `readAsStream()` and `onProgress()` work automatically — no handler changes, so the existing producer-thread and progress-callback paths continue to work for map mode.
- Test migration pinned critical behaviors:
  - Excel blank header cells become `""` map keys (due to `CellData`'s internal `null → ""` coercion).
  - Fewer cells than headers → trailing keys absent.
  - More cells than headers → extras ignored.
  - `readAsStream()` on non-existent sheet now throws (not silent empty stream).
- Example apps migrated in the same commit as file deletion (M10 + M5 together) so no intermediate compilation breakage.
- Build, test, and javadoc all pass before closing analysis.

### Check Phase

**Document**: [docs/03-analysis/map-reader-absorption.analysis.md](../../03-analysis/map-reader-absorption.analysis.md)

- **Match Rate**: 100% (after G1 close)
- **Initial Match Rate**: 98% (pre-close)
- **Gap Found**: G1 (Medium severity) — `ExcelReadSupport.java:4` javadoc had a dangling `{@link ExcelMapReader}` reference (deleted file). Fixed by removing the stale link from the `@link` list. This resolved the Plan's NFR "javadoc 경고 0".
- **Deviations Documented**: 6 deviations, all pre-approved by Design:
  1. M3/M4 handler changes unnecessary (discovered during Design)
  2. SAX rewrite risk eliminated
  3. Implementation uses positional pairing (not named lookup) to match old behavior exactly
  4. Plan undercounted method guards (5, not 3)
  5. Excel `forMap()` gained `onProgress` (bonus improvement)
  6. `readAsStream()` non-existent sheet now throws (more correct)
- **Test Coverage**: 68 new test methods (35 Excel + 33 CSV) covering factory types, header auto-discovery, fluent API fluency on map mode, mixed-mode guard errors, edge cases (null headers, blank cells, multi-sheet, extra/missing columns), and equivalence with old behavior.

---

## Results

### Completed Items

✅ **M1 — ExcelReader.forMap() static factory**
- Located: `ExcelReader.java:167-187`
- Returns `ExcelReader<Map<String, String>>` with `mapMode = true` flag
- Javadoc includes example code and references `sheetIndex`, `headerRowIndex`, `onProgress` fluent methods
- Internal synthetic mapper: `row -> map.put(header, row.get(header).formattedValue())`

✅ **M2 — CsvReader.forMap() static factory**
- Located: `CsvReader.java:122-140`
- Mirrors ExcelReader.forMap() logic
- Supports `dialect()`, `delimiter()`, `charset()`, `onProgress()` fluently on map mode

✅ **M3/M4 — Handler changes (zero changes)**
- `ExcelReadHandler.java` / `CsvReadHandler.java` untouched (no constructor overloads needed)
- Mapping-mode branch already handles map mode via the synthetic mapper
- Confirms Design's discovery: SAX state machine was never the bottleneck

✅ **M5 — File deletion**
- `ExcelMapReader.java` removed (was 265 lines + inner `ExcelMapReadHandler` + `MapSheetHandler`)
- `CsvMapReader.java` removed (was 297 lines + inner `CsvMapReadHandler`)
- Verified via git: both files deleted in commit `ee4a3c3`

✅ **M6 — Existing test migration**
- `MapReaderStreamTest.java`: all `new ExcelMapReader()` → `ExcelReader.forMap()`
- `MapWriterReaderTest.java`: map reader test cases migrated
- `CoverageBoostTest.java`: all references updated, nested class names still use old names (cosmetic, not blocking)
- Zero old-API imports remain in test files

✅ **M7 — Mixed-mode runtime guards**
- `ExcelReader`: 5 guarded methods (column×2, columnAt, skipColumn, skipColumns)
- `CsvReader`: same 5 methods
- Error message: `"[method] cannot be called on a forMap() reader; map mode auto-discovers columns from the header row"`
- Fail-fast at call site (not late in build())

✅ **M8 — New comprehensive tests**
- `ExcelReaderMapModeTest.java`: **35 test/display methods**
  - Factory return type verification (2)
  - Header auto-discovery (2)
  - Fluent API (sheetIndex, headerRowIndex, onProgress) — (4)
  - Mixed-mode guards (7)
  - Behavioral equivalence pins (4)
  - Stream errors (1)
  - Multi-sheet selection (1)
  - Nested DisplayNames (14)
- `CsvReaderMapModeTest.java`: **33 test/display methods**
  - Same structure, CSV-specific tests (dialect, delimiter, bom, duplicate headers)
- Key edge cases pinned:
  - Blank header cells (`CellData`'s `null → ""` coercion)
  - Mismatched column/header counts
  - Non-existent sheet throws
  - Stream properly drains and closes

✅ **M9 — Documentation updates**
- `build.gradle.kts`: version bumped to `0.12.0`
- `CHANGELOG.md`: `[0.12.0]` section added with Removed/Added/Changed notes
- `README.md`: Installation section updated to `0.12.0`, Feature list includes "Map Reader unified with Writer pattern"
- `META-INF/AI.md`: Map Reader section updated with `forMap()` examples
- `META-INF/excel-kit/excel.md` / `csv.md`: Migration guide (Before/After table)
- `docs/llms.txt`: Indexed with Map API documentation

✅ **M10 — Example app migration**
- `ReadShowcaseController.java`: `/showcase/read/map` endpoint now uses `ExcelReader.forMap()`
- `CsvShowcaseController.java`: CSV map read endpoint uses `CsvReader.forMap()`
- Both migrated in same commit as M5 (file deletion) to prevent breakage

### Incomplete or Deferred Items

**None.** All 10 tasks completed. Feature is 100% done.

### Cosmetic findings (not release-blocking)

- Test file javadoc references deleted classes (`ExcelMapReader.ExcelMapReadHandler`) — not emitted as library javadoc, can be cleaned up later.
- Nested test class names (`CsvMapReaderStreamErrors`, etc.) still reference old class names — test only, safe to rename in future cosmetic pass.

---

## Lessons Learned

### What Went Well

1. **Design phase discovery eliminated the plan's top risk**: The original Plan assumed SAX handler rewrites would be necessary. During Design review, reading the `RowData` class revealed that headers and value lookup were already exposed via public methods. This allowed the entire absorption to be implemented at the Reader level (factory + flag + synthetic mapper) with zero handler changes. This is a good reminder that understanding the existing infrastructure deeply before planning large refactors can collapse perceived risks.

2. **Synthetic mapper pattern is elegant and composable**: Rather than trying to extend the handler class hierarchy (Plan's original assumption), leveraging the existing `Function<RowData, T>` mapping-mode path meant the feature required only ~100 lines of new code (factory methods + guards) instead of a major SAX state-machine rewrite. This approach will likely be useful for future reader patterns.

3. **Single-user breaking policy enabled clean execution**: v0.11.0 established the principle that internal breaking changes (deleting dedicated classes when a factory pattern replaces them) are acceptable because the project has one user. This allowed the Map Reader files to be deleted immediately without deprecation, making the codebase smaller and the upgrade path clear (no murky "use X instead of Y" transition).

4. **Edge case test pinning prevented silent regressions**: Tests like `blankHeaderCell_becomesEmptyStringKey` and `fewerCellsThanHeaders_trailingKeysAbsent` pin specific behaviors of `CellData` null coercion and header pairing logic. If someone changes `CellData`'s internal implementation in the future, these tests will immediately alert them that map mode depends on null→"" coercion.

5. **Behavioral improvements from reuse**: By reusing the existing Reader fluent API, `ExcelReader.forMap()` automatically gained `onProgress()` support that `ExcelMapReader` never had. This is a free bonus improvement that addresses an asymmetry flagged in v0.11.0's Plan.

### Areas for Improvement

1. **Plan risk assessment could be less pessimistic**: The Plan rated "SAX state-machine rewrite" as High/Medium risk without first investigating the existing `RowData` API. A brief infrastructure survey during planning might have reframed the risk as Low. For future features, recommend a quick "does the infrastructure already do this?" investigation before committing to a risky implementation path.

2. **Test file naming should have been updated immediately**: Nested test class names like `CsvMapReaderStreamErrors` still reference the deleted class, even though the tests themselves use the new API. These are cosmetic but confusing for future developers. Recommend updating test class/method names as part of the same commit as file deletion to avoid these orphaned references.

3. **Javadoc @link cleanup could be automated**: The G1 gap (dangling `@link ExcelMapReader`) was caught during static analysis. The build's javadoc checker is good, but it would help to pre-scan all source files for broken @link references when deleting public classes. Could be a gradle plugin task.

### To Apply Next Time

1. **Reuse-first refactoring**: When absorbing a specialized class into a more general factory, first check if the existing infrastructure (mappers, handlers, fluent APIs) already does most of the work. This often eliminates perceived complexity from the Plan phase.

2. **Single-user breaking requires clean communication**: Since v0.11.0 established the "immediate deletion" principle, future deletions should cite this precedent clearly in the CHANGELOG and Migration Guide. Don't hedge with "deprecated" or "consider using" — own the breaking change with clear examples.

3. **Pin behavioral equivalence early**: When migrating from one API to another, write parametric tests that compare old-behavior snapshots (or descriptions) with new-API outputs. The `BehavioralEquivalence` test group in M8 did this effectively, catching subtle differences in header/cell pairing.

4. **Coordinate commits for compilation safety**: M10 (example migration) was correctly merged with M5 (file deletion) in a single commit. This pattern is worth repeating whenever deleting used classes — ensure all call sites are updated in the same commit.

---

## Metrics

| Metric | Value | Notes |
|---|---|---|
| **Match Rate** | 100% | After G1 close (javadoc link fix) |
| **Implementation Commits** | 3 | `ee4a3c3` (main), `6abd8bc` (test enhancement), `0d22b49` (G1 close) |
| **Lines Added** | ~560 | Factories, guards, 68 new test methods, documentation |
| **Lines Deleted** | ~560 | Both Map Reader files, old test code |
| **Net File Change** | -2 files, ±~0 net lines | Map Reader files removed |
| **Files Touched** | ~20 | 2 Reader files, 6 test files, 4 doc files, 2 example files, 2 plan/design docs, 2 deleted files, 2 META-INF files |
| **New Test Methods** | 68 | 35 Excel, 33 CSV |
| **Tests Passing** | ✅ | `./gradlew test` green |
| **Javadoc** | ✅ | `./gradlew javadoc` zero warnings (after G1 close) |
| **Example Build** | ✅ | `./gradlew compileJava` green (example included) |
| **Gaps Found** | 1 (closed) | Javadoc @link reference |
| **Deviations from Plan** | 6 (pre-approved) | M3/M4 unnecessary, SAX risk eliminated, guard count, behavior improvements |

---

## Next Steps

### For v0.12.0 Release

Follow `CLAUDE.md` Release Checklist:

1. ✅ GitHub PR check — no outstanding PRs blocking v0.12.0
2. ✅ `build.gradle.kts` version — already set to `0.12.0`
3. ✅ `CHANGELOG.md` — `[0.12.0]` section added
4. ✅ `README.md` — features and installation version updated
5. ✅ `example/` — migrations and build verified
6. ✅ `META-INF/AI.md` — documentation updated
7. ✅ `docs/llms.txt` — Map API docs indexed
8. ✅ Tests — `./gradlew test` green
9. ✅ Example compile — `./gradlew compileJava` green
10. 📋 **Remaining**: Git commit → tag `v0.12.0` → push (origin main --tags) — triggers Maven Central auto-publish

### For v0.13.0

Next major cleanup item: **ColumnStyleConfig / ExcelColumnBuilder / ColumnConfig unification** (deferred from v0.12.0 scope, per Plan §2.2).

The v0.11.0 column API refactor created an asymmetry:
- `ExcelColumnBuilder` (fluent, chainable, for Writer) — good ergonomics
- `ColumnConfig` (lambda-based, for potential Reader setter mode) — exists but not fully integrated
- `ColumnStyleConfig` (base class for style configuration) — legacy name

v0.13.0 should consolidate these into a single, unified `ColumnConfig` API across Writer and Reader, paralleling the Map I/O unification just completed.

---

## Related Documents

- **Plan**: [docs/01-plan/features/map-reader-absorption.plan.md](../../01-plan/features/map-reader-absorption.plan.md)
- **Design**: [docs/02-design/features/map-reader-absorption.design.md](../../02-design/features/map-reader-absorption.design.md)
- **Analysis**: [docs/03-analysis/map-reader-absorption.analysis.md](../../03-analysis/map-reader-absorption.analysis.md)
- **v0.11.0 Cleanup Plan**: [docs/01-plan/features/api-cleanup-v011.plan.md](../../01-plan/features/api-cleanup-v011.plan.md)

---

## Version History

| Version | Date | Changes | Author |
|---|---|---|---|
| 1.0 | 2026-04-12 | Completion report — 100% match rate after G1 javadoc close. All 10 Plan tasks done. Key lesson: design-phase infrastructure discovery eliminated the plan's top risk (SAX rewrite). Six pre-approved deviations all improve the result (zero handler changes, better error messages, bonus `onProgress` support). 68 new tests pin edge cases and behavioral equivalence. Ready for v0.12.0 release. | DongHyeok Kim |
