---
template: design
version: 1.2
feature: map-reader-absorption
project: excel-kit
project-version: 0.11.0 → 0.12.0
author: DongHyeok Kim
date: 2026-04-12
status: Draft
---

# map-reader-absorption Design Document

> **Summary**: Reuse the existing mapping-mode infrastructure via a synthetic `Function<RowData, Map<String, String>>`. Zero changes to SAX handlers. `ExcelMapReader` / `CsvMapReader` files deleted. Major risk from the Plan (SAX state-machine rewrite) is eliminated.
>
> **Project**: excel-kit
> **Version**: 0.11.0 → 0.12.0
> **Author**: DongHyeok Kim
> **Date**: 2026-04-12
> **Status**: Draft
> **Planning Doc**: [map-reader-absorption.plan.md](../../01-plan/features/map-reader-absorption.plan.md)

---

## 1. Overview

### 1.1 Design Goals

1. Complete the Map I/O symmetry deferred from v0.11.0 (Writer had `forMap()`; Reader didn't)
2. **Zero risk to the SAX parsing path** — do not rewrite `ExcelReadHandler.SheetHandler` or the equivalent CSV logic
3. Preserve all `ExcelMapReader` / `CsvMapReader` features (header auto-detect, `readAsStream`, `onProgress`, `sheetIndex`, `headerRowIndex`, `dialect`)
4. Delete the two Map Reader files entirely (single-user breaking allowed, per v0.11.0 policy)

### 1.2 Design Principles

- **Reuse, don't rewrite**: the mapping-mode path (`rowMapper: Function<RowData, T>`) already does exactly what a map reader needs — it builds a `RowData` from the SAX callbacks and calls a user function. A synthetic mapper that builds `Map<String, String>` from `RowData` is all we need.
- **No changes to ExcelReadHandler / CsvReadHandler**: the entire absorption happens in the Reader classes, not the Handler classes. This keeps the SAX state machine and `readAsStream()` producer-thread logic untouched.
- **Runtime mixed-mode guard**: `forMap()` reader rejects `column()` / `columnAt()` / `skipColumn()` / `skipColumns()` at runtime with `IllegalStateException`.
- **Single-user breaking**: delete the Map Reader files immediately, no deprecation (v0.11.0 policy).

---

## 2. Architecture

### 2.1 Plan vs Reality — SAX rewrite is unnecessary

The Plan (§5 Risks) rated "SAX callback state-machine rewrite" as **High/Medium** risk. Reading `ExcelReadHandler.java` and `RowData.java` reveals that this rewrite is **not needed**:

| | Setter mode | Mapping mode | Proposed map mode |
|---|---|---|---|
| `columns` field | set | null | null |
| `rowMapper` field | null | user-supplied | **synthetic (built by `forMap()`)** |
| `endRow()` branch | `mapValuesToInstance()` | `mapWithRowMapper(rowData)` | **reuses mapping-mode branch** |
| SAX handler changes | — | — | **none** |

The mapping-mode path already constructs a `RowData` for every row, which exposes `headerNames()` and `get(String headerName)`. A trivial synthetic rowMapper converts that to `Map<String, String>`:

```java
Function<RowData, Map<String, String>> mapMapper = row -> {
    Map<String, String> map = new LinkedHashMap<>();
    for (String header : row.headerNames()) {
        if (header == null) continue;   // match ExcelMapReader's null-filter behavior
        map.put(header, row.get(header).formattedValue());
    }
    return map;
};
```

This is the entire "absorption" of the Map reading logic. The rest of the feature is factory methods, runtime guards, file deletion, and test migration.

### 2.2 Type Hierarchy (After)

```
excel/
├── ExcelReader<T>
│   ├── private boolean mapMode                           [NEW]
│   ├── static forMap() : ExcelReader<Map<String,String>> [NEW]
│   ├── column(...) / columnAt(...) / skipColumn*()       [guarded: throws if mapMode]
│   ├── existing: mapping(rowMapper), sheetIndex, headerRowIndex,
│   │   onProgress, build(InputStream)                    [unchanged]
│   └── ...
│
├── ExcelReadHandler<T>                                   [unchanged]
│   └── SheetHandler                                      [unchanged]
│
├── ExcelMapReader.java                                   [FILE DELETED]
└── ExcelReadColumn, ExcelReadException, etc.             [unchanged]

csv/
├── CsvReader<T>
│   ├── private boolean mapMode                           [NEW]
│   ├── static forMap() : CsvReader<Map<String,String>>   [NEW]
│   ├── column(...) / columnAt(...) / skipColumn*()       [guarded]
│   └── ...                                               [unchanged]
│
├── CsvReadHandler<T>                                     [unchanged]
├── CsvMapReader.java                                     [FILE DELETED]
└── ...

shared/
└── AbstractReadHandler, RowData                          [unchanged]
```

### 2.3 Mixed-mode guard placement

Where to place the runtime check that map-mode readers can't register columns?

| Option | Pros | Cons |
|--------|------|------|
| **A. In the Reader's `column()`/`columnAt()`/`skipColumn()` methods** | Fails at the call site (clear stack trace). No Handler changes. | Four methods each to update per Reader (8 total). |
| B. In `build(InputStream)` | One check per Reader. | Fails late — user gets error after building, confusing stack trace. |
| C. In `ExcelReadHandler` constructor | Handler gets involved. | Same "fails late" issue. Handler doesn't need to know about modes it doesn't own. |

**Selected: Option A.** Fail fast at the ambiguous call site. The 8 checks are two-liners.

```java
// ExcelReader<T>
public ExcelReader<T> column(BiConsumer<T, CellData> setter) {
    if (mapMode) {
        throw new IllegalStateException(
            "column() cannot be called on a forMap() reader; map mode auto-discovers columns from the header row");
    }
    columns.add(new ExcelReadColumn<>(setter));
    return this;
}
```

Same pattern for `column(name, setter)`, `columnAt(idx, setter)`, `skipColumn()`, `skipColumns(int)`.

---

## 3. Type Model

### 3.1 `ExcelReader.forMap()` (M1)

```java
// ExcelReader.java

private boolean mapMode = false;

/**
 * Creates a reader that parses Excel files into {@code Map<String, String>} rows by
 * auto-discovering columns from the header row.
 * <p>
 * The returned reader exposes the standard fluent API ({@link #sheetIndex(int)},
 * {@link #headerRowIndex(int)}, {@link #onProgress(int, ProgressCallback)}) but
 * rejects {@link #column(BiConsumer) column()} / {@link #columnAt(int, BiConsumer)}
 * and the {@code skipColumn*()} methods at runtime — map mode infers the columns
 * automatically and does not use the setter API.
 *
 * <pre>{@code
 * ExcelReader.forMap()
 *     .sheetIndex(0)
 *     .headerRowIndex(0)
 *     .build(inputStream)
 *     .read(result -> {
 *         Map<String, String> row = result.data();
 *         String name = row.get("Name");
 *     });
 * }</pre>
 *
 * @return a new reader in map mode
 * @since 0.12.0
 */
public static ExcelReader<Map<String, String>> forMap() {
    Function<RowData, Map<String, String>> mapMapper = row -> {
        Map<String, String> map = new LinkedHashMap<>();
        for (String header : row.headerNames()) {
            if (header == null) continue;
            map.put(header, row.get(header).formattedValue());
        }
        return map;
    };
    ExcelReader<Map<String, String>> reader = ExcelReader.mapping(mapMapper);
    reader.mapMode = true;
    return reader;
}
```

**Why `mapping(mapMapper)` + flip flag instead of a dedicated constructor**: `ExcelReader.mapping()` already sets `rowMapper` and leaves `columns` empty — that's exactly the state map mode needs. Flipping `mapMode = true` afterward enables the runtime guard. No new constructor, no new `build()` branch.

### 3.2 `CsvReader.forMap()` (M2)

Mirror of Excel. CSV's `mapping(rowMapper)` factory is at `CsvReader.java:77`.

```java
// CsvReader.java

private boolean mapMode = false;

/**
 * Same semantics as {@link ExcelReader#forMap()} but for CSV. Auto-discovers columns
 * from the header row; rejects column/columnAt/skipColumn calls at runtime.
 *
 * @since 0.12.0
 */
public static CsvReader<Map<String, String>> forMap() {
    Function<RowData, Map<String, String>> mapMapper = row -> {
        Map<String, String> map = new LinkedHashMap<>();
        for (String header : row.headerNames()) {
            if (header == null) continue;
            map.put(header, row.get(header).formattedValue());
        }
        return map;
    };
    CsvReader<Map<String, String>> reader = CsvReader.mapping(mapMapper);
    reader.mapMode = true;
    return reader;
}
```

### 3.3 Mixed-mode runtime guards (M7)

A single helper + 5 call sites per Reader:

```java
// ExcelReader.java (and analogous in CsvReader)

private void requireNotMapMode(String method) {
    if (mapMode) {
        throw new IllegalStateException(
            method + " cannot be called on a forMap() reader; "
            + "map mode auto-discovers columns from the header row");
    }
}

public ExcelReader<T> column(BiConsumer<T, CellData> setter) {
    requireNotMapMode("column(BiConsumer)");
    columns.add(new ExcelReadColumn<>(setter));
    return this;
}

public ExcelReader<T> column(String headerName, BiConsumer<T, CellData> setter) {
    requireNotMapMode("column(String, BiConsumer)");
    columns.add(new ExcelReadColumn<>(headerName, setter));
    return this;
}

public ExcelReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
    requireNotMapMode("columnAt(int, BiConsumer)");
    columns.add(new ExcelReadColumn<>(null, columnIndex, setter));
    return this;
}

public ExcelReader<T> skipColumn() {
    requireNotMapMode("skipColumn()");
    columns.add(new ExcelReadColumn<>((instance, cellData) -> {}));
    return this;
}

public ExcelReader<T> skipColumns(int count) {
    requireNotMapMode("skipColumns(int)");
    // ... existing body
}
```

### 3.4 `readAsStream()` and other existing features

**No changes needed.** Because `forMap()` piggybacks on mapping mode:

- `ExcelReadHandler.readAsStream()` already exists and handles mapping mode identically to setter mode (the producer thread doesn't care which branch in `endRow()` is taken).
- `onProgress(int, ProgressCallback)` already fires from `SheetHandler.endRow()` after `consumer.accept(result)` — map mode will get progress callbacks automatically.
- `sheetIndex(int)`, `headerRowIndex(int)` — already handled at the `ExcelReader` fluent-API level and passed through to the Handler.
- CSV's `dialect(...)`, `delimiter(char)`, `charset(Charset)` — same, already handled at `CsvReader` level.

### 3.5 Equivalence with the deleted `ExcelMapReader`

| Behavior | `ExcelMapReader` (before) | `ExcelReader.forMap()` (after) |
|---|---|---|
| Header auto-detect | `MapSheetHandler.endRow` at `rowNum == headerRowIndex` | `SheetHandler.extractHeaderNames()` (mapping-mode path) |
| Null header filter | Filters nulls via `.filter(Objects::nonNull)` | Filtered inside synthetic mapMapper (`if (header == null) continue;`) |
| Map construction | Positional pairing: `map.put(headerNames[i], currentRow[i].formattedValue())` | Named lookup: `map.put(header, row.get(header).formattedValue())` via `RowData.get(headerName)` |
| Map type | `LinkedHashMap<String, String>` | `LinkedHashMap<String, String>` |
| `readAsStream()` | Own producer-thread impl inside `ExcelMapReadHandler` | Reuses `ExcelReadHandler.readAsStream()` (existing producer-thread impl) |
| `onProgress(...)` | **Not supported** on `ExcelMapReader` (only `CsvMapReader` had it) | **Supported** — inherited from `ExcelReader` |

**Two differences from the old behavior, both intentional**:

1. **Named lookup instead of positional** — old code uses `headerNames[i]` and `currentRow[i]` in lockstep. If a header has null cells at specific indices, positional pairing can desync with named lookup. The new code uses `RowData.get(headerName)` which finds the correct cell via `headerIndexMap`. In the common case (no nulls) the two produce identical results; in the null-header edge case the new behavior is more correct.
2. **`ExcelReader.forMap()` gains `onProgress`** — `ExcelMapReader` never supported it (an asymmetry flagged in v0.11.0 Plan as "T5-b"). Absorption automatically fixes it.

Both changes are improvements. CHANGELOG will note them.

---

## 4. New Public API Specification

### 4.1 Changes summary

| Task | File | Added | Deleted |
|------|------|------|---------|
| M1 | `ExcelReader.java` | `mapMode` field, `forMap()` static factory, `requireNotMapMode` helper | — |
| M1 | `ExcelReader.java` | guards in `column(×2)`, `columnAt`, `skipColumn`, `skipColumns` | — |
| M2 | `CsvReader.java` | same additions as M1 | — |
| M3 | `ExcelReadHandler.java` | — | — *(no changes)* |
| M4 | `CsvReadHandler.java` | — | — *(no changes)* |
| M5 | `ExcelMapReader.java` | — | **full file** |
| M5 | `CsvMapReader.java` | — | **full file** |
| M6 | existing tests | new-API calls | old-API imports + class refs |
| M8 | `ExcelReaderMapModeTest.java` | full new file | — |
| M8 | `CsvReaderMapModeTest.java` | full new file | — |
| M10 | `example/**/*.java` | new-API calls | old-API refs (ReadShowcaseController, CsvShowcaseController) |

Note M3/M4 are effectively no-ops — kept in the Plan's task list but don't require code changes. Design §2.1 explains why.

### 4.2 Example usage

```java
// Before (v0.11.0)
List<Map<String, String>> rows = new ArrayList<>();
new ExcelMapReader()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> rows.add(r.data()));

// After (v0.12.0)
List<Map<String, String>> rows = new ArrayList<>();
ExcelReader.forMap()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> rows.add(r.data()));
```

```java
// CSV — before
try (Stream<ReadResult<Map<String, String>>> stream = new CsvMapReader()
        .dialect(CsvDialect.EXCEL)
        .build(inputStream)
        .readAsStream()) {
    stream.forEach(r -> process(r.data()));
}

// CSV — after
try (Stream<ReadResult<Map<String, String>>> stream = CsvReader.forMap()
        .dialect(CsvDialect.EXCEL)
        .build(inputStream)
        .readAsStream()) {
    stream.forEach(r -> process(r.data()));
}
```

---

## 5. Migration Matrix

### 5.1 Excel Map Reader

| Before (v0.11.0) | After (v0.12.0) |
|-------------------|------------------|
| `new ExcelMapReader()` | `ExcelReader.forMap()` |
| `new ExcelMapReader().sheetIndex(1)` | `ExcelReader.forMap().sheetIndex(1)` |
| `new ExcelMapReader().headerRowIndex(2)` | `ExcelReader.forMap().headerRowIndex(2)` |
| `ExcelMapReader.ExcelMapReadHandler` | `ExcelReadHandler<Map<String, String>>` |
| `new ExcelMapReader().build(in).read(r -> ...)` | `ExcelReader.forMap().build(in).read(r -> ...)` |
| `new ExcelMapReader().build(in).readAsStream()` | `ExcelReader.forMap().build(in).readAsStream()` |

### 5.2 CSV Map Reader

| Before | After |
|--------|-------|
| `new CsvMapReader()` | `CsvReader.forMap()` |
| `new CsvMapReader().dialect(EXCEL)` | `CsvReader.forMap().dialect(EXCEL)` |
| `new CsvMapReader().delimiter('\t')` | `CsvReader.forMap().delimiter('\t')` |
| `new CsvMapReader().charset(UTF_8)` | `CsvReader.forMap().charset(UTF_8)` |
| `new CsvMapReader().onProgress(1000, cb)` | `CsvReader.forMap().onProgress(1000, cb)` |
| `new CsvMapReader().build(in).read(...)` | `CsvReader.forMap().build(in).read(...)` |
| `CsvMapReader.CsvMapReadHandler` | `CsvReadHandler<Map<String, String>>` |

---

## 6. Error Handling

### 6.1 Mixed-mode guard errors

All five guarded methods throw the same exception type with a method-specific message:

```
IllegalStateException: column(BiConsumer) cannot be called on a forMap() reader;
    map mode auto-discovers columns from the header row
```

### 6.2 Existing error paths

No changes. `ExcelReadException` / `CsvReadException` / `ExcelKitException` still wrap underlying failures the same way.

---

## 7. Test Plan

### 7.1 New tests (M8)

`ExcelReaderMapModeTest.java` — at least 8 tests:

| # | Test | Verifies |
|---|------|----------|
| 1 | `forMap_returnsMapValuedReader` | Return type is `ExcelReader<Map<String, String>>` |
| 2 | `forMap_readsAllColumnsAutomatically` | Header row becomes map keys; data rows become `Map<String, String>` entries |
| 3 | `forMap_withSheetIndex` | `sheetIndex()` reaches the correct sheet |
| 4 | `forMap_withHeaderRowIndex` | Non-zero `headerRowIndex` works (rows before it are skipped) |
| 5 | `forMap_readAsStream_worksLikeRead` | `readAsStream` produces the same rows as `read` |
| 6 | `forMap_onProgress_firesForMapMode` | `onProgress` callback fires — this is **new** functionality (old `ExcelMapReader` didn't support it) |
| 7 | `forMap_column_throwsIllegalStateException` | Calling `.column(setter)` after `forMap()` throws |
| 8 | `forMap_columnAt_skipColumn_allThrow` | All five guarded methods throw |
| 9 | `forMap_nullHeaderCell_isSkipped` | A null header cell is filtered from the output map |
| 10 | `forMap_equivalenceWithOldBehavior` *(parametric)* | For a sample workbook, `ExcelReader.forMap()` produces the same `Map<String, String>` entries a `new ExcelMapReader()` snapshot would have produced (captured manually) |

`CsvReaderMapModeTest.java` — mirror of Excel, plus:

| # | Test | Verifies |
|---|------|----------|
| 11 | `forMap_dialect_TSV_worksInMapMode` | `dialect()` propagates |
| 12 | `forMap_delimiter_pipe` | Custom delimiter works |
| 13 | `forMap_withBomDisabled` | `bom(false)` works on read side |

### 7.2 Existing test migration (M6)

Files to migrate:
- `kit/src/test/java/io/github/dornol/excelkit/excel/MapReaderStreamTest.java`
- `kit/src/test/java/io/github/dornol/excelkit/excel/MapWriterReaderTest.java` — Map Reader cases only
- `kit/src/test/java/io/github/dornol/excelkit/shared/CoverageBoostTest.java` — any references to `ExcelMapReader` / `CsvMapReader`

Each `new ExcelMapReader()` → `ExcelReader.forMap()`, each `new CsvMapReader()` → `CsvReader.forMap()`. Imports of `ExcelMapReader` / `CsvMapReader` / `ExcelMapReadHandler` / `CsvMapReadHandler` are deleted.

### 7.3 Example migration (M10)

Files to migrate:
- `example/src/main/java/io/github/dornol/excelkit/example/app/showcase/ReadShowcaseController.java`
- `example/src/main/java/io/github/dornol/excelkit/example/app/showcase/CsvShowcaseController.java`

Same find/replace. Must land in **the same commit** as M5 (file deletion) so example compiles.

---

## 8. Implementation Order

1. **M1 + M7 on ExcelReader** — add `mapMode`, `forMap()`, guards. Write test file M8 Excel. Verify tests pass (no file deletion yet).
2. **M2 + M7 on CsvReader** — same, with CSV test file. Verify tests pass.
3. **M6 — Existing test migration** — `MapReaderStreamTest`, `MapWriterReaderTest`, `CoverageBoostTest` to new API. Verify `./gradlew test` still passes.
4. **M10 — Example migration** — `ReadShowcaseController`, `CsvShowcaseController` to new API.
5. **M5 — Delete files** — `ExcelMapReader.java`, `CsvMapReader.java`. Verify `./gradlew test` + `compileJava` pass. **This is the same commit as M10.**
6. **M9 — Documentation** — CHANGELOG `[0.12.0]` + Before/After migration table. README, META-INF/AI.md, docs/llms.txt.

### 8.1 Expected file change footprint

| File | Task | Rough size |
|------|------|------------|
| `ExcelReader.java` | M1, M7 | +50 / -0 |
| `CsvReader.java` | M2, M7 | +45 / -0 |
| `ExcelMapReader.java` | M5 | **-265 (file deleted)** |
| `CsvMapReader.java` | M5 | **-297 (file deleted)** |
| `MapReaderStreamTest.java` | M6 | ±20 |
| `MapWriterReaderTest.java` | M6 | ±15 |
| `CoverageBoostTest.java` | M6 | ±10 |
| `ExcelReaderMapModeTest.java` | M8 | +220 (new) |
| `CsvReaderMapModeTest.java` | M8 | +180 (new) |
| `example/**/*.java` | M10 | ±30 |
| `CHANGELOG.md` | M9 | +50 |
| `README.md` | M9 | ±15 |
| `META-INF/AI.md`, `META-INF/excel-kit/*.md`, `docs/llms.txt` | M9 | ±10 |

**Net**: roughly `-400 / +560` across the tree. Much smaller than the Plan estimated because no SAX handler changes happen.

---

## 9. Breaking Change Summary (CHANGELOG material)

### Removed
1. `ExcelMapReader` class (and inner `ExcelMapReadHandler`, `MapSheetHandler`) — use `ExcelReader.forMap()`
2. `CsvMapReader` class (and inner `CsvMapReadHandler`) — use `CsvReader.forMap()`

### Added
1. `ExcelReader.forMap()` static factory
2. `CsvReader.forMap()` static factory
3. Runtime `IllegalStateException` on `ExcelReader.column*()` / `CsvReader.column*()` when called on a `forMap()` reader

### Behavioral notes
1. `ExcelReader.forMap()` **now supports `onProgress(int, ProgressCallback)`** — the deleted `ExcelMapReader` didn't. Symmetric with `CsvMapReader`, which always had it.
2. Map building uses **named lookup via `RowData.get(headerName)`** instead of positional pairing. For the common case (no null header cells) behavior is identical; the edge case where header cells are null in the middle of the row is now more robust.

### 변경 없음 (범위 밖)
- `ExcelReadHandler`, `CsvReadHandler`, `AbstractReadHandler`, `RowData` — no code changes
- `ColumnStyleConfig` / `ExcelColumnBuilder` / `ColumnConfig` — separate feature
- `ExcelTemplateWriter`, `TemplateListWriter`

---

## 10. Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 0.1 | 2026-04-12 | Initial draft — discovery that mapping-mode infrastructure enables absorption via synthetic rowMapper with no SAX handler changes. Plan's High risk (SAX rewrite) eliminated. Approach reduced to Reader-level additions only. | DongHyeok Kim |
