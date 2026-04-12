# Changelog

All notable changes to this project will be documented in this file.

## [0.15.0] - 2026-04-12

### Added

- **`FileHandler.toFile(Path)`** — convenience default method that writes directly to a
  file path without manually opening an OutputStream. Works for both Excel and CSV:
  `writer.write(stream).toFile(Path.of("output.xlsx"))`.
- **`ExcelReader.password(String)`** — enables reading password-encrypted Excel files.
  Symmetric with `ExcelWriter.password()` for writing. Uses POI's Decryptor API.
- **`CsvWriter.constColumnIf(name, condition, value)`** — conditional constant column,
  symmetric with `ExcelWriter.constColumnIf()`.

### Fixed

- **BookJpaRepository JPQL path** — example app's JPQL query referenced a wrong package
  (`book.domain.BookDto` instead of `book.application.port.out.BookDto`). Pre-existing bug
  unrelated to library code.

### Changed

- **Release checklist strengthened** — CLAUDE.md now requires `./gradlew :example:bootRun`
  to verify Spring context initialization (catches JPQL/HQL errors that `compileJava` misses),
  plus `./gradlew :kit:javadoc` for zero-warning verification.

## [0.14.0] - 2026-04-12

v0.14.0 is a comprehensive API cleanup and internal refactoring release.

### Changed (Breaking)

- **Package renamed: `shared` → `core`** — `io.github.dornol.excelkit.shared` is now
  `io.github.dornol.excelkit.core`. All imports must be updated. The package holds the
  library's central abstractions (FileHandler, CellData, ReadColumn, RowData, etc.) —
  "core" describes their role more accurately.
- **`ExcelRowFunction` and `CsvRowFunction` deleted** — both were empty marker interfaces
  extending `RowFunction` with zero additional methods. Use `RowFunction` (in `core/`)
  directly.
- **`ExcelConsumer` renamed to `WriteRowCallback`** — better signals that it's a per-row
  callback during writes, not a generic consumer.
- **`ExcelReadColumn` and `CsvReadColumn` merged into `ReadColumn`** — identical records
  consolidated into `core/ReadColumn<T>`.
- **`ExcelWorkbook` constructors removed** — use `ExcelWorkbook.builder()` (same pattern
  as `ExcelWriter.builder()`).
- **`ExcelReader.configureLargeFileSupport()` moved to `ExcelKitConfig`** — JVM-global
  POI configuration belongs at application bootstrap level, not on a reader class.
- **`ExcelWorkbook.finish()` now throws on duplicate calls** — second call to `finish()`
  throws `ExcelWriteException` instead of silently creating a new handler.

### Added

- **`ExcelReader.setter()` / `CsvReader.setter()` static factories** — all three read
  modes now have symmetric entry points: `setter()`, `mapping()`, `forMap()`.
- **`ExcelReader(Supplier)` / `CsvReader(Supplier)` no-validator constructors** —
  eliminates the `null` parameter: `new ExcelReader<>(User::new)` instead of
  `new ExcelReader<>(User::new, null)`.
- **`ExcelWriter.forMap(Builder, String...)` overload** — allows setting color, maxRows,
  rowAccessWindowSize on map-mode writers.
- **`ExcelReader.forMap(String...)` / `CsvReader.forMap(String...)` column selection** —
  include only specified headers in the output map.
- **`ExcelKitConfig` utility class** — centralized JVM-global POI configuration.
- **`ExcelColumn.of(name, fn, style, setter)` factory** — 4-param shortcut for the
  17-param constructor (mainly useful in tests).
- **`CellStyleParams.of()` convenience factories** — eliminates 11-null padding in
  common call patterns.

### Internal improvements

- **Javadoc warnings: 100 → 0** — all public APIs now have complete javadoc.
- **`ExcelWriteSupport` shared methods** — `writeAfterDataAndSummary()` and
  `applyPostProcessing()` extracted, eliminating duplication between ExcelWriter and
  ExcelSheetWriter. `writeRowCells()` now accepts `SheetConfig<T>` directly (9→7 params).
- **`TemplateListWriter` → `SheetConfig` delegation** — 8 duplicate fields replaced with
  a shared `SheetConfig<T>` instance.
- **`AbstractReadHandler` validation helpers** — `validateHeaderRowIndex()` and
  `validateColumns()` eliminate repeated if/throw blocks across 4 handler constructors.
- **Mutable array hack removed** — `ExcelSheetWriter.write()` uses an Iterator loop
  instead of `SXSSFSheet[]` array workaround.
- **Magic numbers extracted** — `WIDTH_PER_CHAR`, `WIDTH_BASE_PADDING`,
  `DEFAULT_FONT_SIZE`, `FONT_HEIGHT_MULTIPLIER`, `LUMINANCE_DARK_THRESHOLD`,
  `DEFAULT_ROW_HEIGHT_POINTS` are now named constants.
- **Defensive array copies** — `ExcelColumn.getDropdownOptions()` and
  `getHeaderFontColor()` return cloned arrays.
- **`ColumnStyleConfig` fields reorganized** — grouped by category (Layout, Font, Color,
  Borders, Validation, Protection, Grouping) with section comments.
- **Package-info expanded** — all three package descriptions upgraded from one-liners to
  meaningful summaries listing key classes and capabilities.
- **Error message consistency** — duplicate-name errors now use consistent `'name'`
  quoting format.
- **ExcelSheetWriter javadoc** — documents the `ColumnConfig` vs `ExcelColumnBuilder`
  design difference.

### Migration Guide

```java
// ─── Package rename ───
// Before
import io.github.dornol.excelkit.shared.FileHandler;
import io.github.dornol.excelkit.shared.CellData;
// After
import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.core.CellData;

// ─── RowFunction ───
// Before
ExcelRowFunction<User, Object> fn = (user, cursor) -> user.getName();
CsvRowFunction<User, Object> fn = (user, cursor) -> user.getName();
// After
RowFunction<User, Object> fn = (user, cursor) -> user.getName();

// ─── ExcelConsumer → WriteRowCallback ───
// Before
writer.write(stream, (ExcelConsumer<User>) (row, cursor) -> log(row));
// After
writer.write(stream, (WriteRowCallback<User>) (row, cursor) -> log(row));

// ─── ExcelWorkbook ───
// Before
new ExcelWorkbook(ExcelColor.STEEL_BLUE)
// After
ExcelWorkbook.builder().color(ExcelColor.STEEL_BLUE).build()

// ─── Reader setter mode ───
// Before
new ExcelReader<>(User::new, null)
// After (pick one)
new ExcelReader<>(User::new)          // no-validator constructor
ExcelReader.setter(User::new)         // static factory (symmetric with mapping/forMap)

// ─── Large file support ───
// Before
ExcelReader.configureLargeFileSupport();
// After
ExcelKitConfig.configureLargeFileSupport();
```

## [0.13.0] - 2026-04-12

### Changed (Breaking)

- **`ExcelSheetWriter.ColumnConfig` and `TemplateListWriter.ColumnConfig` inner classes deleted** —
  replaced by a single top-level `ColumnConfig<T>` class that extends `ColumnStyleConfig<T, ColumnConfig<T>>`.
  Both inner classes were empty wrappers (zero methods) inheriting all 47 styling methods from
  `ColumnStyleConfig`. The new class is identical in behavior; only the qualified name changes.
  `ExcelColumn.ExcelColumnBuilder` is unaffected (it has its own unique `style()` / `build()` methods).

### Migration Guide

```java
// Before (v0.12.0) — qualified references
ExcelSheetWriter.ColumnConfig<MyRow> config = new ExcelSheetWriter.ColumnConfig<>();
TemplateListWriter.ColumnConfig<MyRow> config2 = new TemplateListWriter.ColumnConfig<>();

// After (v0.13.0) — top-level class
ColumnConfig<MyRow> config = new ColumnConfig<>();

// Lambda-style callers are unchanged (type inferred):
sheetWriter.column("Name", row -> row.getName(), cfg -> cfg.bold(true).type(ExcelDataType.STRING));
```

## [0.12.0] - 2026-04-12

v0.12.0 completes the Map I/O symmetry that v0.11.0 deferred: the Map Reader
classes are absorbed into `ExcelReader.forMap()` / `CsvReader.forMap()` static
factories, matching the Writer side.

### Changed (Breaking)

- **`ExcelMapReader` class removed** — use `ExcelReader.forMap()`. The returned
  reader is an `ExcelReader<Map<String, String>>` with the full fluent API
  (`sheetIndex`, `headerRowIndex`, `onProgress`, `readAsStream`). Gains
  `onProgress` support that `ExcelMapReader` never had.
- **`CsvMapReader` class removed** — use `CsvReader.forMap()`. Same benefits:
  full `CsvReader` API including `dialect`, `delimiter`, `charset`,
  `headerRowIndex`, `onProgress`, `readAsStream`.
- **Mixed-mode runtime guard** — calling `column(setter)`, `column(name, setter)`,
  `columnAt(idx, setter)`, `skipColumn()`, or `skipColumns(int)` on a reader
  obtained from `forMap()` now throws `IllegalStateException`. Map mode
  auto-discovers columns; the setter API doesn't apply.

### Behavioral notes

- **`readAsStream` on a non-existent sheet** now throws `ExcelReadException`
  via `ExcelReadHandler`'s missing-sheet check. The deleted `ExcelMapReader`
  silently returned an empty stream, which hid caller bugs.
- Map building still uses positional pairing truncated at
  `min(headerCount, cellCount)` — matches the deleted Map Readers exactly.

### Migration Guide

```java
// ─── Excel map reading ───
// Before (v0.11.0)
new ExcelMapReader()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> process(r.data()));

// After (v0.12.0)
ExcelReader.forMap()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> process(r.data()));

// ─── CSV map reading ───
// Before
new CsvMapReader()
    .dialect(CsvDialect.EXCEL)
    .onProgress(1000, (count, total) -> System.out.println(count))
    .build(inputStream)
    .read(r -> process(r.data()));

// After
CsvReader.forMap()
    .dialect(CsvDialect.EXCEL)
    .onProgress(1000, (count, total) -> System.out.println(count))
    .build(inputStream)
    .read(r -> process(r.data()));

// ─── Excel map reading now supports onProgress (new) ───
ExcelReader.forMap()
    .onProgress(1000, (count, total) -> System.out.println("read " + count))
    .build(inputStream)
    .read(r -> process(r.data()));
```

### Internal note

Absorption reuses the existing mapping-mode infrastructure
(`Function<RowData, T>`) via a synthetic `Function<RowData, Map<String, String>>`.
No SAX handler or `ExcelReadHandler` / `CsvReadHandler` changes were needed —
`RowData` already exposes `headerNames()` and `get(name)`, so the entire
"map reader" can be expressed as a 5-line row mapper. This eliminated the
"SAX callback state-machine rewrite" risk called out in the Plan.

## [0.11.0] - 2026-04-12

v0.11.0 is an **API cleanup release**. It removes a handful of parallel or
stale entry points that accumulated through v0.9.x–v0.10.0 and lands the
Reader-side half of the "unified column API" work that started in v0.10.0.
No new features.

### Changed (Breaking)

- **`ExcelWriter` constructors removed** — all 5 public constructors
  (`ExcelWriter()`, `ExcelWriter(maxRows)`, `ExcelWriter(color)`,
  `ExcelWriter(color, maxRows)`, `ExcelWriter(color, maxRows, windowSize)`)
  are deleted. Use `ExcelWriter.<T>builder()` instead.
- **`FileHandler` interface + `write()` rename** — `ExcelHandler` and
  `CsvHandler` now implement `shared.FileHandler` and expose
  `write(OutputStream)` instead of `consumeOutputStream(OutputStream)`.
  Both handler classes are now `final`. `ExcelHandler`'s
  `consumeOutputStreamWithPassword` Excel-only overloads are unchanged.
  `FileHandler` is a plain `interface` rather than `sealed` because
  excel-kit ships as an automatic module (no `module-info.java`); the
  `final` implementations preserve the closed-hierarchy intent and
  third-party implementations remain unsupported.
- **Reader column API unified** — `ExcelReader` / `CsvReader` no longer have
  `addColumn`, `columnAtBuilder`, `ExcelReadColumnBuilder`, or
  `CsvReadColumnBuilder`. `column(setter)` and `column(name, setter)` now
  return the reader itself (previously returned a chain-continuation
  builder). `columnAt(int, setter)` is unchanged.
- **`ExcelMapWriter` and `CsvMapWriter` deleted** — replaced by
  `ExcelWriter.forMap(...)` and `CsvWriter.forMap(...)` static factories
  that return the underlying writer. Use the writer's full fluent API
  directly instead of reaching through `.writer()` or a limited set of
  shortcut methods.

### Migration Guide

```java
// ─── ExcelWriter construction ───
// Before
new ExcelWriter<User>()
new ExcelWriter<User>(ExcelColor.STEEL_BLUE)
new ExcelWriter<User>(ExcelColor.STEEL_BLUE, 500_000, 500)
// After
ExcelWriter.<User>builder().build()
ExcelWriter.<User>builder().color(ExcelColor.STEEL_BLUE).build()
ExcelWriter.<User>builder()
    .color(ExcelColor.STEEL_BLUE).maxRows(500_000).rowAccessWindowSize(500).build()

// ─── Writing the output ───
// Before
handler.consumeOutputStream(out)
// After
handler.write(out)

// ─── Reader column binding ───
// Before
reader.addColumn(User::setName)
reader.addColumn("Name", User::setName)
reader.column(User::setName)          // (returned a builder)
reader.columnAtBuilder(2, User::setAge)
// After
reader.column(User::setName)          // now returns Reader<T>
reader.column("Name", User::setName)
reader.columnAt(2, User::setAge)

// ─── Map writers ───
// Before
new ExcelMapWriter("Name", "Age").write(stream).consumeOutputStream(out)
new CsvMapWriter("Name", "Age").dialect(CsvDialect.EXCEL).write(stream).consumeOutputStream(out)
// After
ExcelWriter.forMap("Name", "Age").write(stream).write(out)
CsvWriter.forMap("Name", "Age").dialect(CsvDialect.EXCEL).write(stream).write(out)
```

### Deferred to v0.12.0

- Map Reader absorption — `ExcelMapReader` and `CsvMapReader` remain as
  standalone classes. Their header-auto-detect logic is woven into the
  SAX-style row callbacks, so folding them into `ExcelReader.forMap()` /
  `CsvReader.forMap()` needs a separate refactor and is scheduled for
  v0.12.0.

## [0.10.0] - 2026-04-09

### Changed (Breaking)
- **Unified column API** — `ExcelWriter` now uses the same `.column()` / `.columnIf()` / `.constColumn()`
  pattern as `ExcelSheetWriter` and `CsvWriter`. All column methods return `ExcelWriter<T>` for chaining.
  Column configuration is done via lambda configurer: `.column("Name", fn, cfg -> cfg.type(...).bold(true))`.
- **Removed builder-chaining style** — The old `ExcelColumnBuilder` navigation methods (`column()`,
  `columnIf()`, `constColumn()`, `write()`, `beforeHeader()`, `afterData()`, `onProgress()`) are removed.
  `ExcelColumnBuilder` is now only used internally for column configuration.
- **Renamed `addColumn` → `column`** on `ExcelWriter` for consistency across all writer APIs.

### Added
- **`columnIf` on `ExcelWriter`** — conditional column with all 4 overloads
  (Function, Function+Consumer, ExcelRowFunction, ExcelRowFunction+Consumer).
- **`constColumn` with configurer** — `.constColumn("name", value, cfg -> cfg.type(...))`.
- **`constColumnIf`** — conditional constant column.

### Migration Guide
```java
// Before (0.9.x)
writer.column("Price", Product::price).type(ExcelDataType.INTEGER).format("#,##0")
writer.addColumn("Price", Product::price, cfg -> cfg.type(ExcelDataType.INTEGER))

// After (0.10.0)
writer.column("Price", Product::price, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0"))
```

## [0.9.6] - 2026-04-08

### Added
- **AI context documentation** in JAR (`META-INF/AI.md` + `META-INF/excel-kit/*.md`) —
  structured documentation for AI agents to discover library usage when exploring dependencies.
- **`llms.txt`** published to GitHub Pages for web-accessible AI context.
- **`CLAUDE.md`** with release checklist and project conventions.

## [0.9.5] - 2026-04-08

### Added
- **Per-column header font color** via `headerFontColor(ExcelColor)` / `headerFontColor(int, int, int)` —
  override the header font color for individual columns. Useful for conditionally highlighting
  specific column headers (e.g., error indicators). Available on both `ExcelWriter` (builder chaining)
  and `ExcelSheetWriter` (lambda config). Passes `null` to use the default header style.

## [0.9.4] - 2026-04-01

### Added
- **`ExcelWriter.password(String)`** / **`ExcelWorkbook.password(String)`** — set encryption password
  at the writer level. `write()` automatically encrypts without needing
  `consumeOutputStreamWithPassword()`. Consistent with `protectSheet()` / `protectWorkbook()` API pattern.
- **CsvMapReader** — read CSV files into `Map<String, String>` without typed POJOs,
  matching the `ExcelMapReader` API. Supports `dialect()`, `delimiter()`, `charset()`,
  `headerRowIndex()`, `onProgress()`, and `readAsStream()`.
- **CsvMapWriter** — `dialect()`, `delimiter()`, `charset()`, `bom()` configuration
  methods for symmetry with `CsvMapReader`.
- README: Quick Reference table for all read/write classes.
- Benchmark CI workflow — runs performance tests on every main-branch push.
- Read benchmarks — Excel MapReader/typed 100K rows, CSV MapReader/typed 1M rows.
- Multi-JDK CI — test matrix across JDK 17, 21, and 25.
- Javadoc GitHub Pages workflow — auto-deploy API docs on release.
- Example app: CSV Map Reader upload endpoint.

### Improved
- `ExcelHandler` internals refactored: `markConsumed()` and `writePlain()` extracted
  to eliminate duplication between encrypted and plain-text write paths.
- `char[]` password validation now rejects blank (whitespace-only) arrays,
  consistent with `String` password validation.
- Calling `consumeOutputStreamWithPassword()` when `password()` is already set throws
  `IllegalStateException` with a descriptive message instead of silently using
  the wrong password.
- Test coverage boost: 38 new targeted tests for CsvMapReader, CsvWriter quoting,
  AbstractReadHandler validation, and TempResourceContainer edge cases.

### Fixed
- `ExcelHandler.consumeOutputStreamWithPassword(char[])`: char array is now zeroed
  even when `IllegalStateException` is thrown due to password conflict.
  Previously the array was left intact if the exception occurred before the
  `try/finally` block.

### Dependencies
- poi-ooxml 5.4.1 → 5.5.1
- junit-bom 5.10.0 → 6.0.3
- actions/checkout v4 → v6
- actions/upload-artifact v4 → v7
- actions/setup-java v4 → v5
- Dependabot configured for Gradle and GitHub Actions.

## [0.9.2] - 2026-03-30

### Added
- **Data bar conditional formatting** via `dataBar(ExcelColor)` — gradient bars
  proportional to cell values. Supports single-color and 2-color gradient
  (`dataBar(minColor, maxColor)`).
- **Icon set conditional formatting** via `iconSet(IconSetType)` — 10 icon set
  types including arrows, traffic lights, flags, signs, symbols, ratings, quarters.
- **Timezone-aware date parsing** via `CellData.asZonedDateTime(ZoneId)` and
  `CellData.asZonedDateTime(String format, ZoneId)`.
- **CSV dialect presets** via `CsvDialect` enum — RFC4180, EXCEL, TSV, PIPE.
  Apply with `CsvWriter.dialect()` and `CsvReader.dialect()`.
- **CSV quoting strategies** via `CsvQuoting` enum — MINIMAL (default), ALL
  (quote everything), NON_NUMERIC (quote strings, leave numbers unquoted).
  Configure with `CsvWriter.quoting()`.
- README: Supported Formats table, Notes section (JVM-global config warning,
  readAsStream try-with-resources requirement).

### Changed
- **ExcelMapReader.readAsStream()**: Converted from List-collect approach to true
  streaming via BlockingQueue + producer thread (same pattern as ExcelReadHandler).
  Now memory-efficient for large datasets.

### Improved
- Branch test coverage: 84% → 89% (+82 new tests, +46 branches covered).
- Test assertion quality: replaced `assertTrue(out.size() > 0)` patterns with
  actual POI API content verification (validation rules, chart types, cell values,
  font styles, formula content).

## [0.9.0] - 2026-03-19

### Added
- **ExcelTemplateWriter** — fill data into existing .xlsx templates while preserving
  formatting, images, charts, and merged regions. Supports cell-level writes,
  list streaming via `TemplateListWriter`, and mixed mode (cell + list + afterData).
- **TemplateListWriter** — stream tabular data into a template sheet with column
  definition, afterData callbacks, summary rows, and all existing styling options.
- Write benchmarks (`WriteBenchmarkTest`) — 7 scenarios measuring Excel/CSV throughput.
- Migration guide in README for 0.8.1 → 0.8.2 breaking changes.
- Performance section in README with throughput table.

## [0.8.2] - 2026-03-18

### Changed
- **CellData default locale**: Changed from `Locale.KOREA` to `Locale.getDefault()`.
  Korean environments are unaffected (JVM default is already `Locale.KOREA`).
  Users who relied on the hard-coded Korean locale should call
  `CellData.setDefaultLocale(Locale.KOREA)` explicitly.
- **ExcelDataType.IMAGE**: Now throws `ExcelWriteException` when a non-`ExcelImage` value
  is passed to an IMAGE column. Previously it silently fell back to `String.valueOf()`.
- **RowData.get(int)**: Now throws `IllegalArgumentException` on negative index.
  Previously it silently mapped negative indices to 0.

### Added
- `SheetConfig<T>` — internal shared configuration class for `ExcelWriter` and
  `ExcelSheetWriter`, eliminating 17 duplicated field declarations.
- JaCoCo test coverage verification with 70% minimum threshold.
- CI workflow uploads JaCoCo report as artifact for PR visibility.
- `readAsStream()` Javadoc warns about try-with-resources requirement.
- `SXSSFSheetHelper` logs warning on reflection failure instead of silent null.
- `ExcelChartConfig.categoryColumn()` / `valueColumn()` reject negative indices.
- `CellData` date format add/reset methods are now synchronized (`FORMAT_LOCK`).
- `ExcelReadHandler` producer thread named `"excel-kit-reader"` for debuggability.
- `CsvReadHandler` has its own dedicated logger.

### Fixed
- Misplaced Javadoc blocks across 6 source files (caused by linter reordering).
- Silent exception swallowing in `ExcelWorkbook.close()` and `CsvReadHandler.closeQuietly()`.
- `assert` statements in production code replaced with `IllegalStateException`
  (`CsvReadHandler`, `ExcelReadHandler`, `AbstractReadHandler`).
- `ExcelHandler.consumeOutputStreamWithPassword(char[])` error message unified to "blank".
- `CsvReadHandler` BOM error message clarified.

## [0.8.1] - 2025-07-29

### Added
- `CellData.as(Function)` — custom type conversion (e.g., `UUID::fromString`).
- `CellData.as(Function, defaultValue)` — custom conversion with default.
- Default value overloads: `asInt(int)`, `asLong(long)`, `asDouble(double)`, `asString(String)`.
- `CsvWriter.csvInjectionDefense(boolean)` — toggle CSV injection defense.
- Round-trip integration tests for CellData conversions.

## [0.8.0] - 2025-07-28

### Added
- **Mapping mode** for immutable object / Java record reading via `ExcelReader.mapping()` and `CsvReader.mapping()`.
- `RowData` class for positional and named cell access.
- Map-based reading via `ExcelMapReader` and `CsvMapReader`.

## [0.7.2] - 2025-07-27

### Added
- Workbook protection via `protectWorkbook()`.
- Header font customization via `headerFontName()` and `headerFontSize()`.
- Default column style via `defaultStyle()`.
- Summary/footer rows via `summary()` — SUM, AVERAGE, COUNT, MIN, MAX.
- Named ranges via `SheetContext.namedRange()`.
- List validation from cell range via `ExcelValidation.listFromRange()`.
