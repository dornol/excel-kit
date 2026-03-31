# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
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
- Test coverage boost: 38 new targeted tests for CsvMapReader, CsvWriter quoting,
  AbstractReadHandler validation, and TempResourceContainer edge cases.

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
