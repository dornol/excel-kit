# Changelog

All notable changes to this project will be documented in this file.

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
