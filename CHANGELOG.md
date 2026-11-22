# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Changed
- **Breaking:** `io.github.dornol.excelkit.core` is split by purpose (change the imports; names are unchanged):
  reading infrastructure (`AbstractReader`, `ReaderOptions`, `ReadLimits`, `CellData`, `RowData`, `ReadResult`,
  `CancellationToken`, the `Read*` exceptions and the rest) moved to `core.read`; `InputStreamSource` and the
  `TempResource*` types to `core.io`; `TabularFileDetector`, `TabularFileType`, `TabularDetectionResult` and
  `DetectionConfidence` to `core.detect`. `ExcelKitException`, `FileHandler`, `ProgressCallback`, `Cursor`,
  `AtomicFileWriter` and `RuntimeDependencyGuard` stay in `core`. The empty `core.internal` package is removed.
- **Breaking:** `io.github.dornol.excelkit.excel` is split (change the imports; names are unchanged): `ExcelReader`,
  `ExcelWorkbookReader`, `ExcelReadException`, `ExcelSheetInfo`, `SheetReadReport`, `WorkbookReadReport` and
  `SheetFailurePolicy` moved to `excel.read`; `ExcelColor`, `ExcelBorderStyle`, `ExcelDataFormat` and
  `CellColorFunction` to `excel.style`; `ExcelCellComment`, `ExcelHyperlink` and `ExcelImage` to `excel.feature`.
- `constColumn(name, value)`, `columnIf(name, condition, function)`, `maxRows(int)` and `table(...)` are declared once in the
  shared sheet-writer base instead of separately on `ExcelWriter` and `ExcelSheetWriter`. Calls compile unchanged;
  recompile code built against an earlier version.
- The reference guide and `AI.md` list every public type by package (with typical imports), and say that
  `ReadAbortException` (from `readStrict()`, `readAll()` and `maxErrors()`) does not extend `ExcelKitException`.
  `DocumentedPackagesTest` keeps both tables equal to the source.
- Every public type now has a version in its class javadoc `@since` (32 used the date `2025-07-19`, 45 had none) and
  every public method and constructor has a javadoc, mostly the `Path`/`InputStream`/`InputStreamSource` overloads of the
  readers and the Spring helpers. `ReadSecurityException.reason()` had its description attached to the wrong constructor.
  `@since` values come from the first release tag that contains the type; `ReadLimits`, `AbstractReader` and
  `HeaderStyleConfig` were corrected to the release both git and this changelog show. `PublicApiDocumentationTest`
  keeps the rule.
- `PackageDependencyTest` pins the package layering (`core` <- `core.io` <- `core.read`/`core.detect` <-
  `csv`/`excel.read`/`excel` (+ `excel.style`, `excel.feature`) <- `schema`), so `core` cannot depend on `excel` or `csv` again.

### Fixed

- `ExcelWorkbookReader.read(InputStream)` now stops copying once the input exceeds the readers' `ReadLimits.maxInputBytes` (it used to copy the whole stream first) and keeps the copy in a restricted-permission temp directory.
- A summary row's default label (`Min`, `Max`, ...) no longer depends on the JVM locale (it was `Mın` under a Turkish locale).
- `ExcelWriter.write(Stream, ...)` now also releases the workbook's temp files when an `Error` such as `OutOfMemoryError` is thrown.
- Writing to a path that already holds a file (`writeTo(Path)`) now keeps that file's permissions; the atomic replace used to give it the default ones (`rw-rw-r--` became `rw-------` under a 077 umask).
- Internal: `readWhile` now lives once in `AbstractReadHandler` (the CSV and Excel handlers each had a copy), and `ExcelReadHandler` shares one error-translation path across `read`, `readWhile` and shared-package reads. Behaviour is unchanged.


## [0.27.0] - 2026-11-21

### Added
- `ExcelReader.getSheetNames(input, password)` and `getSheetHeaders(input, sheetIndex, headerRowIndex, password)`
  (`InputStream`, `Path` and `InputStreamSource` overloads) read a password-protected workbook. A wrong password is an
  `ExcelReadException: Invalid password for encrypted Excel file`.
- `maxCellStyles(int)` on `ExcelWriter` and `ExcelSheetWriter` caps the cell styles a write may create for
  `cellColor`, `rowColor` and `rowStyle` colors (default 64,000, the .xlsx limit), so a function returning too many
  distinct colors fails fast instead of spending minutes creating styles.
- `ReadLimits.UPLOAD`, a preset for files from untrusted users (50 MiB, 50 sheets, 1,000 columns,
  32,767 characters per cell, 10 million cells). Readers still apply no limits unless given some; the
  javadoc and guides now say so.
- `CsvWriter.lineSeparator(String)` sets the row terminator (`"\r\n"`, `"\n"` or `"\r"`).

### Changed
- **Breaking:** `ExcelKitSchema` moved from `io.github.dornol.excelkit.core` to the new
  `io.github.dornol.excelkit.schema` package (change the import). `core` no longer depends on `excel` and `csv`, which
  used to depend on each other in a cycle. `ExcelWorkbookReader.sheet(name|index, schema, supplier, validator, consumer)`
  is removed for the same reason: use `sheet(name|index, schema.excelReader(supplier, validator), consumer)`.
  `ReaderOptions.toBuilder()` is public. In `excel-kit-spring`, the `ExcelKitTemplateResponse` and `ExcelKitUpload`
  overloads that take an `ExcelKitSchema` now take the one from the `schema` package (same methods, recompile).
- Rich text takes the font of its column: unformatted text and each formatted segment start from the column's font
  name, size, colour and bold, and the segment's own options are laid over it. Before, they started from POI's default
  font (Calibri 11), so a column in another font changed font inside a rich text cell.
- `ExcelTemplateWriter.cell(...)` accepts `ExcelRichText` and keeps its formatting (it used to write the object's text
  or fail). The template writer now keeps shared strings.
- A `char[]` value is written as its characters in Excel cells (including templates) and CSV, not as `[C@1b6d3586`.
- When a write rolls over to further sheets, a configured chart is drawn on every sheet over that sheet's own rows
  (it used to appear only on the last sheet and plot just the remainder).
- **Breaking:** `ExcelChartConfig.position(...)` refuses a negative or backwards column pair and two given rows that run
  upwards (`IllegalArgumentException`); a chart whose category or value column is not among the sheet's columns fails
  with `ExcelWriteException` instead of plotting an empty range.
- **Breaking:** download responses send `Cache-Control: private, max-age=N` instead of `max-age=N`, so a shared cache
  (proxy, CDN) does not keep a generated file and hand it to another user within the max-age.
- `ExcelKitMultipartFile.requireExcelOrCsvContent` checks the real signatures: an `.xlsx`/`.xlsm` upload must start with
  the ZIP signature (a text file starting with `PK` used to pass), and a `.csv` upload that is a workbook or other binary
  is refused as "not text CSV".
- The "Allowed extensions" list in the upload error message is sorted, so it reads the same on every JVM.
- **Breaking:** `CsvReader.ignoreLeadingWhiteSpace` defaults to `true`. With `false`, a hand-written file such as
  `1, "x,y"` (a space after the comma) read as `b= "x,y`: the opening quote kept and the closing one dropped, which is
  neither the text of the file nor the quoted value. Now the value is `x,y`. Whitespace around values that are not
  quoted is kept either way; `ignoreLeadingWhiteSpace(false)` is still there for files that mean a space before a
  quote literally.
- **Breaking:** `CsvReader` refuses a delimiter, quote and escape character it cannot tell apart before it reads the
  file, with `IllegalArgumentException`: a delimiter equal to the quote or the escape character, a quote equal to the
  escape character, a NUL delimiter, or any of them a line break. These used to fail inside opencsv
  (`UnsupportedOperationException` under `Failed to read CSV`), or, for a line break as delimiter, be read as nonsense.
  The delimiter the detector picks for `readDetected` is checked the same way. A quote or escape character of NUL still
  means "none". `CsvWriter.delimiter(...)` refuses the quote character, a line break and NUL, whose output could not be
  read back.
- A CSV read that fails because of what the file holds now says so in its message: `Failed to read CSV: CSV input
  contains a NUL character`, `Failed to read CSV: Unterminated quoted CSV field in the record starting at line 2 ...`,
  and, for text that is not valid in the charset, `Failed to read CSV: the input is not valid UTF-8 text. If the file
  uses another encoding (EUC-KR, windows-1252, UTF-16, ...), set it with charset(...)`. Before, the message was only
  `Failed to read CSV` and the reason sat in the cause.
- **Breaking:** an `ExcelRichText` value written into a workbook without shared strings is refused
  (`ExcelWriteException`, whose cause says to create the writer with `useSharedStrings(true)`; with
  `writeErrorPolicy(LENIENT)` it writes the plain text). Streaming writes a string cell as its plain text, so the
  formatting was lost without a word: the cell held `<is><t>Hello world</t></is>` where the code and the guide
  promised a bold word. Only a workbook created with `useSharedStrings(true)` keeps the runs.
  `ExcelRichText.FontStyle.fontSize` accepts 1–409 points and `text`/`bold`/`italic`/`styled` refuse `null`, at the call.
- **Breaking:** a page setup that Excel cannot use is refused where it is made, with `IllegalArgumentException`,
  instead of writing a broken sheet or failing after the whole stream: a margin that is negative, NaN or infinite;
  `fitToPage` pages outside 0–32,767 (40,000 used to wrap to a negative number and 65,536 to 0, and a negative count
  wrote XML the schema does not allow); `repeatRows` with a negative or reversed range or a row past 1,048,575 (a
  negative first row used to be ignored silently, the others failed with `Failed to write excel`); a header or
  footer over Excel's 255 characters (the `&L`, `&C`, `&R` marks count), checked when the setup is handed to
  `printSetup(...)`.
- `repeatHeaderRows()` and `repeatRows(first, last)` replace each other, the later call winning; `repeatHeaderRows()`
  used to win whatever the order.
- **Breaking:** data validation setup mistakes are refused where they are made, with `IllegalArgumentException`,
  instead of writing a broken validation or failing after the whole stream: `dropdown(...)` without options, with a
  `null` option, or with an option that contains a comma (Excel splits the list there; use `listFromRange(...)`);
  `listFromRange` and `formula` with blank text; `decimalBetween` with NaN or infinite bounds; bounds that no value
  could meet (`integerBetween(10, 1)`, `decimalBetween`, `textLength` with min above max or below 0, `dateRange` with
  the start after the end); an error title over 32 or a message over 255 characters; a column with both a `dropdown`
  and a `validation` (a cell has one data validation; the column's own choice is no longer joined by a
  `defaultStyle` of the other kind).
- A list written into a template (`ExcelTemplateWriter.list(...)`) keeps the column widths the template set, instead of
  overwriting every list column's width with one fitted to its content. A column given `width(n)` still uses that width,
  and a column the template left at the default width is still fitted to its content. A template without custom
  widths behaves as before.
- **Breaking:** removed `ExcelConditionalRule.dataBar(minColor, maxColor)`. It wrote a second `<color>` element
  into the data bar, which the OOXML schema does not allow (the worksheet failed schema validation), and Excel's
  data bars have no two-color gradient. Use `dataBar(color)`.
- **Breaking:** more writer arguments are rejected where they are passed (`NullPointerException` naming the
  argument) instead of failing later: `ExcelSummary` (`sum`, `average`, `min`, `max`, `count`, `label`),
  `ExcelConditionalRule` (every comparison, `dataBar`, `iconSet`), `ExcelChartConfig.type/barGrouping/barDirection`
  and the series title of `valueColumn`, `documentProperty` values, `ExcelWorkbook.headerFontName`, and the stream
  given to `ExcelSheetWriter.write` (before, a missing column list was reported first). Arguments whose `null` means
  "not set" are now annotated `@Nullable`: the optional column `configurer`, style alignments, borders, formats and
  fonts of `ColumnConfig`, `HeaderStyleConfig`, `RowStyleConfig`, `ExcelPrintSetup` headers/footers/orientation/paper,
  chart titles and legend position, and `ExcelReader.password`.
- **Breaking:** the writers reject a null argument where it is passed, with a `NullPointerException` naming it,
  instead of failing on the first row (or not at all): the callbacks and functions of `beforeHeader`,
  `afterData`, `afterAll`, `rowColor`, `rowStyle`, `onProgress`, `sheetName`, `column`/`columnIf` (every
  overload, including the `Function` ones whose null used to hide inside a lambda), `groupComment`, plus
  `headerFontName`, `CsvWriter.charset/quoting`, and a null `Stream` given to `write`. `protectSheet` and
  `protectWorkbook` reject a null or blank password with `IllegalArgumentException`, like `password(...)`;
  a null password used to leave the sheet or workbook silently unprotected.
- **Breaking:** `ExcelSheetWriter` (sheets of an `ExcelWorkbook`) rolls over after 1,000,000 data rows by
  default, like `ExcelWriter`. Without `maxRows(...)` it never rolled over, so a sheet fed more than
  1,048,575 rows failed with POI's `Invalid row number (1048576) outside allowable range` once the
  limit was reached. Sheets with fewer rows are unchanged.
- **Breaking:** `maxRows(...)` of `ExcelWriter` and `ExcelSheetWriter` rejects a value above 1,048,575, the
  most data rows an Excel sheet can hold below its header row.
- **Breaking:** `ExcelKitUpload.excel/csv(file, schema, supplier[, validator])` without explicit limits are
  now bounded by `ReadLimits.UPLOAD` (50 MiB, 50 sheets, 1,000 columns, 32,767 characters per cell, 10 million
  cells); they used to read without limits. Pass `ReadLimits.UNLIMITED` to the overload that takes limits to
  keep the old behavior. Uploads read through an `UploadReader` are unchanged.
- **Breaking:** `CsvWriter` ends rows with CRLF as RFC 4180 requires, on every operating system. It used
  the operating system's line separator, so the same export was CRLF on Windows and LF elsewhere.
  `CsvDialect` now carries a line separator (`getLineSeparator()`): `RFC4180` and `EXCEL` use CRLF,
  `TSV` and `PIPE` use LF. Call `lineSeparator("\n")` for LF output.

### Fixed
- Exceptions that are converted into another one keep what went wrong as their cause: the `IllegalArgumentException`
  of `CellData.asBigDecimal()` carries the `NumberFormatException`, and the `ReadSecurityException` raised when a
  worksheet cannot be scanned for formulas carries the scan failure (new constructor
  `ReadSecurityException(Reason, String, Throwable)`).
- A `*` in a download filename is escaped as `%2A` in the `filename*` parameter (RFC 8187 does not allow it as is).
- An emoji (a surrogate pair) split between two `ExcelRichText` segments was stored as two `?`; the run boundary now
  moves behind the whole character.
- A `dropdown(...)` option containing a double quote wrote a broken list formula (`"say "hi",c"`); quotes are now
  written twice, as Excel's formula syntax asks, and count twice towards the 255-character limit. A single-bound
  validation (`integerGreaterThan`, `integerLessThan`) no longer writes an empty `<formula2/>`.
- `ExcelTemplateWriter.cell(...)` with a date or time in a cell that has no number format (any cell below the template)
  showed it as a serial number (`45356`). Such a cell now gets `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss` or `hh:mm:ss`, keeping
  its font and borders; a cell whose template gave it a format keeps that one.
- A `list(startRow)` on a sheet whose template has rows at or below `startRow` failed deep in POI with `Attempting to
  write a row[3] in the range [0,6] that is already written to disk`. `list(...)` now says what is wrong: the list must
  start below the template's last row, and which row that is.
- A `cell(...)` or `list(...)` that was refused (a row or column off the sheet, a text over 32,767 characters) still
  moved the writer's row position, so the next valid write could fail with `Cells must be written in top-down row
  order`. A refused write now leaves the position as it was. A row or column off the sheet is an
  `IllegalArgumentException` at the call, also for `list(...)`.
- A `summary` on a template list written with `write(...)` (the template already has the headers) left out the first
  data row: `SUM(B5:B6)` for data in rows 4 to 6. The list took its own start row for the header row; the header is the
  template's row above it. `afterData` was told the wrong header row too (`SheetContext.getHeaderRowIndex()`).
  `writeWithHeaders(...)` was not affected.
- `ExcelWorkbookReader` can read a password-protected workbook. It listed the sheets without decrypting the file,
  so every read failed with `Failed to read sheet names` (`OLE2NotOfficeXmlFileException`), however the sheet readers
  were set up with `password(...)`. The sheets are now listed from a decrypted copy (deleted afterwards), and a wrong
  password is reported as `Invalid password for encrypted Excel file`.
- A CSV read with `skipBlankRows` (or `stopAtBlankRows`) counted the skipped blank rows in its progress: `onProgress`
  and `onReadProgress` reported more processed rows than were delivered, and the last `onReadProgress` event
  (`completed`) then reported fewer. Skipped rows are no longer counted, as in the Excel reader.
- With a multi-row (grouped) header the writers took the first header row for "the header row", so everything
  placed relative to it was one or more rows too high: a `summary` range, a structured `table`, a
  `conditionalFormatting` range, dropdown and `validation` ranges, a `chart`'s series, the rows repeated by
  `printSetup(p -> p.repeatHeaderRows())` and named ranges all included the column-name row, and a table started on
  the group row (which Excel may repair on open). `SheetContext.getHeaderRowIndex()` is now the row with the
  column names, the last row of the header.
- A `summary` ends at the last data row. When `afterData` wrote rows too, the range reached down through them (and
  double counted any numbers there).
- `namedRange(...)` with a rollover no longer fails with `Failed to write excel` (or POI's "workbook already
  contains this name"): the first sheet keeps the name, later sheets get `name_2`, `name_3`, ..., as tables do.
- A write that needs more than 64,000 distinct cell styles (every distinct `cellColor`, `rowColor` or `rowStyle`
  color needs one) fails with `ExcelWriteException: Too many distinct cell styles: an .xlsx workbook holds at most
  64000 …` instead of a generic `Failed to write excel`. The javadoc and the styling guide now say that the time
  grows with the square of the number of distinct colors (1,000 ≈ 1s, 4,000 ≈ 10s, 8,000 ≈ 40s), so a continuous
  gradient should be rounded to a palette.
- A structured table with `perRolloverSheet` off now fails when the first rollover is due, with a message that
  names the table, instead of after the rest of the stream has been written.
- A sheet that runs out of rows (`maxRows` set too high for the rows `beforeHeader`, the header, `afterData`
  or a summary also use) now fails with `ExcelWriteException: Sheet 'S' is full: Excel sheets hold at most
  1048576 rows. Lower maxRows(...) …` instead of POI's bare `Invalid row number` (or a generic
  `Failed to write excel`).
- The RGB overloads `tabColor`, `backgroundColor`, `fontColor`, `headerFontColor`, `headerBackgroundColor`
  (and `ExcelRichText`'s `color`) now reject a component outside 0–255 with `IllegalArgumentException`, like
  `ExcelColor` and `RowStyleConfig` already did. Before, `fontColor(300, 0, 0)` silently became red 44.
- `autoWidthSampleRows(n)` samples the first `n` data rows. It counted sheet rows instead, so rows written
  by `beforeHeader` (and the header) used up the sample: with a long enough preamble no data row was
  sampled and the column kept the header's width. The sample is also taken once from the start of the write,
  not again on every rollover sheet.
- `rowStyle(...)` no longer drops the column's RGB font color and underline (and charset, type offset and
  font family) when it changes only some font attributes such as `bold`. It copied the font name, size,
  bold, italic, strikeout and indexed color only, so a red underlined column turned plain after a row
  style that merely set `bold(true)`.
- The javadoc of `headerPolicy`, `cancellationToken`, `onReadProgress`, `HeaderPolicy`, `CancellationToken`,
  `ReadProgressCallback` and every `ReaderOptions.Builder` setting now says what it does.
- The javadoc of every `write(Stream)` now says that the stream is read but not closed, and the writing
  guide shows try-with-resources for JPA/JDBC streams.
- The external-link check of `ReadSecurityPolicy` (`allowExternalLinks(false)`, including `STRICT`) reads
  relationship parts in 8 KiB chunks instead of byte by byte. The verdicts are unchanged.
- `CsvQuoting.NON_NUMERIC` and the CSV injection defense now share one definition of a number (an optional
  sign, ASCII digits with an optional fraction, an optional exponent). Before, `NON_NUMERIC` quoted `1e5`,
  left a lone `.` unquoted and treated non-ASCII digits as numbers.
- `CsvWriter` formats a row without per-field and per-row stream allocations: about 25–30% faster on the
  1M-row × 5-column benchmark (0.61s → 0.44s on the development machine). Output is unchanged.
- A CSV write that fails now says where. A character the charset cannot encode throws
  `CsvWriteException: Cannot write CSV column 'name' at row 5: character 'é' (U+00E9) cannot be encoded in US-ASCII`
  (or `header of column '…'`); before, the writer only said `Could not encode or write CSV` after it had
  read the whole stream. An I/O failure on the temp file now stops reading rows at the next row instead of
  at the end, and the message carries the underlying cause.
- `reader.options(ReaderOptions)` no longer drops the Bean Validation provider given to
  `setter(...)`/`mapping(...)` (or `validator(...)`) when the options carry none; before, rows
  that violated constraints silently passed validation. Options that do carry a validator still override it.

## [0.26.0] - 2026-11-21

### Added
- `ExcelHandler.discard()` and `CsvHandler.discard()` release a handler that will not be written
  (idempotent, a no-op after `writeTo`). A handler that is garbage collected without being
  consumed now releases its temp files as a safety net, at a time the garbage collector chooses.
- `kit` and `excel-kit-spring` jars declare `Automatic-Module-Name` (`io.github.dornol.excelkit`,
  `io.github.dornol.excelkit.spring`) so module-path users get stable module names.

### Changed
- **Breaking:** a column whose header is blank is no longer a `""` key in `forMap()` rows of CSV and
  Excel readers. It has no name to look it up by, so it is left out of the header index; before, the
  first blank column became the `""` key and every further blank column's values were silently lost.
  Columns read by position are unaffected, and a `headerNormalizer` that names blank headers still
  keeps those columns reachable by that name.
- **Breaking:** removed the process-wide `CellData` conversion defaults: `setDefaultLocale`,
  `getDefaultLocale`, `addDateFormat`, `addDateTimeFormat`, `getDateFormats`,
  `getDateTimeFormats`, `resetDateFormats` and `resetDateTimeFormats`. Use
  `reader.cellConversion(c -> c.locale(...).addDateFormat(...))` instead. A `CellData` without a
  config, and a reader without `cellConversion(...)`, use the immutable
  `CellConversionConfig.defaults()` (JVM default locale, built-in date patterns).
- **Breaking:** `ExcelKitUpload.excel/csv(file, schema, supplier[, validator])` now check the content signature
  like the `ReadLimits` overloads, so a file whose content does not match throws
  `ExcelKitUploadException` instead of a reader-specific failure.
- A read that fails because of what is in the file now says what in its message, not only in the cause:
  `Failed to read CSV: Header aliases [Price] not found in CSV. Available headers: [Name, Cost]`
  (the same for Excel, and for duplicate or unusable headers and a missing header row). Before, the
  message was just `Failed to read CSV` / `Failed to read excel`, so an API that reports `getMessage()`
  gave no clue. Other causes keep the generic message.
- A CSV read that fails because `strictQuotes(true)` blanked an unquoted header row now says so
  (`a header cell is blank … every header cell must be quoted`) instead of only
  `Failed to read CSV`; the original cause is still attached. The javadoc of `strictQuotes` and
  `ignoreLeadingWhiteSpace` now describes what they do.
- `CsvHandler.close()` now behaves like `discard()`: it also marks the handler consumed, so a later
  `writeTo` throws `CsvWriteException` instead of failing on the deleted file.
- `ExcelKitResponse.excel(handler, filename, password)` rejects a null or blank password when the
  response is built (`IllegalArgumentException`) instead of while the body streams, and releases the
  handler's workbook.

### Fixed
- Excel writers delete the temp files behind a workbook's sheets on every supported Apache POI
  version. POI 5.2.5's `SXSSFWorkbook.close()` does not do this, so each export used to leave one
  temp file per sheet in `java.io.tmpdir/poifiles`; the workbook is now also disposed explicitly.
- A reader whose configuration is rejected while it opens its input (for example a negative
  `headerRowIndex`) deletes the temp copy of the input instead of leaving it behind.
- `ExcelWriter.write` releases the workbook's temp files when writing the first sheet's
  preamble or header fails (for example a throwing `beforeHeader` callback); the original
  exception is still thrown unchanged.
- `ExcelHandler.writeTo(..., password)` with a null or blank password now consumes the handler and
  closes its workbook before throwing `IllegalArgumentException`, so a rejected password no longer
  leaves the workbook's temp files behind. A conflict with a writer-level `password(...)` still
  throws `IllegalStateException` and leaves the handler usable with `writeTo(OutputStream)`.

## [0.25.0] - 2026-11-20

### Added
- `ExcelReader.forMapRaw()` / `forMapRaw(String...)` return stored XLSX numbers as plain
  strings instead of displayed text.
- `CellData.rawNumber()` reads the stored XLSX number (including cached formula results,
  date serials, and percentage fractions); `rawValues()` on read results and errors use it.
- `CellData.asZonedDateTimeStrict(...)` rejects local times in a gap or overlap.
- CSV delimiter detection ignores quoted fields and honors `headerRowIndex(...)`; explicit
  charset, delimiter, and dialect settings take precedence.
- `ZonedDateTime`, `OffsetDateTime`, `java.util.Date`, `java.sql.Timestamp`, `java.sql.Date` and
  `java.sql.Time` are written as Excel dates by inference, typed `DATETIME`/`DATE`/`TIME` columns,
  and template cells. `ZonedDateTime` and `OffsetDateTime` keep their own wall-clock time; the
  others use the JVM default time zone. `Instant` is still written as text because its zone is
  ambiguous.
- XLSX text cells preserve literal `_xHHHH_` text and control characters, and ISO
  date, date-time, and time cells are read across POI versions.
- `AtomicFileWriter` writes a file through a temporary sibling and replaces the target only on
  success; `writeTo(Path)` and the password `writeTo(Path, ...)` overloads use it.

### Changed
- **Breaking:** the default `ExcelWriteErrorPolicy` is now `FAIL_FAST`. A failing value
  extractor or a value that cannot be stored as the column's type throws
  `ExcelWriteException`; use `writeErrorPolicy(LENIENT)` for blank cells and text fallback.
  An explicit `nullValue(...)` is always written, as text when it does not fit the column type.
- **Breaking:** the default `DuplicateHeaderPolicy` is now `FAIL` (blank headers are ignored).
  Use `FIRST` or `LAST` to keep the previous resolution.
- **Breaking:** for XLSX numeric cells, `asNumber/asInt/asLong/asDouble/asFloat/asBigDecimal`
  convert the stored number instead of the displayed text (`2.5` in a `0` cell stays `2.5`,
  `0.25` in a `0%` cell stays `0.25`). Date-formatted cells convert to their serial number.
- The error for an unterminated quoted CSV field names the line where the record starts and how
  to escape the quote.
- `BIG_DECIMAL_TO_LONG` accepts any `Number`, `DATETIME` accepts a `LocalDate` (as midnight), and
  a rejected typed value now reports its value type and column type in the error message.
- `forMap()` on CSV and Excel returns `""` for missing trailing columns instead of omitting
  the key.
- Summary rows throw `ExcelWriteException` for an unknown summary or label column, and write
  no formulas when the sheet has no data rows.
- `dropdown(...)` throws `IllegalArgumentException` when the comma-joined options exceed
  Excel's 255-character limit.
- Date-times and times with sub-millisecond precision, dates outside Excel's date system,
  and numbers that Excel cannot store exactly are rejected (`FAIL_FAST`) or written as text
  (`LENIENT`) instead of being rounded.
- Updated hibernate-validator and the Gradle wrapper to 9.9.1.

### Fixed
- Conditional formatting (rules, data bars, icon sets) stops at each sheet's last data row, so
  summary and `afterData` rows are no longer formatted as data or used to scale data bars.
- `ExcelWorkbookReader` reports a missing named sheet as a per-sheet failure that follows the
  failure policy, like a missing index, instead of aborting the whole read; `SheetReadReport`
  uses index `-1` for a sheet that was not found.
- Charts are placed below summary and `afterData` rows instead of covering them, and are
  skipped when a sheet has no data rows instead of failing the write with an invalid range.
- Group headers no longer merge equal inner labels across different parent groups.
- Rolled-over sheet names keep their ` (n)` suffix within Excel's 31-character limit, so long
  base names no longer produce truncated or duplicate names.
- `writeTo(Path)` no longer destroys an existing file when the write fails or the handler was
  already consumed, and no longer leaves a partial file behind.
- `ReadSecurityPolicy` formula detection can no longer be bypassed with a namespace prefix, an
  upper-case part path, or a worksheet stored outside `xl/worksheets/`; the preflight scan is
  also about 35 times faster on large workbooks.
- Upload content checks accept valid UTF-8 CSV files whose 4096-byte inspection prefix ends
  inside a multi-byte character.
- `defaultStyle(...)` now applies every column property (widths, header colors, cell colors,
  hidden, dropdown, validation, comments) unless a column overrides it.
- `UploadError` reports setter conversion failures as `MAPPING` instead of `VALIDATION`.
- `TIME` columns no longer silently truncate sub-millisecond values.
- `Float` values are stored by their decimal value (`0.1f` is `0.1`, not `0.10000000149011612`).
- Named ranges work for sheet names containing an apostrophe.
- `maxRows(1)` starts a new sheet for every row.
- XLSX setter columns receive an empty cell for missing trailing cells, like blank cells and CSV.
- `countRows()` totals use row numbers, so leading empty rows no longer skew them.
- CSV values, line endings inside quoted fields, numeric and time conversions, and
  scientific notation are parsed without silent changes; invalid CSV data is rejected.
- Excel numeric and temporal conversions no longer silently change values (1900 leap day,
  1904 date system, time zones, formula results, cell positions without references).

## [0.24.0] - 2026-11-02

### Added
- Same-runner benchmark repetitions and environment-specific advisory noise bands.
- Custom cache-duration overloads for Spring download and error responses.
- Frozen external Excel/Mac Excel/LibreOffice workbook fixtures with pinned upstream
  sources, checksums, producer metadata and offline compatibility tests.
- Advisory benchmark comparisons against compatible successful main runs, with
  percentage changes, aggregate increase warnings, and persisted comparison reports.
- Sampled heap and temporary-file resource peaks for all benchmark scenarios, with
  an isolated scratch directory and machine-readable CSV reports uploaded by CI.
- Independent Excel-only, CSV-only, Bean Validation, and Spring consumer smoke tests
  against published artifacts using Gradle metadata and Maven POM resolution.
- Release tag/version validation and post-publication Maven Central consumer checks.

### Changed
- The default tested POI version is 6.0.0; the same compiled library is also tested
  against POI 5.2.5 and 5.5.1 without forcing a runtime version on applications.
- Updated SLF4J to 2.0.20, Hibernate Validator to 9.1.5.Final, the example Spring Boot
  version to 4.1.2, and the Gradle wrapper to 9.8.2.
- Public API compatibility is checked against 0.23.0 without legacy API exclusions.

### Fixed
- Retained the published `ExcelKitUpload.validateExcel` and `validateCsv` aliases,
  deprecated in favor of `excel` and `csv`, preserving 0.23.0 binary compatibility.
- Streaming cell and header formatting now respects the workbook's 1904 date system.
- Path-based reads and metadata discovery open OOXML packages read-only instead of
  rewriting caller-owned workbook bytes on close.
- Encrypted Path reads allocate an owned decryption workspace, clean plaintext on
  preflight/parse failures, and preserve the encrypted source.
- Streaming sheets now resolve their backing XSSF sheet through public workbook APIs,
  removing reflection on POI's private `_sh` field.
- Configured sheet feature failures propagate as write errors instead of silently
  omitting tab colors, tables, data bars, icon sets or chart data labels.
- Structured tables preserve actual column headers after streaming rows are flushed,
  including rollover sheets and existing or generated template headers.
- Bounded reader reads/skips now request only the remaining input-byte budget plus
  one overflow-detection byte, stop consuming after failure, and avoid count overflow.
- Post-copy byte-limit errors retain their observed size and type after temp-file cleanup.
- Custom compression-ratio limits reject NaN and infinities.
- CI test JVMs now follow the JDK matrix instead of always using the JDK 21 compiler toolchain.
- Java compilation targets the Java 17 API surface with `--release 17`.
- Spring Maven POM integration dependencies are optional, matching the documented installation policy.
- Both library modules are uploaded in one Central deployment; GitHub Release creation waits
  for publication visibility and consumer verification.

## [0.23.0] - 2026-08-13

### Added
- Fluent `validator(Validator)` configuration for Excel and CSV readers.
- Explicit `ReadResult.errorType()` classification for mapping and validation failures.
- Shared Excel/CSV reader contract tests covering limits, cancellation, early stop, and resource ownership.
- Writer cleanup contract tests covering iterator and output stream failures.
- Benchmark coverage for Excel `Path` input, `countRows()`, and multi-sheet package sharing.
- Actionable runtime dependency diagnostics for missing POI or OpenCSV integrations.
- ZIP entry count protection through `ReadSecurityPolicy.maxZipEntries()`.

### Changed
- `ReadResult.messages()` is now always an immutable, non-null list.
- Setter conversion failures are reported as `RowError.Type.MAPPING`, while required-field
  and Bean Validation failures remain `RowError.Type.VALIDATION`.
- Reader callbacks and configuration functions now reject `null` arguments immediately.
- Schema reader/writer construction is delegated to dedicated internal factories.
- `ReadResult` creation now uses explicit `success`, `failure`, `mappingFailure`, and
  `validationFailure` factories; legacy overloaded constructors were removed.
- Callback, writer, and cleanup exception boundaries now catch only the checked or runtime
  exception categories they can actually handle.
- API compatibility CI now compares against `0.22.0` and explicitly excludes only the
  intentionally redesigned reader/result members until the next major release.
- Common null-check expressions now use consistent `Objects` imports across reader and writer
  implementations.
- Excel and CSV handler creation now shares the common input stream and path validation helpers.
- Reader handler creation and execution now uses shared internal stream/path execution helpers.
- Reader execution aggregation is now isolated from fluent reader configuration in an internal
  support component.
- Mutable format-independent reader settings are now isolated in an internal configuration object;
  execution receives an immutable `ReadOptions` snapshot.
- Read handlers now retain the immutable `ReadOptions` snapshot directly instead of copying its
  values into mutable execution fields.
- Reader validator configuration is now owned by the same internal configuration object as the
  other common reader settings.
- Legacy and detailed reader progress callbacks now keep independent intervals.
- CSV and Excel readers now share row-policy handling for cancellation, blank rows, and row limits.
- Internal row-policy implementation is no longer exposed as public API.
- Legacy progress settings now travel through the immutable reader options snapshot alongside
  detailed progress settings.
- Map-mode selection and column collections are encapsulated and snapshotted when a session starts.
- Legacy progress callback dispatch is now shared by CSV and Excel read handlers.
- Added reusable `ReaderOptions`, null-safe row/cell accessors, CSV detection source/summary APIs,
  configurable detection sampling, and `readAll(...)` convenience methods.
- Added `ExcelWorkbookReader` with per-sheet readers/schemas, sheet-level reports, bounded errors,
  and continue/stop failure policies.
- Added an optional cumulative `ReadLimits.maxTotalCells` guard for untrusted CSV/Excel input.
- Added explicit `ReadSession<T>` lifecycle APIs and schema mapping overloads without validators.
- Multi-sheet Excel reads now reuse one `OPCPackage` for non-encrypted readers; password readers
  retain the isolated fallback path.
- `maxColumns` now applies to data rows as well as headers.
- STRICT security scanning now detects external workbook links declared in worksheet relationship files.
- Detected CSV charset and delimiter settings now apply only to the current read operation.
- Dependency compatibility CI now verifies the minimum POI/OpenCSV combination on JDK 17.

## [0.22.0] - 2026-08-06

### Added
- Automatic Excel column type inference for non-null values in `ExcelWriter` and
  `ExcelSheetWriter`, including numeric, date/time, hyperlink, image, and rich-text values.
- Automatic default number/date formats for inferred numeric and date/time columns.
- Comprehensive regression coverage for inferred types, null-leading columns, explicit type
  overrides, custom formats, and both column writer APIs.

### Changed
- Columns without an explicit `.type(...)` now infer their writer type from the first non-null
  value. Explicit types remain authoritative.
- Columns whose values are all `null` retain the `STRING` fallback behavior.

## [0.21.0] - 2026-07-11

### Added
- Symmetric `Path` and `InputStreamSource` overloads for `read`, `readStrict`, and `readWhile`.
- `Path` and `InputStreamSource` overloads for Excel sheet metadata discovery.
- `ReadSummary`, bounded `ReadReport`, detailed `ReadProgress`, and cooperative cancellation APIs.
- Header normalization presets through `HeaderPolicy`.
- Defensive input, sheet, column, and cell-size limits through `ReadLimits`.
- Signature-based `TabularFileDetector` for XLSX, XLS, CSV, and unknown inputs.
- Excel structured table creation through `ExcelWriter.table(...)`.
- SXSSF temporary-file compression and shared-string initialization options.
- CI-generated japicmp reports against the latest published baseline.
- Structured `ReadLimitExceededException` details and pre-copy stream byte enforcement.
- Detailed file detection confidence, charset, delimiter, and read-support metadata.
- Shared immutable `StreamingOptions` and rollover/multi-sheet structured tables.
- Signature-validated Spring uploads with core `ReadLimits` integration.
- Optional Excel formula and external-link rejection through `ReadSecurityPolicy`.
- Structured tables for template list writers and `Iterable` template inputs.
- Bounded XLSX security scanning with entry, total-byte, and compression-ratio limits.
- Spring upload result collection limits with independent truncation indicators.
- CSV automatic charset/delimiter detection and deterministic malformed-input fuzz coverage.

### Changed
- Excel metadata methods no longer close caller-provided `InputStream` instances.
- Reader executions now snapshot column and selected-header collections.
- Reader `Path` inputs are read directly without copying the source file to a temporary file.
- Excel and CSV read-handler implementations are no longer public API types.
- Exceptions thrown by a `readWhile` predicate are propagated unchanged.
- `readWhile` returns `ReadSummary`; summary/report methods now support all input source types.
- Detailed read progress always emits a terminal completion/cancellation event.
- API compatibility CI now fails for non-allowlisted binary incompatibilities.
- Reader summary/report aggregation is shared internally to keep Excel and CSV behavior aligned.
- japicmp exceptions are signature-scoped and document their baseline expiry.

## [0.20.0] - 2026-07-11

### Added
- Direct `read(input, ...)`, `readWhile(...)`, `Path`, and `InputStreamSource` APIs for Excel and CSV readers.
- Reader `maxErrors(...)` guards with structured `ReadAbortReason` details.
- Excel/CSV header normalization via `headerNormalizer(UnaryOperator<String>)`.
- `Iterable` inputs for Excel and CSV writers.
- Spring `UploadReader` callback integration without exposing read handlers.
- CI compatibility coverage for minimum-tested and current POI/OpenCSV versions.

### Changed
- Caller-provided input streams and writer streams are now caller-owned and are not closed by excel-kit.
- Reader executions use immutable common option snapshots.
- Excel stream reading no longer uses a producer thread and blocking queue.

### Removed
- `readAsStream()` in favor of callback-based `read(...)` and `readWhile(...)`.
- Handler-based `build(input)` in favor of direct Reader execution methods.

## [0.19.0] - 2026-07-01

### Added
- Reader-scoped `CellConversionConfig` via `cellConversion(...)` on Excel and
  CSV readers for locale/date parsing without mutating global `CellData` state.
- `ExcelWriteErrorPolicy` with `writeErrorPolicy(FAIL_FAST)` for aborting Excel
  writes on value extraction or cell type write failures.
- CSV reader parser options for quote character, escape character, strict quotes,
  and leading whitespace handling.
- Schema convenience APIs for Excel writer init options and validator-free
  schema readers.
- Schema-level reader/writer defaults for strict headers, duplicate header
  policy, cell conversion, row limits, blank-row handling, and Excel write
  error policy.
- Spring upload shortcuts for schema-backed Excel/CSV uploads.
- Spring upload content validation helper `requireExcelOrCsvContent(...)`.
- Reader row guards via `maxRows(...)`, `skipBlankRows(...)`, and
  `stopAtBlankRows(...)`.
- Raw source row values on read errors, upload errors, and downloadable error
  reports.
- Upload summary metadata with row counts, duration, filename, and file size.
- Guided Excel template response helper with required-column guidance.

### Changed
- `CellData` number cleanup now respects the configured locale's grouping and
  decimal separators.

## [0.18.4] - 2026-06-30

### Added
- Release checklist guide for version/tag/publish verification workflow.
- `excel-kit-spring` module with Spring MVC download response helpers and
  `MultipartFile` upload stream adapter.
- Spring upload result DTOs, read-error report responses, and schema-based
  template response helpers.
- `ExcelKitUpload` helper and sample-row overloads for schema template responses.
- Spring upload validation helpers for max file size and Excel/CSV filename
  extension checks.

### Changed
- Example app now uses the new `excel-kit-spring` module for download responses.
- Example read showcase now uses Spring module upload/result/error-report helpers.
- Example schema template downloads now include sample rows.
- Maven publish workflow now publishes and verifies `excel-kit-spring`.
- Example write showcase endpoints are split across focused controllers.
- Test dependency baseline updated to JUnit 6.1.1.
- Spring download filenames are sanitized before being written to
  `Content-Disposition` headers.
- `ExcelImage.fromUrl` now accepts only HTTP(S), applies a 10 MiB default
  download limit, and offers a `fromUrl(url, maxBytes)` overload.

### Tests
- Added direct coverage for example download content-disposition headers.
- Added Spring module tests for upload results, error reports, and template responses.
- Added security-focused tests for download filename sanitization and bounded
  URL image downloads.

## [0.18.3] - 2026-06-29

### Added
- Manual Maven Central visibility verification workflow backed by a reusable
  `scripts/verify-maven-central.sh` script.
- README installation examples now show the runtime dependencies that
  applications should declare because library integrations remain `compileOnly`.

### Changed
- Maven publish workflow now reuses the same Maven Central visibility script as
  manual verification.
- Example app static CSS/JavaScript is split out of `index.html`.
- Example app read report rendering is separated from the read showcase
  controller.
- Example app download response helpers are unified under `DownloadResponse`.
- Example app name-based upload area now links directly to Excel and CSV sample
  templates for correction workflows.
- Release and publish workflows now use JDK 21 to match the project toolchain.
- CI workflow permissions now explicitly include read-only repository contents
  access.
- Documentation examples now avoid `System.out` in favor of logger-style
  snippets, and data type docs use concrete `ExcelDataType` examples.

### Tests
- Example smoke tests now verify served static assets and inspect the generated
  XLSX read-error report contents.

## [0.18.2] - 2026-06-24

### Added
- Example app upload results now render inline with success rows, structured
  cell errors, showcase filtering, and client-side error CSV download.
- Example app now provides server-generated read error reports as CSV and Excel
  downloads.
- Maven publish workflow now verifies that a tagged release becomes visible on
  repo1.maven.org.
- README and Spring guide now document structured read errors and downloadable
  correction-report patterns.
- Schema reader tests now cover Excel/CSV alias, required-column, and duplicate
  header parity in a dedicated read-options test class.

## [0.18.1] - 2026-06-24

### Changed
- Maven Central publishing now uses the automatic release path so tag publishes
  do not remain as user-managed Central Portal deployments.
- Generated Javadocs now suppress missing-doc comments while keeping other
  doclint checks active.
- Showcase upload read failures now include structured cell error details in the
  response.

### Added
- Smoke coverage for CSV upload conversion errors in the example app.
- `CellError` handling examples in the reading guide and bundled docs.

## [0.18.0] - 2026-06-24

### Added
- Header aliases for Excel and CSV readers via `column(List<String>, ...)`.
- Strict reader header validation via `strictHeaders()` / `requireHeaders()`.
- Duplicate header handling via `DuplicateHeaderPolicy` (`FIRST`, `LAST`, `FAIL`).
- Physical source row tracking via `ReadResult.fileRowNum()` and `RowError.fileRowNum()`.
- Structured cell-level read errors via `CellError` and `RowError.cellErrors()`.
- Schema read aliases and required columns via `ExcelKitSchema.Builder.column(name, aliases, ...)`
  and `requiredColumn(...)`.
- Strict selected-column validation for `ExcelReader.forMap(...)` and `CsvReader.forMap(...)`.
- Example app MockMvc upload smoke tests for schema-based Excel and CSV reads.

### Changed
- Map-mode readers now apply the configured duplicate header policy consistently
  when building output maps.

## [0.17.1] - 2026-06-24

### Added
- Example app smoke test coverage now starts the Spring context without Docker
  and verifies a real MyBatis mapper query against an H2 database.

### Changed
- CSV and Excel read handlers are now explicitly one-shot. Calling `read()`,
  `readStrict()`, or `readAsStream()` consumes the handler; a second read attempt
  now fails with a clear `ExcelKitException` instead of a lower-level temp-file
  error.

### Fixed
- CSV `readAsStream()` cleanup coverage now verifies temporary resources are
  removed after failed row mapping results are consumed.

## [0.17.0] - 2026-06-22

### Added
- `ExcelWriter.create(opts -> opts.rowAccessWindowSize(...))` and
  `ExcelWorkbook.create(opts -> opts.rowAccessWindowSize(...))` for configuring
  SXSSF's row access window at workbook construction time.

### Fixed
- `ExcelReadHandler.readAsStream()` now releases temporary reader resources when
  the stream is fully consumed, even if callers do not explicitly close the stream.
- `ExcelReader.getSheetHeaders(...)` now throws `ExcelReadException` when the
  requested sheet index does not exist, matching the main read path.
- `ExcelReader.sheetIndex(...)` now accepts any non-negative index and lets the
  workbook contents determine whether that sheet exists.
- `password(char[])` now rejects blank passwords, matching `password(String)`.

## [0.16.15] - 2026-04-19

### Added
- **`countRows()`** on `ExcelReader` — opt-in pre-scan that counts data rows before
  parsing, making percentage-based progress reporting possible (e.g., via SSE).
  The total is available via `Cursor.getTotalRows()` in the `ProgressCallback`.
  ```java
  ExcelReader.setter(MyDto::new)
      .countRows()
      .onProgress(500, (processed, cursor) -> {
          int percent = (int) (processed * 100 / cursor.getTotalRows());
      })
      .build(inputStream).read(result -> { ... });
  ```
- **`rowStyle(predicate, config)`** — conditional row-level styling on `ExcelWriter`,
  `ExcelSheetWriter`, and `TemplateListWriter`. Unlike `rowColor()` (background only),
  supports bold, italic, strikethrough, font size, font color, and background color.
  Multiple rules can be registered; the first matching predicate wins.
  ```java
  writer.rowStyle(p -> p.price() > 10000,
      style -> style.bold(true).backgroundColor(ExcelColor.LIGHT_YELLOW));
  ```
- **`ExcelImage.size(width, height)`** — control image cell span (columns × rows).
  Default remains 1×1. Previously images were always anchored to a single cell.
- **`ExcelImage.fromUrl(url)`** — download an image from a URL and embed it directly.
  Auto-detects PNG/JPEG from URL extension or magic bytes. 10-second timeout.
  ```java
  .column("Photo", user -> ExcelImage.fromUrl(user.getPhotoUrl()),
      c -> c.type(ExcelDataType.IMAGE))
  ```
- **`Cursor.getTotalRows()`** — returns the total row count when known (via
  `countRows()`), or `-1` if unknown.

### Changed
- **`ExcelImage`** refactored from `record` to `class` to support `size()` chaining
  and `fromUrl()` factory. The constructor signature and `data()`/`imageType()` accessors
  are unchanged — existing code compiles without modification.

## [0.16.14] - 2026-04-19

### Added
- **`documentProperty(key, value)`** — set Excel document metadata (title, author,
  subject, keywords, description, category) on `ExcelWriter` and `ExcelWorkbook`.
  Standard keys map to core properties; other keys become custom properties.
- **`namedRange(name, columnIndex)`** — fluent named range registration on any writer.
  Automatically creates a workbook-scoped named range covering all data rows in the
  specified column, replacing manual `afterData` callback usage.
- **`headerStyle(cfg -> ...)`** — configure header cell alignment, bold, border style,
  and wrap text via `HeaderStyleConfig`. Available on `ExcelWriter` and `ExcelWorkbook`.
- **`password(char[])`** — char-array overload for `ExcelWriter.password()` and
  `ExcelWorkbook.password()`. The array is copied internally and zeroed after encryption,
  preventing password from lingering in heap as an immutable String.

### Changed
- **Internal: `ExcelColumn` constructor** simplified from 22 positional parameters to
  `(name, function, style, setter, ColumnStyleConfig)`. Adding new column fields now
  requires changes in only 2 places instead of 5+.
- **Internal: `AbstractSheetWriter` base class** extracted from `ExcelWriter` and
  `ExcelSheetWriter`, eliminating 22 duplicated sheet-config methods (~320 lines removed).
- **Internal: `AbstractReader` base class** extracted from `ExcelReader` and `CsvReader`,
  eliminating 10 duplicated column-registration methods (~177 lines removed).
- **Internal: `writeGroupAndColumnHeaders`** split into 5 focused helpers
  (`createHeaderRows`, `buildGroupGrid`, `populateHeaderCells`,
  `applyHorizontalMerges`, `applyVerticalMerges`).
- **Internal: `ExcelHandler`** stores password as `char[]` internally (zeroed after use).

### Documentation
- **Guide restructured** — `docs/guide.md` (2219 lines) split into `docs/guide/` directory
  with 13 topic files. `index.md` contains inline code examples covering 80% of use cases;
  topic files provide full details.
- **AI/LLM integration section** added to README with classpath vs HTTP agent guidance.
- **Security notes** expanded in `docs/guide/protection.md`: ZIP bomb protection,
  encrypted temp file handling, Windows ACL fallback, XXE delegation to POI.

## [0.16.13] - 2026-04-15

### Added
- **Multi-row header support** — `ExcelReader.headerRows(int)` lets the reader
  combine N header rows into the effective column names, using the bottom-most
  non-blank value per column. This round-trips files written with multi-level
  `group(...)` on the writer side (where column header cells are often blank
  because they're part of a vertical merge with a group label above).
  ```java
  reader
      .headerRowIndex(1)     // last header row (0-based)
      .headerRows(2)         // 2 header rows total (group row + column header row)
      .build(in)
      .read(record -> ...);
  ```
  Default remains `headerRows(1)`; existing single-row behavior unchanged
  (including preserving empty-string headers).

## [0.16.12] - 2026-04-15

### Added
- **Split success/error read callbacks** — `read(Consumer<T> onSuccess,
  Consumer<RowError> onError)` delivers successfully parsed rows to one
  callback and failed rows (validation or mapping) to another. The library
  buffers nothing — the caller decides how to manage error memory (log,
  keep top N, stream to elsewhere, or abort by throwing).
  ```java
  reader.read(
      record -> process(record),
      err -> {
          if (err.type() == RowError.Type.VALIDATION) { ... }
          else { log.warn("row {} mapping failed", err.rowNum(), err.cause()); }
      });
  ```
- **`RowError` record** — carries row ordinal (1-based, excluding header),
  category (`VALIDATION` / `MAPPING`), human-readable messages, and the
  original cause when available.

### Changed
- `ReadResult<T>` record gains a nullable `cause` field so mapping-stage
  exceptions are preserved alongside the usual messages. The 3-arg
  constructor is retained for backward compatibility.

## [0.16.11] - 2026-04-15

### Added
- **Group header comments** — `writer.groupComment(String text, String... path)`
  and `groupComment(ExcelCellComment, String... path)` attach a note to a
  merged group header cell, identified by its path (outermost first). No-op
  if no column declares that path.
  ```java
  .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
  .groupComment("Quarterly revenue", "Financial", "Revenue")
  ```
- **Per-column header background color** — `headerBackgroundColor(ExcelColor)` /
  `headerBackgroundColor(int r, int g, int b)`. Overrides the workbook-wide
  `headerColor` for a single column only. Useful for highlighting alert or
  required columns.
- **`headerRowHeight(float points)`** — writer-level setter that applies the
  given height to every header row (including group header rows in multi-level
  setups). Pass `0` for default.
- **`rowNumberColumn(String name)`** — convenience for a 1-based sequential
  row-number column. Equivalent to a manual `column(name, (r, cursor) ->
  cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))` call. Works
  across auto-rollover sheets.

## [0.16.10] - 2026-04-15

### Fixed
- Vertically merged group header cells rendered with bottom-aligned text
  because the column name value existed in both the top-left and the
  original column header row cell of the merge region. Non-top cells in
  a vertical merge are now explicitly blanked so Excel honors the
  header style's `VerticalAlignment.CENTER`.

## [0.16.9] - 2026-04-15

### Added
- **Multi-level group headers** — `group(String... levels)` accepts N levels,
  ordered outermost (top) to innermost (just above the column header). Previously
  only one level was supported.

  ```java
  .column("Q1", Row::q1, c -> c.group("Financial", "Revenue", "2025"))
  ```
  Produces 3 group rows + 1 column header row. Adjacent columns with equal values
  on the same row are horizontally merged; columns with fewer levels are vertically
  merged with the column header cell. Header comments attach to the top-left cell
  of the merged region.

### Changed (Breaking)
- `ColumnStyleConfig.group(String)` signature changed to `group(String... levels)`.
  Source-compatible — existing `.group("X")` calls still work unchanged. The
  internal field `groupName` (String) became `groupNames` (String[]); if you
  reflected on it or read `ExcelColumn.getGroupName()`, use `getGroupNames()`
  instead.

## [0.16.8] - 2026-04-14

### Added
- `ColumnStyleConfig.headerComment(ExcelCellComment)` — overload for full
  customization of header comment (author, width, height) via the
  `ExcelCellComment` record.
- `ColumnStyleConfig.commentSize(int width, int height)` — column-level
  comment box size applied to both header and data-cell comments.
- `ExcelCellComment` fluent builders: `ExcelCellComment.of(text)`,
  `.author(String)`, `.size(int, int)`.
- `Comment.setAuthor(...)` is now actually propagated to POI (previously
  stored in the record but never written out).

  Precedence: `ExcelCellComment.size()` > `commentSize()` > POI default (2×3).

  ```java
  // Simple
  .headerComment("Enter in YYYY-MM-DD format")

  // Column-level size for all comments in the column
  .column(..., c -> c.comment(p -> "note").commentSize(4, 5))

  // Full customization
  .headerComment(ExcelCellComment.of("Reviewed").author("System").size(4, 3))
  ```

### Changed (Breaking)
- **`ExcelCellComment` record canonical constructor signature changed** from
  `(String text, String author)` to `(String text, String author, int width, int height)`.
  The old convenience constructor `new ExcelCellComment(text)` / `(text, author)`
  were removed. Use `ExcelCellComment.of(text)` going forward.

  This record was previously dead code (never instantiated outside its own
  file), so real-world impact is nil.

## [0.16.7] - 2026-04-14

### Added
- `ColumnStyleConfig.headerComment(String)` — attach a static comment (note) to
  a column's header cell. Useful for documenting expected input format or units
  (e.g., `"YYYY-MM-DD format"`, `"In KRW"`) as an Excel cell comment visible on
  hover. Complements the existing per-row `comment(Function)` targeting data
  cells.
  ```java
  .column("Birth Date", User::birth, c -> c
      .type(ExcelDataType.DATE)
      .headerComment("Enter in YYYY-MM-DD format"))
  ```
  With group headers, the comment attaches to the column header row (not the
  group row).

## [0.16.6] - 2026-04-13

### Added
- `ExcelHandler.writeTo(Path, String)` and `ExcelHandler.writeTo(Path, char[])` —
  convenience overloads that encrypt and write directly to a file path, mirroring
  the existing `writeTo(OutputStream, String/char[])` overloads. Useful for CLI/batch
  tools that don't need a Spring `StreamingResponseBody`. The `char[]` variant zeroes
  the password array after use.
- `ExcelWriter.freezeCols(int)` / `ExcelSheetWriter.freezeCols(int)` — freeze N
  columns from the left. Previously required `freezePane(N, 0)`.

### Changed (Breaking)
- **`ExcelWriter.freezePane(int rows)` renamed to `freezeRows(int)`.** Single-arg
  `freezePane(N)` was ambiguous — a caller reading `.freezePane(3)` couldn't tell
  whether it froze rows or columns without checking the Javadoc. The new names
  make intent explicit:
  - `freezeRows(int)` — freeze N rows below the header (replaces `freezePane(int)`)
  - `freezeCols(int)` — freeze N columns from the left (new)
  - `freezePane(int cols, int rows)` — freeze both axes (unchanged)
  Same rename applies to `ExcelSheetWriter`.

  Migration: `.freezePane(1)` → `.freezeRows(1)`.
- **`new CsvWriter<T>()` replaced by `CsvWriter.<T>create()`.** The public no-arg
  constructor is now package-private, matching the `ExcelWriter.create()` /
  `ExcelWorkbook.create()` pattern. Users switching between Excel and CSV writers
  now share one mental model.

  Migration: `new CsvWriter<Row>()` → `CsvWriter.<Row>create()`.

### Changed
- Strengthened Javadoc on `ExcelReadHandler` and `CsvReadHandler` with an explicit
  "Resource management" section, clarifying when temp files are released and why
  `readAsStream()` must be used with try-with-resources.

## [0.16.5] - 2026-04-13

### Changed (Breaking)
- **Unified `FileHandler` output API under `writeTo(...)`.** Every way a handler
  emits its payload is now a `writeTo` overload:
  - `FileHandler.write(OutputStream)` → `writeTo(OutputStream)`
  - `FileHandler.toFile(Path)` → `writeTo(Path)`
  - `ExcelHandler.consumeOutputStreamWithPassword(OutputStream, String)` → `writeTo(OutputStream, String)`
  - `ExcelHandler.consumeOutputStreamWithPassword(OutputStream, char[])` → `writeTo(OutputStream, char[])`
  - `CsvHandler.write(OutputStream)` → `writeTo(OutputStream)`
  Motivation: the old names collided conceptually with `ExcelWriter.write(Stream<T>)`
  (which produces a handler) — reading `.write(stream).write(out)` looked like two
  invocations of the same operation. They are not: the first pulls data *in*, the
  second pushes bytes *out*. Renaming the handler side to `writeTo` makes the
  direction explicit and frees the `write(Stream)` verb on the writer.
  `consumeOutputStreamWithPassword` (30 chars, different verb family) was also
  replaced by a plain `writeTo` overload so every output-stage call reads
  `.writeTo(destination, ...)`.

Migration:

```java
// Before
writer.column(...).write(stream).write(out);
writer.column(...).write(stream).toFile(path);
writer.column(...).write(stream).consumeOutputStreamWithPassword(out, "pw");
writer.column(...).write(stream).consumeOutputStreamWithPassword(out, pwdChars);

// After
writer.column(...).write(stream).writeTo(out);
writer.column(...).write(stream).writeTo(path);
writer.column(...).write(stream).writeTo(out, "pw");
writer.column(...).write(stream).writeTo(out, pwdChars);

// Spring StreamingResponseBody
.body(handler::write)   →   .body(handler::writeTo)
```

`ExcelWriter.password(String)` / `ExcelWorkbook.password(String)` is unchanged —
pre-set-then-`writeTo(out)` remains the preferred path. Passing the password to
`writeTo(out, pw)` still throws `IllegalStateException` if `password()` was
already set at the writer level.

## [0.16.4] - 2026-04-13

### Changed (Breaking)
- **`ExcelWriter.builder()` / `ExcelWorkbook.builder()` removed** — use
  `ExcelWriter.create()` / `ExcelWriter.create(opts -> ...)` and
  `ExcelWorkbook.create()` / `ExcelWorkbook.create(opts -> ...)` instead.
  The two-step `builder().xxx.build()` pattern was awkward because most
  configuration (columns, filters, callbacks, etc.) was already a fluent chain
  on the writer itself — `build()` served no real boundary.
- **`InitOptions` is now scoped to what POI forces at construction time.** Only
  `rowAccessWindowSize` remains in `InitOptions` (SXSSFWorkbook's constructor
  argument — cannot be changed after the workbook is created). The former
  `color` and `maxRows` builder methods are now fluent setters on the
  writer/workbook itself.
- **`color()` renamed to `headerColor()`** on both `ExcelWriter` and
  `ExcelWorkbook`, for consistency with the existing `tabColor()` and
  `rowColor()` methods and to make the target of the color unambiguous.
- **`ExcelWriter.forMap(Builder, String...)` → `forMap(Consumer<InitOptions>, String...)`.**
  Header color / max rows for a map writer are now set via fluent
  `.headerColor()` / `.maxRows()` on the returned writer.

Migration:

```java
// Before
ExcelWriter.<User>builder()
    .color(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .rowAccessWindowSize(500)
    .build()
    .column("Name", User::name)
    .write(stream);

// After
ExcelWriter.<User>create(opts -> opts.rowAccessWindowSize(500))
    .headerColor(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .column("Name", User::name)
    .write(stream);

// Or, when the default row window (1000) is fine — the common case:
ExcelWriter.<User>create()
    .headerColor(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .column("Name", User::name)
    .write(stream);
```

### Added
- **`ExcelWriter.headerColor(ExcelColor)` / `ExcelWorkbook.headerColor(ExcelColor)`**
  fluent setters (rebuild header style, preserving already-configured font
  settings).
- **`ExcelWriter.maxRows(int)`** fluent setter.

### Tests
- Strengthened header-color tests to verify exact RGB values instead of merely
  asserting "some fill color is present" (which passed even for the default
  white header).
- Added regression tests for the new `headerColor()` setter composition with
  `headerFontName/Size` (ensures font settings survive the internal
  `headerStyle` rebuild).

## [0.16.3] - 2026-04-12

### Fixed
- **readAsStream() producer thread deadlock** — when consumer closed the stream early
  (e.g., `limit()`, `takeWhile()`), the producer thread could block forever on
  `queue.put()`. Replaced with `offer(timeout)` + interrupt check loop so the producer
  exits cleanly on consumer close.
- **Temp file leak on read initialization failure** — if `Files.copy()` failed during
  `initTempFile()`, already-created temp directory and file were not cleaned up. Now
  calls `close()` before rethrowing the exception.

## [0.16.2] - 2026-04-12

### Fixed
- **Decrypted temp file not cleaned up** — when reading password-protected Excel files,
  the decrypted temp file was not deleted after reading, leaving sensitive data on disk.
  Now explicitly deleted in a finally block.
- **DataBar gradient maxColor not applied** — `dataBar(minColor, maxColor)` stored the
  max color but never wrote it to the Excel XML. 2-color gradient data bars now correctly
  render both colors.
- **ExcelDataType.FORMULA DDE security warning** — added javadoc warning that DDE formulas
  (e.g., `cmd|'/c calc'`) can execute commands. No runtime guard is applied because DDE
  pipe syntax is used by legitimate integrations (Bloomberg, Reuters). Input validation
  is the caller's responsibility.

## [0.16.1] - 2026-04-12

### Fixed
- **CSV injection whitespace bypass** — values like `" =cmd|'/c calc"` (leading spaces
  before formula characters) were not caught by the injection defense. Now detects formula
  characters after leading spaces.
- **Required column bypass in sparse Excel rows** — when trailing cells were empty (SAX
  omits them), required column validation was skipped. Now correctly reports missing
  required columns in sparse rows.
- **ExcelSheetWriter rollover header color loss** — custom `headerFontColor()` was not
  preserved on auto-rollover sheets. Fixed by passing workbook and style cache to the
  rollover header writer.
- **beforeHeaderWriter position ignored on rollover** — if `beforeHeaderWriter` wrote rows,
  rollover sheets placed headers at row 0 instead of after the preamble. Now captures and
  applies the preamble row offset.
- **Summary row overlap with afterData** — when both `afterData()` and `summary()` were
  used, the summary writer's return value was discarded, causing rows to overlap. Now
  correctly chains the row position.
- **ExcelImage null/empty data** — `ExcelImage.png(null)` and `ExcelImage.jpeg(new byte[0])`
  now throw `IllegalArgumentException` at creation time instead of failing later during write.

### Changed
- **ExcelSheetWriter.write() single-call guard** — calling `write()` twice on the same
  sheet now throws `ExcelWriteException` instead of silently corrupting data.

### Improved
- **docs/guide.md** — all `new ExcelReader<>(T::new, null)` / `new CsvReader<>(T::new, null)`
  patterns updated to modern `ExcelReader.setter(T::new)` / `CsvReader.setter(T::new)` API.
- **Progress callback warning** — documentation now warns that callbacks run on the
  read/write thread and must be fast and non-blocking.
- **Removed `docs/llms.txt`** — redundant with `META-INF/AI.md` in the JAR. Removed from
  project, CI workflow, and release checklist.

## [0.16.0] - 2026-04-12

### Added
- **`nullValue(Object)`** on `ColumnStyleConfig` — sets a default value to write when a
  column function returns null (e.g., `c -> c.nullValue("N/A")`). Works with `defaultStyle()`
  for writer-level fallback, overridable per column.
- **`freezePane(int cols, int rows)`** overload on `ExcelWriter` and `ExcelSheetWriter` —
  freezes both columns and rows simultaneously. The existing `freezePane(int rows)` is unchanged.
- **`required()`** on `ExcelReader` and `CsvReader` — marks the last registered column as
  required. Blank/empty cells produce a validation error in `ReadResult`.
- **`ReadColumn.required()`** — returns a new `ReadColumn` with `isRequired=true`.

### Changed (Breaking)
- **`FileHandler.write()` no longer throws `IOException`** — I/O errors are wrapped as
  unchecked exceptions (`ExcelWriteException`, `CsvWriteException`). `toFile()` wraps as
  `ExcelKitException`. Callers no longer need try-catch for checked exceptions.
- **`ExcelHandler.write()` and `consumeOutputStreamWithPassword()` no longer throw
  `IOException`** — same unchecked wrapping as above.
- **`AbstractReadHandler.readAsStream()` is now abstract** — the default implementation
  (which loaded all results into memory) is removed. Both `ExcelReadHandler` and
  `CsvReadHandler` already had proper lazy overrides.

### Fixed
- **ExcelWriter workbook leak on stream exception** — if the data stream threw during
  `write()`, the `SXSSFWorkbook` was never closed, leaking temp files. Now wrapped in
  try-catch with `closeWorkbookQuietly()`.
- **CsvReadHandler.readAsStream() CSVReader leak** — if initialization failed after creating
  the CSVReader but before returning the stream, the reader was not closed. Fixed with
  null-safe cleanup in all error paths.
- **CellData.asBigDecimal() precision loss** — previously routed through `asNumber()` →
  `Double` → `BigDecimal(toString())`, which lost precision for very large numbers. Now
  parses the cleaned string directly as `BigDecimal`.
- **CsvHandler.write() javadoc** — documented that IOException is wrapped as
  `CsvWriteException` (unchecked), clarifying the `FileHandler` contract.

### Improved
- **CellData regex pre-compiled** — currency symbol pattern (`[$,₩€%원]`) compiled once
  as `static final Pattern` instead of per-call `replaceAll()`.
- **Exception catch simplified** — `catch (A) throw; catch (B) throw;` patterns replaced
  with `catch (A | B) throw;` in both read handlers.
- **AtomicLong for lambda counters** — `long[]` array wrapper replaced with `AtomicLong`
  in `readStrict()` and `CsvReadHandler.readAsStream()`.
- **Duplicate header detection** — `ExcelReadHandler` now logs a warning when duplicate
  header names are found, noting that only the first occurrence is used in mapping mode.
- **README restructured** — reduced from 2043 to 231 lines. Detailed guide moved to
  `docs/guide.md`.
- **Release checklist updated** — reflects new README structure and `docs/guide.md`.

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
ExcelWorkbook.create(opts -> opts.color(ExcelColor.STEEL_BLUE))

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
    .onProgress(1000, (count, total) -> log.info("{}", count))
    .build(inputStream)
    .read(r -> process(r.data()));

// After
CsvReader.forMap()
    .dialect(CsvDialect.EXCEL)
    .onProgress(1000, (count, total) -> log.info("{}", count))
    .build(inputStream)
    .read(r -> process(r.data()));

// ─── Excel map reading now supports onProgress (new) ───
ExcelReader.forMap()
    .onProgress(1000, (count, total) -> log.info("read {}", count))
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
ExcelWriter.<User>create()
ExcelWriter.<User>create(opts -> opts.color(ExcelColor.STEEL_BLUE))
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
