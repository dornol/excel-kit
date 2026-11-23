# excel-kit

Fluent API-based Excel/CSV read-write Java library built on Apache POI (streaming SXSSF).
Not annotation-based — columns are defined programmatically via builder chains or lambda configs.

- Group: `io.github.dornol`
- Artifact: `excel-kit`
- Requires: Java 17+
- Runtime dependencies are intentionally declared as `compileOnly` by the
  library so applications can control versions:
  - `org.apache.poi:poi-ooxml` for Excel read/write
  - `com.opencsv:opencsv` for CSV read/write
  - `org.slf4j:slf4j-api` for logging
  - `jakarta.validation:jakarta.validation-api` for optional Bean Validation
- GitHub: https://github.com/dornol/excel-kit
- Javadoc: https://dornol.github.io/excel-kit/

## Quick Reference

| Task | Class | Pattern |
|------|-------|---------|
| Write Excel (typed) | `ExcelWriter<T>` | `.column("Name", T::getName).write(stream).writeTo(path)` |
| Write Excel (map) | `ExcelWriter.forMap(...)` | `ExcelWriter.forMap("Name", "Age").write(stream).writeTo(path)` |
| Write Excel (multi-sheet) | `ExcelWorkbook` | `wb.<T>sheet("Sheet1").column(...).write(stream)` → `wb.finish().writeTo(path)` |
| Write Excel (template) | `ExcelTemplateWriter` | `new ExcelTemplateWriter(template).list(...).write(stream, out)` |
| Read Excel (setter) | `ExcelReader<T>` | `ExcelReader.setter(T::new).validator(validator).column("Name", T::setName).required().read(in, r -> ...)` |
| Read Excel (map) | `ExcelReader.forMap()` / `forMapRaw()` | `ExcelReader.forMap().read(in, r -> r.data().get("Name"))` |
| Read Excel (mapping) | `ExcelReader.mapping()` | `ExcelReader.mapping(row -> new Record(row.get("Name").asString())).read(in, r -> ...)` |
| Summarize a read | `ReadSummary` / `ReadReport` | `reader.readWithSummary(in, r -> ...)` / `reader.readReport(in, 100)` |
| Detect tabular input | `TabularFileDetector` | `TabularFileDetector.detect(path)` or `.detectDetailed(in)` |
| Write CSV | `CsvWriter<T>` | `.column("Name", T::getName).write(stream).writeTo(path)` |
| Write CSV (map) | `CsvWriter.forMap(...)` | `CsvWriter.forMap("Name", "Age").write(stream).writeTo(path)` |
| Read CSV (setter) | `CsvReader<T>` | `CsvReader.setter(T::new).validator(validator).column("Name", T::setName).read(in, r -> ...)` |
| Read CSV (map) | `CsvReader.forMap()` | `CsvReader.forMap().read(in, r -> r.data().get("Name"))` |

## Packages and Imports

Every public type and its package, so imports can be written without guessing. The same table is in the
reference guide.

| Package | Public types | Use for |
|---------|--------------|---------|
| `io.github.dornol.excelkit.core` | `AtomicFileWriter`, `Cursor`, `ExcelKitException`, `FileHandler`, `ProgressCallback`, `RuntimeDependencyGuard` | Shared contracts: `FileHandler` (what a writer returns), progress callbacks, the base exception |
| `io.github.dornol.excelkit.core.read` | `AbstractReadHandler`, `AbstractReader`, `CancellationToken`, `CellConversionConfig`, `CellData`, `CellError`, `DuplicateHeaderPolicy`, `HeaderPolicy`, `ReadAbortException`, `ReadAbortReason`, `ReadColumn`, `ReadLimitExceededException`, `ReadLimits`, `ReadOptions`, `ReadProgress`, `ReadProgressCallback`, `ReadReport`, `ReadResult`, `ReadSecurityException`, `ReadSecurityPolicy`, `ReadSession`, `ReadStoppedException`, `ReadSummary`, `ReaderOptions`, `RowData`, `RowError`, `RowFunction` | Reading, shared by Excel and CSV: options, limits, security policy, cell/row access, results and errors |
| `io.github.dornol.excelkit.core.io` | `InputStreamSource`, `TempResourceCleaner`, `TempResourceContainer`, `TempResourceCreateException`, `TempResourceCreator`, `TempWorkspace` | Input sources and temp-file management |
| `io.github.dornol.excelkit.core.detect` | `DetectionConfidence`, `TabularDetectionResult`, `TabularFileDetector`, `TabularFileType` | Telling an .xlsx from a CSV by content |
| `io.github.dornol.excelkit.csv` | `CsvAfterDataWriter`, `CsvColumn`, `CsvDialect`, `CsvHandler`, `CsvQuoting`, `CsvReadException`, `CsvReader`, `CsvWriteException`, `CsvWriter` | CSV reading and writing |
| `io.github.dornol.excelkit.excel` | `AfterDataWriter`, `BeforeHeaderWriter`, `ColumnConfig`, `ColumnStyleConfig`, `ExcelChartConfig`, `ExcelColumn`, `ExcelConditionalRule`, `ExcelDataType`, `ExcelErrorReport`, `ExcelHandler`, `ExcelKitConfig`, `ExcelPrintSetup`, `ExcelRichText`, `ExcelSheetWriter`, `ExcelSummary`, `ExcelTemplateWriter`, `ExcelValidation`, `ExcelWorkbook`, `ExcelWriteErrorPolicy`, `ExcelWriteException`, `ExcelWriter`, `HeaderStyleConfig`, `RowStyleConfig`, `SheetContext`, `StreamingOptions`, `TableOptions`, `TemplateListWriter`, `WriteRowCallback` | Excel writing: `ExcelWriter`, `ExcelWorkbook`, templates, column and sheet configuration |
| `io.github.dornol.excelkit.excel.read` | `ExcelReadException`, `ExcelReader`, `ExcelSheetInfo`, `ExcelWorkbookInspector`, `ExcelWorkbookReader`, `SheetFailurePolicy`, `SheetReadReport`, `WorkbookReadReport` | Excel reading: `ExcelReader`, multi-sheet `ExcelWorkbookReader`, sheet metadata |
| `io.github.dornol.excelkit.excel.style` | `CellColorFunction`, `ExcelBorderStyle`, `ExcelColor`, `ExcelDataFormat` | Colors, borders and number formats |
| `io.github.dornol.excelkit.excel.feature` | `ExcelCellComment`, `ExcelHyperlink`, `ExcelImage` | Cell content beyond text: comments, hyperlinks, images |
| `io.github.dornol.excelkit.schema` | `ExcelKitSchema` | One column definition for both reading and writing, Excel or CSV |
| `io.github.dornol.excelkit.spring` | `CellErrorResponse`, `DownloadFileType`, `ErrorReportRow`, `ExcelKitErrorResponse`, `ExcelKitMultipartFile`, `ExcelKitResponse`, `ExcelKitTemplateResponse`, `ExcelKitUpload`, `ExcelKitUploadException`, `UploadCollectionLimits`, `UploadError`, `UploadReader`, `UploadResult`, `UploadSummary` | Spring MVC download and upload helpers (`excel-kit-spring` artifact) |

Typical imports:

```java
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
```

`ReadAbortException` (thrown by `readStrict()`/`readAll()` and when `maxErrors()` is exceeded) does not extend
`ExcelKitException`; catch it separately. Limits, security policy and read/write errors do extend it.

## Detailed Documentation

The following files are located in `META-INF/excel-kit/` within this JAR:

- **[quick-start.md](excel-kit/quick-start.md)** — Basic write/read examples for Excel and CSV
- **[column-config.md](excel-kit/column-config.md)** — Column styling, data types, header font color, default styles
- **[reading.md](excel-kit/reading.md)** — Excel/CSV reading (name-based, index-based, mapping mode, map reader)
- **[advanced.md](excel-kit/advanced.md)** — Multi-sheet, protection, charts, conditional formatting, images, validation
- **[csv.md](excel-kit/csv.md)** — CSV-specific features (dialect, delimiter, BOM, quoting, injection defense)

## Two Writer APIs

`ExcelWriter<T>` (single-type, auto-rollover):
```java
ExcelWriter.<Person>create()
    .column("Name", Person::name)
    .column("Age", Person::age, cfg -> cfg.type(ExcelDataType.INTEGER))
    .write(stream)
    .writeTo(out);
```

`ExcelWorkbook` (multi-sheet, different types per sheet):
```java
try (var wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users").column("Name", User::getName).write(userStream);
    wb.<Order>sheet("Orders").column("ID", Order::getId).write(orderStream);
    wb.finish().writeTo(out);
}
```

All writer APIs (`ExcelWriter`, `ExcelSheetWriter`, `CsvWriter`) use the same `.column("Name", fn, cfg -> cfg.type().bold())` pattern.

## Key API Notes (v0.16.0+)

- `FileHandler.writeTo()` throws unchecked exceptions only — no `throws IOException`
- `writeTo(Path)` writes atomically via `AtomicFileWriter` (temporary sibling, then replace); a failed write keeps the existing file
- `nullValue(Object)` on column config — default value for null cells (e.g., `c -> c.nullValue("N/A")`); always written, as text when it does not fit the column type
- Date types written as Excel dates: `Local*`, `ZonedDateTime`/`OffsetDateTime` (own wall-clock), `java.util.Date`/`java.sql.*` (JVM default zone); `Instant` stays text (convert it first)
- Freeze panes:
  - `freezeRows(int)` — freeze N rows below the header
  - `freezeCols(int)` — freeze N columns from the left
  - `freezePane(int cols, int rows)` — freeze both
- Encrypted output:
  - `.password("pw")` on writer → `handler.writeTo(out)` auto-encrypts
  - `handler.writeTo(out, "pw")` / `handler.writeTo(path, "pw")` — late-binding encryption (when the response is written);
    `handler.isPasswordSet()` — whether the writer already set one (both at once throws `IllegalStateException`);
    Spring: `ExcelKitResponse.excel(handler, filename, password)` rejects that before the response is built
- Writers use static factory `create()`: `ExcelWriter.<T>create()`, `ExcelWorkbook.create()`,
  `CsvWriter.<T>create()` (no public constructors)
- `required()` on reader columns — blank cells produce validation errors
- `column(List<String>, ...)` on readers — read by header aliases
- `strictHeaders()` — fail fast when configured headers are absent
- `DuplicateHeaderPolicy` — choose first, last, or fail for duplicate headers
- `CellError` and `RowError.cellErrors()` — structured cell-level read errors
- Numeric `as*()` conversions (`asInt/asLong/asDouble/asFloat/asNumber/asBigDecimal`) use the stored XLSX number, not the display format (`0.25` in a `0%` cell stays `0.25`); CSV/text cells use their text
- `asBigDecimal()` parses the decimal directly, avoiding Double conversion
- `rawNumber()` reads the stored XLSX number before display formatting; `rawValues()` on read results and errors returns the source row's values using it
- `asZonedDateTimeStrict(zone)` rejects local date-times in a DST gap or overlap
- `ExcelSheetWriter.write()` is single-call — second call throws `ExcelWriteException`
- `ExcelHandler.discard()` / `CsvHandler.discard()` release an unwritten handler (idempotent, no-op after `writeTo`); an unconsumed handler that is garbage collected is released as a safety net
- `ExcelImage.png()/jpeg()` validates non-null, non-empty byte array at creation time
- `ExcelImage.size(w, h)` sets cell span; `ExcelImage.fromUrl(url)` downloads and auto-detects type
- `rowStyle(predicate, config)` — conditional row-level bold/font/color (first match wins)
- `countRows()` on `ExcelReader` — opt-in pre-scan; `cursor.getTotalRows()` in progress callback
- CSV injection defense covers leading whitespace + formula characters (e.g., `" =cmd"`)

## Key API Notes (v0.16.9+)

- **Multi-level group headers** — `group(String... levels)` takes N levels, outermost first.
  `.column("Q1", Row::q1, c -> c.group("Financial", "Revenue", "2025"))` produces 3 group rows + 1 header row.
  Source-compatible with single-level `.group("X")`. Reflective readers: field renamed `groupName` → `groupNames`,
  getter `getGroupNames()` replaces `getGroupName()`.
- **Header customization** (writer-level):
  - `.headerRowHeight(float points)` — applies to every header row including group rows; `0` = default
  - `.headerFontName(String)` / `.headerFontSize(int)`
  - `.rowNumberColumn(String name)` — 1-based sequential column; shorthand for
    `column(name, (r, cur) -> cur.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))`, rolls over with `maxRows()`
- **Per-column header background** — `headerBackgroundColor(ExcelColor)` or `(int r, int g, int b)` on column config.
  Overrides workbook-wide `headerColor` for one column only.
- **Group header comments** — `writer.groupComment(String text, String... path)` or
  `groupComment(ExcelCellComment, String... path)`. Path is outermost-first and must match a declared
  `group(...)`; no-op otherwise.

## Key API Notes (v0.16.12+) — Reading

- **Split success/error callbacks** — `read(Consumer<T> onSuccess, Consumer<RowError> onError)` routes
  validated rows vs failed rows. Library buffers nothing — caller decides error memory policy.
- **`RowError`** record — `rowNum` (1-based, header excluded), `fileRowNum` (physical source row),
  `type` (`VALIDATION` / `MAPPING`), `messages`, nullable `cause`, and `cellErrors`.
- **`ReadResult<T>.cause()`** — nullable throwable from mapping stage. 3-arg constructor retained for
  backward compatibility.

## Key API Notes (v0.16.13+) — Reading

- **Multi-row header** — `ExcelReader.headerRows(int)` combines N header rows into effective column
  names, taking the bottom-most non-blank per column. Use with files written via multi-level `group(...)`:
  ```java
  reader.headerRowIndex(1).headerRows(2).read(in, r -> ...);
  ```
  Default `headerRows(1)` preserves existing single-row behavior including empty-string headers.

## Key API Notes (v0.18.0+) — Reading

- **Header aliases** — `column(List.of("Name", "User Name", "이름"), setter)` tries aliases in order.
- **Strict headers** — `strictHeaders()` fails before data rows when configured
  headers or selected map-mode columns are missing.
- **Duplicate headers** — `duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL|FIRST|LAST)` controls
  duplicate header resolution; the default is `FAIL`. A column with a blank header is left out of the header index, so it is not a key in `forMap()` rows.
- **Map mode values** — `ExcelReader.forMap()` returns displayed text; `ExcelReader.forMapRaw()` returns
  stored numbers as plain strings. Missing trailing cells are `""`.
- **Schema read options** — `ExcelKitSchema.Builder.column(name, aliases, ...)` and
  `requiredColumn(...)` flow into generated Excel/CSV readers.
- **Cell errors** — `CellError` is exposed through `ReadResult.cellErrors()` and
  `RowError.cellErrors()` for import UIs.

## Key API Notes (v0.16.14+)

- **Document properties** — `documentProperty(key, value)` on `ExcelWriter` / `ExcelWorkbook`.
  Standard keys (`title`, `subject`, `author`/`creator`, `keywords`, `description`, `category`)
  map to core properties; others become custom properties.
- **Fluent named ranges** — `namedRange(name, columnIndex)` registers a workbook-scoped named
  range covering all data rows in the given column. Replaces manual `afterData` callback usage.
- **Header style config** — `headerStyle(cfg -> cfg.alignment(...).border(...).bold(...).wrapText(...))`
  on `ExcelWriter` / `ExcelWorkbook`. Overrides default header alignment/bold/border/wrap.

## Key API Notes (v0.28.0+)

- **Schema package** — `ExcelKitSchema` lives in `io.github.dornol.excelkit.schema` (it moved from `core`, which no
  longer depends on `excel`/`csv`). Register a schema reader in a workbook reader with
  `sheet(name, schema.excelReader(supplier, validator), consumer)`.
- **Untrusted uploads** — readers have no limits by default; use `limits(ReadLimits.UPLOAD)` (50 MiB, 50 sheets,
  1,000 columns, 32,767 characters per cell, 10 million cells) and, for Excel, `securityPolicy(ReadSecurityPolicy.STRICT)`.
- **Sheet discovery** — `ExcelWorkbookInspector.create().sheetNames(input)` / `.headers(input, sheetIndex, headerRowIndex)`
  (`.headers(input, sheetIndex, headerRowIndex, headerRows)` for a header block of several rows, such as `group(...)`)
  (`InputStream`, `Path`, `InputStreamSource`). `.password(...)` opens an encrypted workbook; `.limits(ReadLimits.UPLOAD)`
  caps an upload (`maxInputBytes`, `maxSheets`, `maxColumns`) and fails with `ReadLimitExceededException`.
- **Style cap** — `maxCellStyles(int)` on `ExcelWriter` / `ExcelSheetWriter` limits the cell styles that `cellColor`,
  `rowColor` and `rowStyle` colors may create (default 64,000); a function returning more distinct colors fails fast.
- **CSV** — `CsvWriter.lineSeparator("\r\n"|"\n"|"\r")`. `CsvReader` refuses a delimiter, quote and escape character
  that cannot be told apart (and line breaks / NUL as delimiter); `CsvWriter.delimiter(...)` refuses `"`, line breaks and
  NUL. `ignoreLeadingWhiteSpace` defaults to `true` (a space before a quote is skipped). A decode failure names the
  charset and says to set `charset(...)`.
- **Templates from users** — `new ExcelTemplateWriter(upload)` fails with `ExcelWriteException` for a stream that is not
  a usable `.xlsx` (not OOXML, damaged, no sheets) and with `IOException` for a stream that cannot be read; `cell`, a
  list's `write` and saving fail with `ExcelWriteException` too, never with POI's own exceptions.
- **Rich text** — unformatted text and every formatted segment start from the column's font; `ExcelTemplateWriter.cell`
  accepts `ExcelRichText`. A `char[]` value is written as its characters.
- **Charts** — with `maxRows()` rollover each sheet gets its own chart; `position(...)` and the chart's columns are
  validated.
- **Spring** — downloads send `Cache-Control: private, max-age=N`; `ExcelKitMultipartFile.requireExcelOrCsvContent`
  checks the real ZIP signature and that a `.csv` is UTF-8 text.

## Spring response compatibility

Spring responses accept an explicit `Duration` for cache max-age through
`ExcelKitResponse.excel/csv` and `ExcelKitErrorResponse.excel/csv` overloads.
For example, `ExcelKitResponse.excel("report", Duration.ofMinutes(1))`
returns a builder whose `.body(handler::writeTo)` streams the download.
The default is 10 seconds. Upload helpers use `ExcelKitUpload.excel/csv`.

## Key API Notes (Unreleased) — Dependent dropdown

`c.dependentDropdown("Category", Map.of("A", List.of("1", "2"), "B", List.of("3")))` makes a cascading dropdown:
the cell offers the options of the value the named parent column holds in the same row. The options go to a hidden
helper sheet (`_excelkit_dep_N`); equal option sets share one sheet. The parent value is matched as text; use a
`LinkedHashMap` to keep the parent order (a numeric parent cell matches its key written as text; an option that is a plain number is
written as a number to the helper sheet, because Excel refuses a typed number against a list of text). It is a validation
of the column: it cannot be combined with `dropdown(...)`, a `validation(...)` set after it replaces it, and an unknown
parent column fails when the sheet is written. `ExcelValidation.list("A","B")` and `ExcelValidation.dependentList(parent,
options)` are the same lists as validations, so `inputMessage(..)` / `allowBlank(false)` apply to them.

`zoom(percent)` (10-400) on any sheet writer; `ExcelSheetWriter.hideSheet()` / `veryHideSheet()` hide a workbook sheet
(`finish()` throws `ExcelWriteException` if no sheet stays visible).

Page breaks: `printSetup(p -> p.rowBreakEvery(n))`, and `pageBreakOnChange(keyFn)` / `pageBreakWhen((previous, current) -> bool)`
on `ExcelWriter` / `ExcelSheetWriter` (a new printed page where the data changes; at most 1026 breaks per sheet).

`writer.stripedRows()` / `stripedRows(color)` (also on `ExcelSheetWriter`) shades every second data row without a `table(...)`; it is the
lowest-priority row background (`cellColor`, `rowStyle`, `rowColor` win).

`c.mergeSameValues()` merges runs of equal values of a column vertically (a department over its rows; rows sorted). A column
merges only inside the merged regions of the merge columns to its left; empty values and one-row runs are not merged;
the cells below the first keep their values; rejected with `table(...)` and on `TemplateListWriter`.

`reader.uniqueKey("Employee ID")` (`ExcelReader`, `CsvReader`, also `ReaderOptions`) fails every row whose key (one or several
headers) was seen in an earlier row: `VALIDATION` `RowError` `Duplicate key [..]; first seen at row N`, key cells in
`cellErrors()`. Empty key cells are not checked; keys are held in memory.

`c.requiredMark()` shows `*` after a column's header text (a number format `@" *"`: the header cell keeps the plain
name, so the file reads back by the same names). A schema `requiredColumn(...)` marks its header automatically.

`ExcelReader.fillMergedCells()` gives every cell of a merged region the value of its top-left cell in the data rows
(a department merged by hand over its rows reads in each row); one extra scan of the sheet, at most 200,000 regions.

`ExcelErrorReport.create(headers).write(errors)` writes the failed rows of an upload (row number, uploaded values,
failed cells shaded with a comment, a messages column) as an `ExcelHandler`, from `RowError`s or a `ReadReport`.

`ExcelValidation` also takes `inputTitle(..)` / `inputMessage(..)` (the tooltip shown when the cell is selected) and
`allowBlank(false)` (an empty cell no longer passes; default is allowed).
