# excel-kit — Reading

> Other topics: [Index](../AI.md) | [Quick Start](quick-start.md) | [Column Config](column-config.md) | [Advanced](advanced.md) | [CSV](csv.md)

## Three Read Modes

### 1. Setter Mode (Mutable Objects)
```java
ExcelReader.setter(User::new)
    .validator(validator)
    .column("Name", (u, cell) -> u.name = cell.asString())
    .column("Age", (u, cell) -> u.age = cell.asInt())
    .read(inputStream, result -> {
        if (result.success()) { User u = result.data(); }
        else { log.warn("Read failed: {}", result.messages()); }
    });
```

### 2. Mapping Mode (Records / Immutable Objects)
```java
ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
    row.get("Name").asString(),
    row.get("Age").asInt(),
    row.get("City").asString()
)).read(inputStream, result -> { ... });
```

### 3. Map Mode (No POJO)
```java
ExcelReader.forMap()
    .read(inputStream, result -> {
        Map<String, String> row = result.data();
        String name = row.get("Name");
    });
```
- `forMap()` returns displayed text (`2.5` in a `0` cell → `"3"`); `ExcelReader.forMapRaw()` returns stored numbers as plain strings
- Rows shorter than the header return `""` for missing trailing columns

## Column Mapping Strategies

### Name-Based (Order-Independent)
```java
reader.column("Name", (u, cell) -> u.name = cell.asString())  // matches header "Name"
```

### Header Aliases
```java
reader.column(List.of("Name", "User Name", "이름"), (u, cell) -> u.name = cell.asString())
      .strictHeaders()
      .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL);
```

### Index-Based
```java
reader.columnAt(0, (u, cell) -> u.name = cell.asString())  // column index 0
      .columnAt(3, (u, cell) -> u.age = cell.asInt())       // column index 3
```

### Positional (Legacy)
```java
reader.column((u, cell) -> u.name = cell.asString())  // first column
      .column((u, cell) -> u.age = cell.asInt())       // second column
```

Can be mixed: name-based + index-based in same reader.

## CellData Conversion Methods

| Method | Return | Description |
|--------|--------|-------------|
| `asString()` | String | Formatted cell text |
| `asInt()` | Integer | Parse as integer (null if empty) |
| `asLong()` | Long | Parse as long |
| `asDouble()` | Double | Parse as double |
| `asFloat()` | Float | Parse as float |
| `asBigDecimal()` | BigDecimal | Parse the stored XLSX number (or the cell text for CSV/text cells) as BigDecimal |
| `rawNumber()` | BigDecimal | Read the stored XLSX number, including cached numeric formula results |
| `asBoolean()` | Boolean | Parse as boolean |
| `asLocalDate()` | LocalDate | Parse as date |
| `asLocalDateTime()` | LocalDateTime | Parse as datetime |
| `asZonedDateTime(ZoneId)` | ZonedDateTime | Parse with timezone |
| `asZonedDateTimeStrict(ZoneId)` | ZonedDateTime | Like `asZonedDateTime`, but rejects times in a DST gap or overlap |
| `as(Function<String, T>)` | T | Custom conversion (e.g., `UUID::fromString`) |
| `asInt(defaultValue)` | int | With default for null/empty |
| `asString(defaultValue)` | String | With default for null/empty |
| `isEmpty()` | boolean | Check if cell is empty |
| `value()` | String | Formatted cell text (nullable) |

`rawNumber()` preserves the digits stored in the XLSX file even when the General
format rounds large numbers for display. It returns the underlying fraction for
percentages and the serial number for dates. For text, blank cells, and CSV input,
it returns `null`.

On read results and errors, `rawValues()` returns the source row's values, using the stored
XLSX number for numeric cells.

## RowData Methods (Mapping Mode)

| Method | Description |
|--------|-------------|
| `get(String headerName)` | Cell by header name (throws if not found) |
| `get(int columnIndex)` | Cell by 0-based index |
| `has(String headerName)` | Check if header exists |
| `size()` | Number of cells |
| `headerNames()` | List of header names |

## Configuration Options

```java
reader
    .sheetIndex(0)          // default: 0
    .headerRowIndex(0)      // default: 0 — 0-based index of the LAST header row
    .headerRows(1)          // default: 1 — total header row count (v0.16.13+, Excel only)
    .strictHeaders()        // fail fast when configured headers are missing
    .onProgress(10_000, (count, cursor) -> log.info("Read {} rows", count))
    .countRows()            // opt-in pre-scan for total row count (v0.16.15+, Excel only)
    .read(inputStream, result -> { ... });
```

With `countRows()`, `cursor.getTotalRows()` returns the total data row count in the progress callback (otherwise `-1`).

### Multi-row headers (v0.16.13+, Excel)

Files written with multi-level `group(...)` have blank column-header cells due to vertical merges.
`headerRows(int)` combines N header rows per column, taking the bottom-most non-blank value:

```java
ExcelReader.<Record>mapping(row -> ...)
    .headerRowIndex(1)     // last header row is row index 1
    .headerRows(2)         // 2 header rows (group row + column header row)
    .read(in, result -> ...);
```

Default `headerRows(1)` = existing single-row behavior. Empty-string headers preserved.

## Read Methods

| Method | Description |
|--------|-------------|
| `.read(Consumer<ReadResult<T>>)` | Process each row (skips failures) |
| `.read(Consumer<T> onSuccess, Consumer<RowError> onError)` | Split success/error callbacks (v0.16.12+) |
| `.readStrict(Consumer<ReadResult<T>>)` | Throws `ReadAbortException` on first validation failure (not an `ExcelKitException`) |
| `.readWhile(Predicate<ReadResult<T>>)` | Stops normally when the callback returns false |

Reader configurations are reusable; each call creates an independent one-shot session.

### Split success/error callbacks (v0.16.12+)

Route valid rows and failed rows to separate callbacks. The library buffers nothing — caller decides
error memory policy (log, keep top N, abort by throwing):

```java
reader.read(
    record -> process(record),
    err -> {
        if (err.type() == RowError.Type.VALIDATION) {
            log.warn("row {} invalid: {}", err.rowNum(), err.messages());
        } else {  // MAPPING
            log.error("row {} mapping failed", err.rowNum(), err.cause());
        }
    });
```

**`RowError` fields**: `rowNum()` (1-based, header excluded), `fileRowNum()` (physical source row),
`type()` (`VALIDATION` / `MAPPING`), `messages()` (human-readable),
`cause()` (nullable throwable from mapping stage), and `cellErrors()` (`CellError` entries with
column index, header name, cell value, and message).

```java
reader.read(
    row -> importRow(row),
    error -> error.cellErrors().forEach(cell ->
        log.warn("file row {}, {}='{}': {}",
            error.fileRowNum(), cell.headerName(), cell.cellValue(), cell.message()))
);
```

Keep `cellErrors()` structured when returning API responses so clients can point
users to a specific row, header, and submitted value.

`ExcelKitSchema` generated readers support the same header controls:

```java
schema.excelReader(User::new)
    .validator(validator)
    .strictHeaders()
    .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
    .read(inputStream, user -> importUser(user), error -> log.warn("{}", error.cellErrors()));
```

`ReadResult<T>.cause()` is also available for fail-paths in the unified `read(Consumer<ReadResult<T>>)` form.

## Duplicate keys

`reader.uniqueKey("Employee ID")` (or several headers for a composite key) fails every row whose key was already seen
in an earlier row with a `VALIDATION` `RowError`: `Duplicate key [Employee ID=1001]; first seen at row 2`, key cells in
`cellErrors()`. `ExcelReader` and `CsvReader`, every mode. Compared as trimmed, case-sensitive text (numeric cells by
their stored number); an empty key cell is not checked (use `required()`); keys are held in memory, so bound huge files
with `maxRows(...)`. Also on `ReaderOptions.builder().uniqueKey(...)`.

## Merged cells

`ExcelReader.fillMergedCells()` gives every cell of a merged region the value of its top-left cell (Excel stores it
there only), in the data rows: a department merged over its rows reads in each row. Header rows are unchanged, a cell
with a value of its own keeps it, a row absent from the file is not created. It scans the sheet once more before the
read (like `countRows()`); at most 200,000 merged regions per sheet.

## Error report workbook

`ExcelErrorReport.create(headers).write(errors)` (`Iterable<RowError>`, `Stream<RowError>` or `ReadReport`) returns an
`ExcelHandler` for an Excel file of the failed rows: row number in the uploaded file, uploaded values, failed cells shaded
with their message as a cell comment, and a last column with the row's messages. Options: `rowColumnName`,
`errorColumnName`, `sheetName`, `errorColor`. It uses `RowError.rawValues()` and `CellError.columnIndex()`; blank or repeated
headers are renamed.

## Summaries, bounded reports, and file detection

`readWithSummary(...)` returns a `ReadSummary` with total, successful, and failed row counts,
early-stop state, and elapsed time. `readReport(input, maxCollectedErrors)` adds a bounded
error sample and an `errorsTruncated` indicator.

`TabularFileDetector.detect(path)` distinguishes XLSX, XLS, CSV, and unknown content by
signature. `detectDetailed(input)` also reports confidence, charset, delimiter, and whether
excel-kit supports reading the detected format. Caller-provided streams remain caller-owned.

## Files that are not an .xlsx

A file that is not an `.xlsx` fails with an `ExcelReadException` that says what it is, judged by content (not by name), from
every way of reading: an Excel 97-2003 `.xls` ("save it as .xlsx and read that"), a password-protected workbook read without
`password(...)`, another OLE2 document, text or a CSV ("not an Excel workbook"), an empty file.

## Multi-Sheet Discovery

```java
ExcelWorkbookInspector inspector = ExcelWorkbookInspector.create();
List<ExcelSheetInfo> sheets = inspector.sheetNames(inputStream);   // ExcelSheetInfo: index(), name()
List<String> headers = inspector.headers(inputStream, sheetIndex, headerRowIndex);
// A header of several rows (written with group(...)): the block ends at headerRowIndex and is headerRows high;
// the lowest non-blank cell of each column is its name, as for ExcelReader.headerRows(...).
List<String> names = inspector.headers(inputStream, sheetIndex, headerRowIndex, headerRows);
```

`InputStream`, `Path` and `InputStreamSource` overloads. Encrypted workbook: `.password("secret")` on the inspector.
Untrusted upload: `.limits(ReadLimits.UPLOAD)` stops the copy at `maxInputBytes` and refuses more than `maxSheets` sheets
or a header wider than `maxColumns` with `ReadLimitExceededException`. `ExcelWorkbookReader` uses the password of the readers it is given.

## Bean Validation

```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

// Setter mode
ExcelReader.setter(User::new).validator(validator).column(...).read(in, result -> {
    if (!result.success()) {
        // result.messages() contains validation errors
    }
});

// Mapping mode
ExcelReader.mapping(row -> new Person(...)).validator(validator).read(in, result -> { ... });
```

## CSV Reading

Same API pattern as Excel reading:

```java
// Setter mode
new CsvReader<>(Product::new, null)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .read(inputStream, result -> { ... });

// Mapping mode
CsvReader.<Product>mapping(row -> new Product(
    row.get("Name").asString(), row.get("Price").asInt()
)).read(inputStream, result -> { ... });

// Map mode
CsvReader.forMap().read(inputStream, result -> {
    Map<String, String> row = result.data();
});
```

## Input byte limits

Bounded reader streams request at most the remaining `maxInputBytes` budget plus
one byte to detect overflow, including skip operations. Exact-size and empty
zero-limit streams are valid. A limit error's `actual()` value is an observed
count, not necessarily the full file size; byte counts saturate at `Long.MAX_VALUE`.
Caller-provided streams and paths remain caller-owned. File size is captured before
temporary-file cleanup so `ReadLimitExceededException` is preserved. Custom
`ReadSecurityPolicy.maxCompressionRatio` values must be finite and positive.

Readers have no limits by default. For untrusted uploads use `limits(ReadLimits.UPLOAD)` and, for Excel,
`securityPolicy(ReadSecurityPolicy.STRICT)`.

## External workbook behavior

Streaming cell and header formatting uses each workbook's 1900/1904 date system.
Formula results use saved caches without recalculation. Path-based data/metadata
reads and shared-package reads use read-only OOXML packages and preserve original
file bytes. Encrypted reads own a separate plaintext workspace and remove it after
reading or security-preflight failure, preserving caller-owned encrypted files.
