# Reading Excel & CSV

> [Back to Index](index.md)

## Setter Mode — Mutable Objects

### Explicit execution sessions

For advanced lifecycle control, open a session explicitly:

```java
try (ReadSession<Person> session = CsvReader.<Person>mapping(this::map).open(path)) {
    List<Person> people = session.readAll();
}
```

Sessions created from `Path` or `InputStreamSource` own their input resources. A session opened
from a caller-owned `InputStream` does not close that stream. The existing direct `read(...)` APIs
remain available for the shorter form.

```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.name = cell.asString()).required()
    .column("Age", (u, cell) -> u.age = cell.asInt())
    .read(inputStream, result -> {
        if (result.success()) {
            User u = result.data();
        } else {
            log.warn("Read failed: {}", result.messages());
        }
    });
```

## Mapping Mode — Immutable Objects / Records

```java
record PersonRecord(String name, Integer age, String city) {}

ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
        row.get("Name").asString(),
        row.get("Age").asInt(),
        row.get("City").asString()))
    .read(inputStream, result -> {
        if (result.success()) {
            PersonRecord p = result.data();
        }
    });
```

Columns are matched by header name — order in the file doesn't matter.

Header aliases are tried in order. The first alias found in the file is used.

```java
ExcelReader.setter(User::new)
    .column(List.of("Name", "User Name", "이름"), (u, cell) -> u.name = cell.asString())
    .read(inputStream, ...);
```

**RowData access methods:**

| Method | Description |
|--------|-------------|
| `get(String headerName)` | Get cell by header name (throws if not found) |
| `get(int columnIndex)` | Get cell by 0-based index |
| `has(String headerName)` | Check if header exists |
| `size()` | Number of cells in this row |
| `headerNames()` | List of header names |

## Map Mode — Schema-less

```java
ExcelReader.forMap()
    .read(inputStream, result -> {
        Map<String, String> row = result.data();
        String name = row.get("Name");
    });

// CSV equivalent
CsvReader.forMap().read(inputStream, result -> { ... });
```

`ExcelReader.forMap()` returns each cell's displayed text, so a `2.5` cell formatted as `0`
is `"3"` and a General-format 15-digit id can be `"1.23457E+14"`. Use `ExcelReader.forMapRaw()`
(or `forMapRaw("Name", "Age")`) to get stored numbers as plain strings (`"2.5"`, `"0.25"`,
`"123456789012345"`). Date cells become Excel serial numbers. Rows with fewer cells than headers
return `""` for the missing trailing columns.

## Column Matching

### Name-Based (default)

```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.name = cell.asString())
    .column("City", (u, cell) -> u.city = cell.asString())  // other columns ignored
    .read(inputStream, result -> { ... });
```

### Index-Based

```java
ExcelReader.setter(User::new)
    .columnAt(0, (u, cell) -> u.name = cell.asString())
    .columnAt(2, (u, cell) -> u.city = cell.asString())
    .columnAt(4, (u, cell) -> u.phone = cell.asString())
    .read(inputStream, result -> { ... });
```

Can be mixed with name-based mapping.

### Positional with Skip

```java
reader
    .column((u, cell) -> u.name = cell.asString())
    .skipColumn()
    .skipColumns(2)
    .column((u, cell) -> u.age = cell.asInt())
    .read(inputStream, result -> { ... });
```

## Required Columns

```java
ExcelReader.setter(User::new)
    .column("Name", (u, c) -> u.setName(c.asString())).required()
    .column("Email", (u, c) -> u.setEmail(c.asString())).required()
    .column("Nickname", (u, c) -> u.setNick(c.asString()))  // optional
    .read(inputStream, result -> {
        if (!result.success()) {
            // result.messages() -> "Required column 'Name' is empty"
        }
    });
```

## Split Success / Error Callbacks (v0.16.12+)

```java
reader.read(
    record -> process(record),
    err -> {
        switch (err.type()) {
            case VALIDATION -> log.warn("row {} invalid: {}", err.rowNum(), err.messages());
            case MAPPING    -> log.error("row {} mapping failed", err.rowNum(), err.cause());
        }
    });
```

`RowError` fields:
- `rowNum()` — 1-based row ordinal (header rows excluded)
- `fileRowNum()` — 1-based physical row number in the original file

For a CSV record spanning several lines, `fileRowNum()` identifies its first line.
On read results and errors, `rawValues()` returns the stored XLSX numeric text,
including date and time serials and cached numeric formula values.
- `type()` — `VALIDATION` or `MAPPING`
- `messages()` — human-readable messages (list)
- `cause()` — nullable throwable (for `MAPPING` errors)
- `cellErrors()` — structured `CellError` entries with `columnIndex`, `headerName`, `cellValue`, and `message`

Example import error response:

```java
reader.read(inputStream,
    user -> importUser(user),
    error -> error.cellErrors().forEach(cell ->
        log.warn("file row {}, column {}, value {}: {}",
            error.fileRowNum(), cell.headerName(), cell.cellValue(), cell.message()))
);
```

For API responses, keep `cellErrors()` structured instead of flattening it into a
single message. This lets clients highlight the exact row, header, and submitted
value:

```json
{
  "fileRowNum": 2,
  "messages": ["Failed to set column 'Price'"],
  "cellErrors": [
    {
      "columnIndex": 2,
      "headerName": "Price",
      "cellValue": "not-a-number",
      "message": "Failed to set column 'Price'"
    }
  ]
}
```

When using `ExcelKitSchema`, the same read options apply to generated readers:

```java
ExcelKitSchema<User> schema = ExcelKitSchema.<User>builder()
    .requiredColumn("Name", List.of("User Name", "이름"), User::getName, User::setName)
    .column("Age", User::getAge, User::setAge)
    .build();

schema.excelReader(User::new)
    .validator(validator)
    .strictHeaders()
    .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
    .read(inputStream, user -> importUser(user), error -> log.warn("{}", error.cellErrors()));
```

## Error Report Workbook

`ExcelErrorReport` turns the `RowError`s of an upload (or a `ReadReport`) into an Excel file to hand back to the
person who uploaded it: the row number in the uploaded file, the uploaded values, the failed cells shaded with their
message in a cell comment, and a last column with the messages of the row.

```java
var errors = new ArrayList<RowError>();
reader.read(input, row -> save(row), errors::add);

ExcelHandler report = ExcelErrorReport.create(List.of("Name", "Age"))   // the headers of the uploaded file
        .errorColumnName("Reason")                                    // optional: also rowColumnName, sheetName, errorColor
        .write(errors);                                               // Iterable, Stream or ReadReport
// e.g. ExcelKitResponse.excel("upload-errors").body(report::writeTo)
```

It holds the failed rows only, from `RowError.rawValues()`; the original file is not read again. Headers start at the
file's first column (as `CellError.columnIndex()` counts) and a blank or repeated one is renamed. A `ReadReport` holds
a bounded number of errors (`errorsTruncated()`).

## Advanced Options

### Reusable reader options and list results

Common reader settings can be shared across reader instances:

```java
ReaderOptions options = ReaderOptions.builder()
    .strictHeaders(true)
    .maxRows(10_000)
    .skipBlankRows(true)
    .limits(new ReadLimits(10 * 1024 * 1024, 10, 100, 10_000))
    .build();

List<Person> people = ExcelReader.<Person>mapping(this::map)
    .options(options)
    .readAll(inputStream);
```

`readAll(...)` returns only successful rows and uses the same fail-fast behavior as
`readStrict(...)`. Both throw `ReadAbortException`, which does not extend `ExcelKitException` (see
[Exception Handling](reference.md#exception-handling)). For mapping code that tolerates optional columns, use `row.getOrNull("Name")`.
For blank numeric cells, `cell.asOptionalInt()`, `asOptionalLong()`, and `asOptionalDouble()` avoid
manual nullable checks.

**Header row index** (files with metadata rows above header):
```java
reader.headerRowIndex(2)  // 3rd row as header (0-based)
```

**Strict headers** fail before data rows when a positional or index-based column has no header:

```java
reader.strictHeaders();
reader.strictHeaders(false);  // default
```

**Duplicate headers** fail before data rows by default, so a repeated column is never dropped
silently. A column whose header is blank is ignored: it has no name to look it up by, so it is not a
key in `forMap()` rows (and several blank headers never collide or count as duplicates). Columns read
by position (`column(setter)`, `row.get(int)`) do not depend on their headers. You can choose another
policy for duplicates:

```java
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL);  // default
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.FIRST);
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST);
```

**Multi-row headers** (v0.16.13+, Excel only):

For files with multi-level group headers, use `headerRows(int)` to combine N header rows.
Takes the bottom-most non-blank value per column.

```java
ExcelReader.<Row>mapping(row -> new Row(
        row.get("Q1").asInt(), row.get("Q2").asInt(), row.get("Profit").asInt()))
    .headerRowIndex(2)  // last header row (0-based)
    .headerRows(3)      // 3 rows: 2 group + 1 column header
    .read(in, result -> ...);
```

**Specific sheet:**
```java
reader.sheetIndex(1)  // 2nd sheet (0-based)
```

**Early completion without exceptions:**
```java
reader.readWhile(inputStream, result -> {
    process(result);
    return shouldContinue(result); // false stops normally
});
```

Callback reading is row-by-row and does not load the whole file into memory. A caller-provided
`InputStream` remains caller-owned; use try-with-resources around it.
`Path` and `InputStreamSource` are available consistently for `read`, `readStrict`, and
`readWhile`. Path inputs are read directly and are never modified or deleted. Streams opened
by an `InputStreamSource` are closed by excel-kit.

Excel sheet discovery follows the same ownership rule: `ExcelWorkbookInspector.sheetNames(InputStream)` and
`headers(InputStream, ...)` consume but do not close the supplied stream, while their
`Path` and `InputStreamSource` overloads manage their own resources.

Use `readWithSummary(...)` when aggregate counts and elapsed time are needed, or
`readReport(input, maxCollectedErrors)` for a bounded error sample. Untrusted inputs can be
guarded with `limits(new ReadLimits(maxBytes, maxSheets, maxColumns, maxCellCharacters, maxTotalCells))`.

> **Reading files from untrusted users?** A reader applies no limits and allows formulas and
> external links unless you say otherwise, and it copies the whole input to a temp file. For an
> upload use `limits(ReadLimits.UPLOAD)` (50 MiB, 50 sheets, 1,000 columns, 32,767 characters per
> cell, 10 million cells) and, for Excel, `securityPolicy(ReadSecurityPolicy.STRICT)`.

`headerPolicy(...)` provides common trim, case-insensitive, whitespace, and Unicode-normalized
matching presets. Long-running reads can use `cancellationToken(...)` and
`onReadProgress(interval, callback)` without transferring execution to a library thread.
Byte limits are enforced while copying a stream, so oversized uploads are stopped before the
entire body is materialized. Each bounded read or skip requests no more than the
remaining byte budget plus one byte to detect overflow. Exact-size inputs and an
empty stream with a zero-byte limit are allowed. `ReadLimitExceededException`
exposes the limit kind, configured value, and observed value; the latter may be
only the first disallowed byte, not the full input size. Observed byte counts
saturate at `Long.MAX_VALUE` instead of wrapping. Post-copy validation captures
file size before cleanup, so cleanup cannot mask a limit error. Caller-owned
streams and paths remain caller-owned. Summary, report, and `readWhile` APIs accept `InputStream`, `Path`,
and `InputStreamSource`; progress callbacks always receive a terminal event.
For untrusted XLSX files, `securityPolicy(ReadSecurityPolicy.STRICT)` rejects formulas and
external workbook links (including links declared through worksheet relationship files) before
sheet rows are mapped.
Strict inspection also bounds ZIP entry count, each decompressed worksheet entry, total scanned
bytes, and compression ratio. Custom compression ratios must be finite and positive;
`NaN` and either infinity are rejected. CSV callers can use `readDetected(...)` to apply sampled charset and
delimiter detection without closing the caller stream.

Before selecting a reader, `TabularFileDetector.detect(path)` identifies XLSX, legacy XLS,
CSV, or unknown content from its signature instead of trusting the filename. Use
`detectDetailed(input)` when confidence, charset, delimiter, and read-support metadata are
needed. Caller-provided streams remain caller-owned.

**Bean Validation:**
```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
ExcelReader<User> reader = ExcelReader.setter(User::new)
    .validator(validator);
```

**Large file support:**
```java
ExcelReader.configureLargeFileSupport();  // call once at startup; JVM-global
```

**Progress callback:**
```java
reader.onProgress(10_000, (count, cursor) -> log.info("Read {} rows", count));
```

**Row guards and blank rows:**
```java
reader
    .skipBlankRows()
    .stopAtBlankRows(3)
    .maxRows(100_000);
```

**Percentage progress with `countRows()`:**
```java
ExcelReader.setter(MyDto::new)
    .column((dto, cell) -> dto.setName(cell.asString()))
    .countRows()   // pre-scan to count total data rows
    .onProgress(500, (processed, cursor) -> {
        long total = cursor.getTotalRows();  // -1 if countRows() not called
        int percent = (int) (processed * 100 / total);
        log.info("{}% ({}/{})", percent, processed, total);
    })
    .read(inputStream, result -> { ... });
```

## Multi-Sheet Discovery

`ExcelWorkbookInspector` looks at a workbook without reading its rows:

```java
ExcelWorkbookInspector inspector = ExcelWorkbookInspector.create();

List<ExcelSheetInfo> sheets = inspector.sheetNames(inputStream);
sheets.forEach(s -> log.info("{}: {}", s.index(), s.name()));

List<String> headers = inspector.headers(inputStream, 0, 0);   // sheet index, header row index
```

A header that spans several rows, as written by `group(...)`, has no single row with every column name: the name of a
column without a group sits in the first row, in a cell merged downwards. Give the size of the header block, as for
`ExcelReader.headerRows(...)`. It ends at `headerRowIndex`, and for each column the lowest non-blank cell is its name:

```java
// First | (G: Grouped1, Grouped2) | No.   ->  [First, Grouped1, Grouped2, No.]
List<String> names = inspector.headers(inputStream, 0, 1, 2);   // sheet index, last header row index, header rows
```

`sheetNames` and `headers` each take an `InputStream` (copied to a temporary file, not closed), a `Path` (opened in
place) or an `InputStreamSource` (opened and closed here). Settings go on the inspector, so they do not multiply the
overloads, and one inspector can be used for any number of files.

A password-protected workbook needs `password(...)`; a wrong one fails with
`ExcelReadException: Invalid password for encrypted Excel file`:

```java
ExcelWorkbookInspector.create().password("secret").sheetNames(path);
```

For a file from an untrusted source (an upload) add `limits(...)`. The copy of a stream stops at `maxInputBytes`, a
file larger than that is refused before it is opened, and a workbook with more sheets than `maxSheets` (or a header row
wider than `maxColumns`) fails with `ReadLimitExceededException`. Without limits nothing is bounded:

```java
var inspector = ExcelWorkbookInspector.create().limits(ReadLimits.UPLOAD);
List<ExcelSheetInfo> sheets = inspector.sheetNames(upload);
```

`ExcelWorkbookReader` does this on its own when the readers it was given have a `password(...)`.

## CellData Conversion Methods

| Method | Return Type |
|--------|-------------|
| `asString()` | `String` |
| `asInt()` / `asInt(default)` | `Integer` / `int` |
| `asLong()` / `asLong(default)` | `Long` / `long` |
| `asDouble()` / `asDouble(default)` | `Double` / `double` |
| `asFloat()` | `Float` |
| `asBigDecimal()` | `BigDecimal` |
| `rawNumber()` | `BigDecimal` (stored XLSX number, null if unavailable) |
| `asBoolean()` | `boolean` (`true`/`1`/`y`/`yes`) |
| `asBooleanOrNull()` | `Boolean` |
| `asLocalDate()` | `LocalDate` |
| `asLocalDateTime()` | `LocalDateTime` |
| `asLocalTime()` | `LocalTime` |
| `asZonedDateTimeStrict(zone)` | `ZonedDateTime` (rejects gaps and overlaps) |
| `asEnum(Class<E>)` | `E` (case-insensitive) |
| `as(Function<String, R>)` | `R` (custom, null if blank) |
| `as(Function, default)` | `R` |
| `asString(default)` | `String` |
| `isEmpty()` | `boolean` |

For XLSX numeric cells, `asNumber()`, `asInt()`, `asLong()`, `asDouble()`, `asFloat()`, and
`asBigDecimal()` convert the stored number, not the displayed text: a `2.5` cell formatted as `0`
converts to `2.5`, and a `25%` cell converts to `0.25`. `formattedValue()` and `asString()` keep
the displayed text. CSV cells and XLSX text cells convert from their text.

`asInt()` and `asLong()` reject fractional, non-finite, and out-of-range values.
Date and date-time conversion with pattern strings rejects invalid calendar dates.
`asNumber()`, `asDouble()`, and `asFloat()` reject non-finite values and nonzero
values that would round to zero. Numeric conversion accepts currency and percent markers at the
edges of a value. Custom time patterns validate the parsed time strictly.
`asFloat()` rounds directly from the decimal text. The strict zoned date-time
overloads accept only local times with exactly one valid offset in the supplied zone.

**Custom conversion:**
```java
UUID id = cell.as(UUID::fromString);
UUID id = cell.as(UUID::fromString, DEFAULT_UUID);
int qty = cell.asInt(0);  // 0 if blank
```

**Number parsing locale and custom date formats:**

Conversion settings are scoped to a reader. `CellData` without an explicit config uses
`CellConversionConfig.defaults()` (the root locale and the built-in date patterns).

Excel readers use the same locale for cell display formatting and numeric
conversion. Configure the locale and extra date patterns per reader:

```java
CsvReader.<Order>mapping(row -> new Order(
        row.get("Ordered At").asLocalDate(),
        row.get("Amount").asBigDecimal()))
    .cellConversion(c -> c
        .locale(Locale.GERMANY)
        .addDateFormat("dd.MM.yyyy"))
    .read(inputStream, result -> { ... });
```

> **Text numbers are read in the root locale by default, not in the machine's.** Where a number is text (a CSV field, an
> Excel text cell), `"1,234.5"` is 1234.5 and `"2.718"` is 2.718 on every machine: a period is the decimal separator and a
> comma the grouping separator. A file written in another convention says so on the reader, for example German
> `1.234,5` with `.cellConversion(c -> c.locale(Locale.GERMANY))`; without it `"2,718"` is read as 2718. Excel readers use
> the same locale for the text of numeric cells and for month names, so by default these are English (`Jan`).
> Numeric Excel cells are read from the stored number and do not depend on the locale.
>
> A date cell shows its digits as 0-9 and its year in the Gregorian calendar, also where the JVM default locale writes
> its own digits (Arabic, Persian, Bengali, ...) or counts years differently
> (`th-TH`, `ja-JP-u-ca-japanese`). Under such a calendar the text of a date cell is the ISO date, `2024-01-05` or
> `2024-01-05 13:45:06`, whatever format the cell has.

## CSV Reading

By default, CSV reading treats backslashes as literal text and preserves line
endings inside quoted fields. Use `.escapeChar('\\')` to read legacy files
that use backslash escapes. Malformed input for the configured charset and NUL
characters raise a `CsvReadException`.

```java
CsvReader.setter(Product::new)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .read(inputStream, result -> { ... });

// Mapping mode
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(), row.get("Age").asInt()))
    .read(inputStream, result -> { ... });
```

CSV-specific options: see [CSV](csv.md).

## External workbook compatibility

Workbook-specific 1900/1904 date systems are applied to formatted numeric cells and
header extraction. ISO date, date-time, and time cells are supported across POI versions.
When XLSX omits a cell reference, the reader uses the next column in that row.
XLSX date and time serials are rendered as local calendar values independently of
the JVM or POI time zone.
In the 1900 date system, a calendar-formatted serial 60 is reported as
`1900-02-29` with its time fraction when present. Java calendar conversions reject
that date, and `rawNumber()` returns the original serial.
Formula values come from saved caches. A formula cell without
a cached result reads as blank. Path-based row reads, metadata discovery and shared-package
reads open packages read-only and preserve the caller's original file bytes.

Encrypted inputs use an owned temporary plaintext workspace. It is removed after
reading or rejection by security preflight, while caller-owned encrypted files
remain unchanged.

The offline compatibility corpus includes Excel/Mac Excel and LibreOffice files
with shared strings/formulas, sparse and merged cells, 1904 dates, cached array
formulas and encryption. See the fixture [provenance and coverage](../../kit/src/test/resources/external-workbooks/README.md)
for exact versions, source commits, checksums and limits of this coverage.
