# Reading Excel & CSV

> [Back to Index](index.md)

## Setter Mode — Mutable Objects

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

schema.excelReader(User::new, validator)
    .strictHeaders()
    .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
    .read(inputStream, user -> importUser(user), error -> log.warn("{}", error.cellErrors()));
```

## Advanced Options

**Header row index** (files with metadata rows above header):
```java
reader.headerRowIndex(2)  // 3rd row as header (0-based)
```

**Strict headers** fail before data rows when a positional or index-based column has no header:

```java
reader.strictHeaders();       // equivalent to requireHeaders()
reader.strictHeaders(false);  // default
```

**Duplicate headers** default to the first occurrence. You can choose another policy:

```java
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.FIRST); // default
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST);
reader.duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL);
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

Excel sheet discovery follows the same ownership rule: `getSheetNames(InputStream)` and
`getSheetHeaders(InputStream, ...)` consume but do not close the supplied stream, while their
`Path` and `InputStreamSource` overloads manage their own resources.

Use `readWithSummary(...)` when aggregate counts and elapsed time are needed, or
`readReport(input, maxCollectedErrors)` for a bounded error sample. Untrusted inputs can be
guarded with `limits(new ReadLimits(maxBytes, maxSheets, maxColumns, maxCellCharacters))`.
`headerPolicy(...)` provides common trim, case-insensitive, whitespace, and Unicode-normalized
matching presets. Long-running reads can use `cancellationToken(...)` and
`onReadProgress(interval, callback)` without transferring execution to a library thread.
Byte limits are enforced while copying a stream, so oversized uploads are stopped before the
entire body is materialized. `ReadLimitExceededException` exposes the limit kind, configured
value, and observed value. Summary, report, and `readWhile` APIs accept `InputStream`, `Path`,
and `InputStreamSource`; progress callbacks always receive a terminal event.
For untrusted XLSX files, `securityPolicy(ReadSecurityPolicy.STRICT)` rejects formulas and
external workbook links before sheet rows are mapped.
Strict inspection also bounds each decompressed worksheet entry, total scanned bytes, and
compression ratio. CSV callers can use `readDetected(...)` to apply sampled charset and
delimiter detection without closing the caller stream.

**Bean Validation:**
```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
ExcelReader<User> reader = ExcelReader.setter(User::new, validator);
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

```java
List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(inputStream);
sheets.forEach(s -> log.info("{}: {}", s.index(), s.name()));

List<String> headers = ExcelReader.getSheetHeaders(inputStream, 0, 0);
```

## CellData Conversion Methods

| Method | Return Type |
|--------|-------------|
| `asString()` | `String` |
| `asInt()` / `asInt(default)` | `Integer` / `int` |
| `asLong()` / `asLong(default)` | `Long` / `long` |
| `asDouble()` / `asDouble(default)` | `Double` / `double` |
| `asFloat()` | `Float` |
| `asBigDecimal()` | `BigDecimal` |
| `asBoolean()` | `boolean` (`true`/`1`/`y`/`yes`) |
| `asBooleanOrNull()` | `Boolean` |
| `asLocalDate()` | `LocalDate` |
| `asLocalDateTime()` | `LocalDateTime` |
| `asLocalTime()` | `LocalTime` |
| `asEnum(Class<E>)` | `E` (case-insensitive) |
| `as(Function<String, R>)` | `R` (custom, null if blank) |
| `as(Function, default)` | `R` |
| `asString(default)` | `String` |
| `isEmpty()` | `boolean` |

**Custom conversion:**
```java
UUID id = cell.as(UUID::fromString);
UUID id = cell.as(UUID::fromString, DEFAULT_UUID);
int qty = cell.asInt(0);  // 0 if blank
```

**Custom date formats:**
```java
CellData.addDateFormat("dd/MM/yyyy");
CellData.addDateTimeFormat("dd/MM/yyyy HH:mm");
CellData.resetDateFormats();
```

**Number parsing locale:**
```java
CellData.setDefaultLocale(Locale.US);  // default: Locale.KOREA
```

Prefer reader-scoped conversion settings in server applications:

```java
CsvReader.<Order>mapping(row -> new Order(
        row.get("Ordered At").asLocalDate(),
        row.get("Amount").asBigDecimal()))
    .cellConversion(c -> c
        .locale(Locale.GERMANY)
        .addDateFormat("dd.MM.yyyy"))
    .read(inputStream, result -> { ... });
```

## CSV Reading

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
