# excel-kit — Reading

> Other topics: [Index](../AI.md) | [Quick Start](quick-start.md) | [Column Config](column-config.md) | [Advanced](advanced.md) | [CSV](csv.md)

## Three Read Modes

### 1. Setter Mode (Mutable Objects)
```java
new ExcelReader<>(User::new, validator)  // validator is optional (null to skip)
    .column("Name", (u, cell) -> u.name = cell.asString())
    .column("Age", (u, cell) -> u.age = cell.asInt())
    .build(inputStream)
    .read(result -> {
        if (result.success()) { User u = result.data(); }
        else { System.out.println(result.messages()); }
    });
```

### 2. Mapping Mode (Records / Immutable Objects)
```java
ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
    row.get("Name").asString(),
    row.get("Age").asInt(),
    row.get("City").asString()
)).build(inputStream).read(result -> { ... });
```

### 3. Map Mode (No POJO)
```java
ExcelReader.forMap()
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
        String name = row.get("Name");
    });
```

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
| `asString()` | String | Raw string value |
| `asInt()` | Integer | Parse as integer (null if empty) |
| `asLong()` | Long | Parse as long |
| `asDouble()` | Double | Parse as double |
| `asFloat()` | Float | Parse as float |
| `asBigDecimal()` | BigDecimal | Parse as BigDecimal |
| `asBoolean()` | Boolean | Parse as boolean |
| `asLocalDate()` | LocalDate | Parse as date |
| `asLocalDateTime()` | LocalDateTime | Parse as datetime |
| `asZonedDateTime(ZoneId)` | ZonedDateTime | Parse with timezone |
| `as(Function<String, T>)` | T | Custom conversion (e.g., `UUID::fromString`) |
| `asInt(defaultValue)` | int | With default for null/empty |
| `asString(defaultValue)` | String | With default for null/empty |
| `isEmpty()` | boolean | Check if cell is empty |
| `value()` | String | Raw value (nullable) |

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
    .build(inputStream);
```

With `countRows()`, `cursor.getTotalRows()` returns the total data row count in the progress callback (otherwise `-1`).

### Multi-row headers (v0.16.13+, Excel)

Files written with multi-level `group(...)` have blank column-header cells due to vertical merges.
`headerRows(int)` combines N header rows per column, taking the bottom-most non-blank value:

```java
ExcelReader.<Record>mapping(row -> ...)
    .headerRowIndex(1)     // last header row is row index 1
    .headerRows(2)         // 2 header rows (group row + column header row)
    .build(in)
    .read(result -> ...);
```

Default `headerRows(1)` = existing single-row behavior. Empty-string headers preserved.

## Read Methods

| Method | Description |
|--------|-------------|
| `.read(Consumer<ReadResult<T>>)` | Process each row (skips failures) |
| `.read(Consumer<T> onSuccess, Consumer<RowError> onError)` | Split success/error callbacks (v0.16.12+) |
| `.readStrict(Consumer<ReadResult<T>>)` | Throws on first validation failure |
| `.readAsStream()` | Returns `Stream<ReadResult<T>>` |

Read handlers are one-shot. `read()`, `readStrict()`, and `readAsStream()` all
consume the handler and its temporary resources. Build a new handler from a new
`InputStream` if you need to read the same file again.

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

`ReadResult<T>.cause()` is also available for fail-paths in the unified `read(Consumer<ReadResult<T>>)` form.

## Multi-Sheet Discovery

```java
List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(inputStream);
// ExcelSheetInfo: index(), name()

List<String> headers = ExcelReader.getSheetHeaders(inputStream, sheetIndex, headerRowIndex);
```

## Bean Validation

```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

// Setter mode
new ExcelReader<>(User::new, validator).column(...).build(in).read(result -> {
    if (!result.success()) {
        // result.messages() contains validation errors
    }
});

// Mapping mode
ExcelReader.mapping(row -> new Person(...), validator).build(in);
```

## CSV Reading

Same API pattern as Excel reading:

```java
// Setter mode
new CsvReader<>(Product::new, null)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .build(inputStream).read(result -> { ... });

// Mapping mode
CsvReader.<Product>mapping(row -> new Product(
    row.get("Name").asString(), row.get("Price").asInt()
)).build(inputStream).read(result -> { ... });

// Map mode
CsvReader.forMap().build(inputStream).read(result -> {
    Map<String, String> row = result.data();
});
```
