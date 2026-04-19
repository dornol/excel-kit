# Reading Excel & CSV

> [Back to Index](index.md)

## Setter Mode — Mutable Objects

```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.name = cell.asString()).required()
    .column("Age", (u, cell) -> u.age = cell.asInt())
    .build(inputStream)
    .read(result -> {
        if (result.success()) {
            User u = result.data();
        } else {
            System.out.println(result.messages());
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
    .build(inputStream)
    .read(result -> {
        if (result.success()) {
            PersonRecord p = result.data();
        }
    });
```

Columns are matched by header name — order in the file doesn't matter.

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
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
        String name = row.get("Name");
    });

// CSV equivalent
CsvReader.forMap().build(inputStream).read(result -> { ... });
```

## Column Matching

### Name-Based (default)

```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.name = cell.asString())
    .column("City", (u, cell) -> u.city = cell.asString())  // other columns ignored
    .build(inputStream);
```

### Index-Based

```java
ExcelReader.setter(User::new)
    .columnAt(0, (u, cell) -> u.name = cell.asString())
    .columnAt(2, (u, cell) -> u.city = cell.asString())
    .columnAt(4, (u, cell) -> u.phone = cell.asString())
    .build(inputStream);
```

Can be mixed with name-based mapping.

### Positional with Skip

```java
reader
    .column((u, cell) -> u.name = cell.asString())
    .skipColumn()
    .skipColumns(2)
    .column((u, cell) -> u.age = cell.asInt())
    .build(inputStream);
```

## Required Columns

```java
ExcelReader.setter(User::new)
    .column("Name", (u, c) -> u.setName(c.asString())).required()
    .column("Email", (u, c) -> u.setEmail(c.asString())).required()
    .column("Nickname", (u, c) -> u.setNick(c.asString()))  // optional
    .build(inputStream)
    .read(result -> {
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
- `type()` — `VALIDATION` or `MAPPING`
- `messages()` — human-readable messages (list)
- `cause()` — nullable throwable (for `MAPPING` errors)

## Advanced Options

**Header row index** (files with metadata rows above header):
```java
reader.headerRowIndex(2)  // 3rd row as header (0-based)
```

**Multi-row headers** (v0.16.13+, Excel only):

For files with multi-level group headers, use `headerRows(int)` to combine N header rows.
Takes the bottom-most non-blank value per column.

```java
ExcelReader.<Row>mapping(row -> new Row(
        row.get("Q1").asInt(), row.get("Q2").asInt(), row.get("Profit").asInt()))
    .headerRowIndex(2)  // last header row (0-based)
    .headerRows(3)      // 3 rows: 2 group + 1 column header
    .build(in).read(result -> ...);
```

**Specific sheet:**
```java
reader.sheetIndex(1)  // 2nd sheet (0-based)
```

**Stream-based reading:**
```java
try (Stream<ReadResult<User>> stream = handler.readAsStream()) {
    stream.filter(ReadResult::success)
          .map(ReadResult::data)
          .forEach(this::process);
}
```

> `readAsStream()` holds resources — always use try-with-resources.

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

## Multi-Sheet Discovery

```java
List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(inputStream);
sheets.forEach(s -> System.out.println(s.index() + ": " + s.name()));

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

## CSV Reading

```java
CsvReader.setter(Product::new)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .build(inputStream)
    .read(result -> { ... });

// Mapping mode
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(), row.get("Age").asInt()))
    .build(inputStream).read(result -> { ... });
```

CSV-specific options: see [CSV](csv.md).
