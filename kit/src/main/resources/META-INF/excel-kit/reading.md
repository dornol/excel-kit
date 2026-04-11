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
    .headerRowIndex(0)      // default: 0
    .onProgress(10_000, (count, cursor) -> log.info("Read {} rows", count))
    .build(inputStream);
```

## Read Methods

| Method | Description |
|--------|-------------|
| `.read(Consumer<ReadResult<T>>)` | Process each row (skips failures) |
| `.readStrict(Consumer<ReadResult<T>>)` | Throws on first validation failure |
| `.readAsStream()` | Returns `Stream<ReadResult<T>>` |

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
