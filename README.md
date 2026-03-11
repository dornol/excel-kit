# excel-kit

A lightweight Java library for streaming Excel (.xlsx) and CSV generation/parsing built on Apache POI.
Designed for large datasets with minimal memory footprint, fluent DSL-style column definitions,
password-encrypted Excel export, and optional Bean Validation support.

- **Group/Artifact:** `io.github.dornol:excel-kit`
- **License:** MIT

## Features

**Excel Writing** (SXSSFWorkbook streaming)
- Fluent column DSL with type, format, alignment, and style options
- Automatic sheet splitting when row limit is reached
- Customizable header color, row height, auto-filter, and freeze panes
- Lifecycle callbacks with `SheetContext`: `beforeHeader`, `afterData`, `afterAll`
- Dropdown data validation (select list) per column
- Row-level conditional styling (background color)
- Explicit multi-sheet workbook with different data types per sheet (`ExcelWorkbook`)
- Password-encrypted Excel output
- Consume-once output via `ExcelHandler`

**Excel Reading** (SAX-based streaming)
- Header auto-detection with column mapping DSL
- Configurable header row index and sheet index
- Optional Bean Validation integration with per-row results
- Stream-based reading via `readAsStream()`
- Large file support configuration

**CSV Writing**
- Streaming write to temp file, then flush to `OutputStream`
- Proper escaping (quotes, commas, newlines)
- UTF-8 BOM for Excel compatibility
- Configurable delimiter and charset

**CSV Reading** (OpenCSV-based)
- Header auto-detection with BOM removal
- Column mapping DSL with Bean Validation support
- Configurable delimiter, charset, and header row index

## Installation

**Gradle (Kotlin DSL)**
```kotlin
dependencies {
    implementation("io.github.dornol:excel-kit:<latest-version>")
}
```

**Maven**
```xml
<dependency>
  <groupId>io.github.dornol</groupId>
  <artifactId>excel-kit</artifactId>
  <version><!-- latest-version --></version>
</dependency>
```

### Runtime Dependencies

The library declares the following as `compileOnly`. You must provide them at runtime:

| Dependency | Required For |
|------------|-------------|
| `org.apache.poi:poi-ooxml` | Excel read/write |
| `org.slf4j:slf4j-api` | Logging |
| `jakarta.validation:jakarta.validation-api` | Bean Validation (optional) |
| `com.opencsv:opencsv` | CSV reading |

You also need runtime implementations:
- **SLF4J:** e.g. `org.slf4j:slf4j-simple` or Logback
- **Bean Validation:** e.g. `org.hibernate:hibernate-validator` (if using validation)

## Quick Start

### Excel Writing

```java
record Person(long id, String name, int age) {}

var data = Stream.of(new Person(1, "Alice", 30), new Person(2, "Bob", 28));

try (ExcelWriter<Person> writer = new ExcelWriter<>()) {
    ExcelHandler handler = writer
            .column("ID", p -> p.id())
                .type(ExcelDataType.LONG)
                .alignment(HorizontalAlignment.RIGHT)
            .column("Name", p -> p.name())
            .column("Age", p -> p.age())
                .type(ExcelDataType.INTEGER)
            .write(data);

    try (var os = Files.newOutputStream(Path.of("people.xlsx"))) {
        handler.consumeOutputStream(os);
    }
}
```

### Excel Reading

```java
class User {
    String name;
    Integer age;
}

ExcelReader<User> reader = new ExcelReader<>(User::new, null);

ExcelReadHandler<User> rh = reader
        .column((u, cell) -> u.name = cell.asString())
        .column((u, cell) -> u.age = cell.asInt())
        .build(Files.newInputStream(Path.of("users.xlsx")));

rh.read(result -> {
    if (result.success()) {
        User u = result.data();
        // process user
    } else {
        System.out.println(result.messages());
    }
});
```

### CSV Writing

```java
record Row(long id, String name) {}

var rows = Stream.of(new Row(1, "Alice"), new Row(2, "Bob"));

CsvWriter<Row> csv = new CsvWriter<>();
CsvHandler ch = csv
        .column("ID", r -> r.id())
        .column("Name", r -> r.name())
        .write(rows);

try (var os = Files.newOutputStream(Path.of("rows.csv"))) {
    ch.consumeOutputStream(os);
}
```

### CSV Reading

```java
class Product {
    String name;
    Integer price;
}

CsvReader<Product> csvReader = new CsvReader<>(Product::new, null);

CsvReadHandler<Product> crh = csvReader
        .column((p, cell) -> p.name = cell.asString())
        .column((p, cell) -> p.price = cell.asInt())
        .build(Files.newInputStream(Path.of("products.csv")));

crh.read(result -> {
    if (result.success()) {
        Product p = result.data();
    }
});
```

## Advanced Usage

### Row Height

```java
new ExcelWriter<Person>()
        .rowHeight(25)                      // data row height (default: 20pt)
        .column("Name", p -> p.name())
        .write(data);
```

### Header Color

```java
// RGB values
new ExcelWriter<>(91, 155, 213, 1_000_000);

// Preset color
new ExcelWriter<>(ExcelColor.STEEL_BLUE);
```

Available presets: `WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `DARK_GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `LIGHT_ORANGE`, `LIGHT_PURPLE`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`

### Column Styling

```java
writer
    .column("Amount", p -> p.amount())
        .type(ExcelDataType.DOUBLE)
        .format("#,##0.00")
        .alignment(HorizontalAlignment.RIGHT)
        .backgroundColor(ExcelColor.LIGHT_YELLOW)
        .bold(true)
        .fontSize(12)
    .column("Status", p -> p.status())
    .write(data);
```

### Dropdown Validation

Add a dropdown (data validation) to a column so users can only select from predefined options:

```java
writer
    .column("Name", p -> p.name())
    .column("Status", p -> p.status())
        .dropdown("Active", "Inactive", "Pending")
    .write(data);
```

The dropdown is applied to all data rows across all sheets (including rollover sheets).

### Row-Level Styling

Apply conditional background colors to entire rows based on data:

```java
writer
    .rowColor(p -> p.isError() ? ExcelColor.LIGHT_RED : null)  // null = no override
    .column("Name", p -> p.name())
    .column("Status", p -> p.status())
    .write(data);
```

When a row color is set, it overrides any column-level `backgroundColor`.

### Conditional Columns

```java
boolean showAge = true;

writer
    .column("Name", p -> p.name())
    .columnIf("Age", showAge, p -> p.age())  // only added when condition is true
    .column("Email", p -> p.email())
    .write(data);
```

### Constant Columns

```java
writer
    .column("Name", p -> p.name())
    .constColumn("Source", "SYSTEM")  // same value for every row
    .write(data);
```

### Auto-Filter and Freeze Panes

```java
writer
    .autoFilter(true)    // dropdown filter on header row
    .freezePane(1)       // freeze 1 row below the header
    .column("Name", p -> p.name())
    .write(data);
```

### Lifecycle Callbacks

Callbacks allow inserting custom rows at specific points during Excel generation.

**Invocation order per sheet:**
```
beforeHeader → header → data rows → afterData
                                      ↓ (last sheet only)
                                    afterAll
```

On sheet rollover: current sheet gets `afterData` → new sheet is created → `beforeHeader` + header.

All callbacks receive a `SheetContext` parameter that provides:
- `getSheet()` — the current `SXSSFSheet`
- `getWorkbook()` — the `SXSSFWorkbook` (useful for creating CellStyles, etc.)
- `getCurrentRow()` — the first available row index for writing
- `getColumnCount()` — the number of configured columns
- `getColumnNames()` — unmodifiable list of column header names

A new `SheetContext` is created for each callback invocation, so the sheet reference is always current (even after rollover).

```java
writer
    .column("Name", p -> p.name())
    .column("Amount", p -> p.amount())
    .beforeHeader(ctx -> {
        // called on every sheet, before the header row
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
                .setCellValue("Generated: 2025-07-19");
        return ctx.getCurrentRow() + 1;  // return the next available row index
    })
    .afterData(ctx -> {
        // called on every sheet after its data rows (e.g. subtotals)
        SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
        row.createCell(0).setCellValue("Subtotal");
        return ctx.getCurrentRow() + 1;
    })
    .afterAll(ctx -> {
        // called once on the last sheet, after afterData (e.g. grand total)
        SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
        row.createCell(0).setCellValue("Grand Total");
        return ctx.getCurrentRow() + 1;
    })
    .write(data);
```

### Sheet Auto-Splitting

When the configured maximum rows per sheet is reached, a new sheet is automatically created with `beforeHeader` and header replicated.

```java
// split every 100,000 rows
new ExcelWriter<>(ExcelColor.STEEL_BLUE, 100_000);
```

### Password-Encrypted Export

```java
try (var os = Files.newOutputStream(Path.of("secret.xlsx"))) {
    handler.consumeOutputStreamWithPassword(os, "P@ssw0rd!");
}
```

### Explicit Multi-Sheet Workbook

Use `ExcelWorkbook` to write different data types to separate sheets:

```java
try (ExcelWorkbook workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    workbook.<User>sheet("Users")
        .column("Name", u -> u.getName())
        .column("Status", u -> u.getStatus(), c -> c.dropdown("Active", "Inactive"))
        .column("Age", u -> u.getAge(), c -> c.type(ExcelDataType.INTEGER))
        .rowColor(u -> u.isError() ? ExcelColor.LIGHT_RED : null)
        .write(userStream);

    workbook.<Order>sheet("Orders")
        .column("ID", o -> o.getId())
        .column("Amount", o -> o.getAmount(), c -> c.type(ExcelDataType.DOUBLE))
        .write(orderStream);

    ExcelHandler handler = workbook.finish();
    handler.consumeOutputStream(outputStream);
}
```

Each `ExcelSheetWriter` supports the same features as `ExcelWriter`:
- Column configuration via `Consumer<ColumnConfig>`: `type`, `format`, `alignment`, `backgroundColor`, `bold`, `fontSize`, `width`, `minWidth`, `maxWidth`, `dropdown`
- `beforeHeader()`, `afterData()`, `autoFilter()`, `freezePane()`, `rowColor()`, `constColumn()`

> **Note:** Unlike `ExcelWriter`, `ExcelSheetWriter` does not auto-split sheets.
> Each `sheet()` call creates exactly one sheet.

### Cursor Access

The `Cursor` provides positional information during streaming. Use `ExcelRowFunction` (with cursor) instead of `Function`:

```java
writer
    .column("No.", (row, cursor) -> cursor.getCurrentTotal())  // row number
        .type(ExcelDataType.LONG)
    .column("Name", (row, cursor) -> row.name())
    .write(data);
```

- `cursor.getCurrentTotal()` — total rows written (across all sheets, `long`)
- `cursor.getRowOfSheet()` — current row index within the sheet

### Excel Reading — Advanced Options

**Header row index** (for files with metadata rows above the header):
```java
new ExcelReader<>(User::new, null)
        .headerRowIndex(2)  // use the 3rd row as header (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .build(inputStream);
```

**Specific sheet:**
```java
new ExcelReader<>(User::new, null)
        .sheetIndex(1)  // read the 2nd sheet (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .build(inputStream);
```

**Skip columns:**
```java
reader
    .column((u, cell) -> u.name = cell.asString())
    .skipColumn()       // skip one column
    .skipColumns(2)     // skip two more columns
    .column((u, cell) -> u.age = cell.asInt())
    .build(inputStream);
```

**Stream-based reading:**
```java
try (Stream<ReadResult<User>> stream = rh.readAsStream()) {
    stream.filter(ReadResult::success)
          .map(ReadResult::data)
          .forEach(user -> { /* process */ });
}
```

**Large file support:**
```java
// call once at application startup
ExcelReader.configureLargeFileSupport();
```

**Bean Validation:**
```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
ExcelReader<User> reader = new ExcelReader<>(User::new, validator);
```

### CellData Conversion Methods

When reading Excel/CSV, `CellData` provides type-safe conversions:

| Method | Return Type |
|--------|-------------|
| `asString()` | `String` |
| `asInt()` | `Integer` |
| `asLong()` | `Long` |
| `asDouble()` | `Double` |
| `asFloat()` | `Float` |
| `asBigDecimal()` | `BigDecimal` |
| `asBoolean()` | `boolean` (`true`/`1`/`y`/`yes`) |
| `asBooleanOrNull()` | `Boolean` |
| `asLocalDate()` | `LocalDate` |
| `asLocalDateTime()` | `LocalDateTime` |
| `asLocalTime()` | `LocalTime` |
| `isEmpty()` | `boolean` |

Custom date formats:
```java
CellData.addDateFormat("dd/MM/yyyy");
CellData.addDateTimeFormat("dd/MM/yyyy HH:mm");
CellData.resetDateFormats();      // restore defaults
CellData.resetDateTimeFormats();
```

Number parsing locale (default: `Locale.KOREA`):
```java
CellData.setDefaultLocale(Locale.US);
```

### ExcelDataType Reference

| Type | Java Type | Default Format |
|------|-----------|---------------|
| `STRING` | String | — |
| `BOOLEAN_TO_YN` | Boolean → "Y"/"N" | — |
| `LONG` | Long | `#,##0` |
| `INTEGER` | Integer | `#,##0` |
| `DOUBLE` | Double | `#,##0.00` |
| `FLOAT` | Float | `#,##0.00` |
| `DOUBLE_PERCENT` | Double | `0.00%` |
| `FLOAT_PERCENT` | Float | `0.00%` |
| `DATETIME` | LocalDateTime | `yyyy-mm-dd hh:mm:ss` |
| `DATE` | LocalDate | `yyyy-mm-dd` |
| `TIME` | LocalTime | `hh:mm:ss` |
| `BIG_DECIMAL_TO_DOUBLE` | BigDecimal | `#,##0.00` |
| `BIG_DECIMAL_TO_LONG` | BigDecimal | `#,##0` |

### ExcelDataFormat Presets

Use with `.format(ExcelDataFormat.NUMBER.getFormat())`:

| Preset | Format String |
|--------|---------------|
| `NUMBER` | `#,##0` |
| `NUMBER_1` | `#,##0.0` |
| `NUMBER_2` | `#,##0.00` |
| `NUMBER_4` | `#,##0.0000` |
| `PERCENT` | `0.00%` |
| `DATETIME` | `yyyy-mm-dd hh:mm:ss` |
| `DATE` | `yyyy-mm-dd` |
| `TIME` | `hh:mm:ss` |
| `CURRENCY_KRW` | `#,##0"원"` |
| `CURRENCY_USD` | `"$"#,##0.00` |

### ExcelKitSchema — Unified Read/Write Definitions

Define columns once, use for both reading and writing:

```java
ExcelKitSchema<Book> schema = ExcelKitSchema.<Book>builder()
        .column("Title",  Book::getTitle,  (b, cell) -> b.setTitle(cell.asString()))
        .column("Price",  Book::getPrice,  (b, cell) -> b.setPrice(cell.asInt()))
        .build();

// writing
ExcelHandler handler = schema.excelWriter().write(bookStream);

// reading
ExcelReadHandler<Book> rh = schema.excelReader(Book::new, null)
        .build(inputStream);
```

Also supports CSV:
```java
CsvHandler ch = schema.csvWriter().write(bookStream);
CsvReadHandler<Book> crh = schema.csvReader(Book::new, null).build(inputStream);
```

### CSV Options

```java
new CsvWriter<Row>()
    .delimiter('\t')                   // tab-separated
    .charset(StandardCharsets.UTF_16)  // custom encoding
    .bom(false)                        // disable UTF-8 BOM
    .column("Name", r -> r.name())
    .write(rows);

new CsvReader<>(Row::new, null)
    .delimiter('\t')
    .charset(StandardCharsets.UTF_16)
    .headerRowIndex(1)
    .column((r, cell) -> r.name = cell.asString())
    .build(inputStream);
```

## Exception Handling

| Exception | Description |
|-----------|-------------|
| `ExcelKitException` | Base class for all library exceptions |
| `ExcelWriteException` | Excel write errors (no columns, handler already consumed, etc.) |
| `ExcelReadException` | Excel read/parse errors |
| `CsvWriteException` | CSV write errors |
| `CsvReadException` | CSV read/parse errors |

- Column mapping exceptions are safely logged and fall back to string conversion (Excel writing).
- Calling `consumeOutputStream` on an already-consumed handler throws the corresponding `WriteException`.
- Empty password on encrypted export throws `IllegalArgumentException`.

## Requirements

- **JDK 17+**
- Apache POI 5.x for Excel operations

## Build & Test

```bash
./gradlew build
./gradlew test
```

## License

MIT License. See [LICENSE](./LICENSE) for details.
