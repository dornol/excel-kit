# excel-kit

A lightweight Java library for streaming Excel (.xlsx) and CSV generation/parsing built on Apache POI.
Designed for large datasets with minimal memory footprint.

- **Group/Artifact:** `io.github.dornol:excel-kit`
- **License:** MIT
- **JDK:** 17+

## Why excel-kit?

Column mappings live in the adapter layer via a fluent API — not on domain objects.
No `@ExcelColumn` annotations, no infrastructure leaking into your models.

```java
ExcelWriter.<User>create()
    .column("Name", User::name)
    .column("Age", User::age, c -> c.type(ExcelDataType.INTEGER))
    .write(userStream)
    .writeTo(Path.of("users.xlsx"));
```

## Installation

**Gradle**
```kotlin
implementation("io.github.dornol:excel-kit:0.18.1")
```

**Maven**
```xml
<dependency>
  <groupId>io.github.dornol</groupId>
  <artifactId>excel-kit</artifactId>
  <version>0.18.1</version>
</dependency>
```

Runtime dependencies (declared as `compileOnly`):

| Dependency | Required For |
|------------|-------------|
| `org.apache.poi:poi-ooxml` | Excel read/write |
| `org.slf4j:slf4j-api` | Logging |
| `com.opencsv:opencsv` | CSV reading |
| `jakarta.validation:jakarta.validation-api` | Bean Validation (optional) |

## Quick Start

### Write Excel

```java
record Product(String name, int price, LocalDate released) {}

ExcelWriter.<Product>create().headerColor(ExcelColor.STEEL_BLUE)
    .column("Name", Product::name)
    .column("Price", Product::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
    .column("Released", Product::released, c -> c.type(ExcelDataType.DATE))
    .autoFilter(true)
    .freezeRows(1)
    .write(productStream)
    .writeTo(Path.of("products.xlsx"));
```

### Read Excel

**Mapping mode** — for records / immutable objects:

```java
record Person(String name, int age) {}

ExcelReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(),
        row.get("Age").asInt()
)).build(inputStream).read(result -> {
    if (result.success()) {
        Person p = result.data();
    }
});
```

**Setter mode** — for mutable objects:

```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.setName(cell.asString()))
    .column("Age", (u, cell) -> u.setAge(cell.asInt())).required()
    .build(inputStream)
    .read(result -> { ... });
```

Header aliases, strict header validation, and duplicate header policy are available for Excel and CSV:

```java
ExcelReader.setter(User::new)
    .strictHeaders()
    .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
    .column(List.of("Name", "User Name", "이름"), (u, cell) -> u.setName(cell.asString()))
    .build(inputStream)
    .read(user -> { ... }, error -> log.warn("file row {}", error.fileRowNum()));
```

Read failures include structured cell-level details that are useful for APIs and
upload screens:

```java
reader.read(
    user -> importUser(user),
    error -> error.cellErrors().forEach(cell ->
        log.warn("row {}, header {}, value {}: {}",
            error.fileRowNum(), cell.headerName(), cell.cellValue(), cell.message()))
);
```

For JSON responses, keep those details as fields instead of flattening them into
one message:

```json
{
  "fileRowNum": 2,
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

Schema-based readers can carry the same read metadata:

```java
ExcelKitSchema.<User>builder()
    .requiredColumn("Name", List.of("User Name", "이름"),
            User::getName, (u, cell) -> u.setName(cell.asString()))
    .build();
```

### Write & Read CSV

```java
// Write
CsvWriter.<Product>create()
    .column("Name", Product::name)
    .column("Price", Product::price)
    .write(productStream)
    .writeTo(Path.of("products.csv"));

// Read
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(),
        row.get("Age").asInt()
)).build(inputStream).read(result -> { ... });
```

### Map-Based I/O (no POJO needed)

```java
// Write
ExcelWriter.forMap("Name", "Age", "City")
    .write(Stream.of(Map.of("Name", "Alice", "Age", 30, "City", "Seoul")))
    .writeTo(Path.of("output.xlsx"));

// Read
ExcelReader.forMap()
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
    });
```

### Multi-Sheet Workbook

```java
try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users")
        .column("Name", User::name)
        .column("Email", User::email)
        .write(userStream);

    wb.<Order>sheet("Orders")
        .column("ID", Order::id, c -> c.type(ExcelDataType.LONG))
        .column("Amount", Order::amount, c -> c.type(ExcelDataType.DOUBLE))
        .write(orderStream);

    wb.finish().writeTo(Path.of("report.xlsx"));
}
```

### Spring MVC

```java
@GetMapping("/download")
public ResponseEntity<StreamingResponseBody> download() {
    ExcelHandler handler = writer.write(dataStream);
    return ResponseEntity.ok()
        .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"report.xlsx\"")
        .body(handler::writeTo);
}
```

## Features at a Glance

**Writing (Excel & CSV)**

| Category | Highlights |
|----------|-----------|
| Column DSL | Type, format, alignment, width, font, color, rotation, indentation |
| Styling | Row/cell conditional colors, conditional row style (bold/font/color), borders (per-side), bold, strikethrough, underline |
| Layout | Auto-filter, freeze panes (row + column), multi-level group headers, column outline, hiding |
| Header | Global/per-column header color, per-column header background, header row height, header style config, group header comments, row-number column |
| Validation | Dropdown, integer/decimal ranges, text length, date ranges, custom formulas |
| Content | Formulas, hyperlinks, rich text, images (byte[] or URL, custom size), cell comments |
| Charts | BAR, LINE, PIE, SCATTER, AREA, DOUGHNUT |
| Protection | Sheet/workbook protection, per-column lock, password encryption (AES-256) |
| Multi-sheet | `ExcelWorkbook` for typed multi-sheet, auto-rollover via `maxRows()` |
| Callbacks | `beforeHeader`, `afterData`, `afterAll` with `SheetContext` |
| Null handling | `nullValue()` default for null cells, `defaultStyle()` for writer-level defaults |
| Metadata | Document properties (title, author, subject, keywords), named ranges (fluent API) |
| Other | Summary rows (SUM/AVG/COUNT), conditional formatting, data bars, icon sets, print setup |

**Reading (Excel & CSV)**

| Category | Highlights |
|----------|-----------|
| Column matching | By name, aliases, by index, positional with skip |
| Read modes | Setter (mutable), Mapping (records), Map (schema-less) |
| Headers | Strict header validation, duplicate header policies, single or multi-row headers (`headerRows(int)`, Excel) |
| Validation | Bean Validation, `required()` per column |
| Callbacks | Unified `read(Consumer<ReadResult>)` or split `read(onSuccess, onError)` with typed `RowError`, `CellError`, and physical file row number |
| Stream | `readAsStream()` with lazy evaluation, `readStrict()` for fail-fast |
| Discovery | `getSheetNames()`, `getSheetHeaders()` |
| Config | Sheet index, header row index, progress callback, `countRows()` for total row pre-scan, password-encrypted files |

Read handlers are one-shot: after calling `read()`, `readStrict()`, or `readAsStream()`,
create a new handler for another pass over the same file.

## Performance

Measured on Apple M-series, JDK 21. Pure write throughput (excludes DB/network I/O).

| Scenario | Rows | Time | Throughput |
|----------|------|------|------------|
| Excel 5 cols | 1M | 3.3s | ~300K rows/s |
| Excel 50 cols | 100K | 2.9s | ~34K rows/s |
| CSV 5 cols | 1M | 0.45s | ~2.2M rows/s |

```bash
./gradlew :kit:benchmark
```

## Documentation

For the full feature guide (all column options, styling, validation, Spring WebFlux integration,
security details, migration guide, etc.), see:

- **[Full Guide](docs/guide/index.md)** — complete API reference with code examples
- **[Changelog](CHANGELOG.md)** — version history
- **[Example App](example/)** — Spring Boot showcase with all features

## AI / LLM Integration

**In-JAR docs** — `META-INF/excel-kit/` is included in the published artifact.
AI agents with classpath access (IDE plugins, code assistants) can read these
files directly without network requests.

**Remote docs** — for HTTP-based agents, start with the
[guide index](docs/guide/index.md). It contains a summary table with key APIs
per topic, plus inline code examples that cover most common use cases.
Follow links to individual topic files only when you need full details.

| Source | Path | When to use |
|--------|------|-------------|
| JAR | `META-INF/excel-kit/*.md` | IDE-integrated agents (classpath access) |
| GitHub | `docs/guide/index.md` | HTTP-based agents (fetch via raw URL) |
| GitHub | `docs/guide/{topic}.md` | Deep-dive into specific feature areas |

## Quick Reference

| Task | Entry Point |
|------|------------|
| Write Excel (typed) | `ExcelWriter.<T>create().column(...).write(stream)` |
| Write Excel (map) | `ExcelWriter.forMap("A", "B").write(stream)` |
| Write Excel (multi-sheet) | `ExcelWorkbook.create()` → `.sheet("name")` |
| Write CSV | `CsvWriter.<T>create().column(...).write(stream)` |
| Read Excel (setter) | `ExcelReader.setter(T::new).column(...).build(in).read(...)` |
| Read Excel (mapping) | `ExcelReader.mapping(row -> ...).build(in).read(...)` |
| Read Excel (map) | `ExcelReader.forMap().build(in).read(...)` |
| Read CSV | `CsvReader.setter(T::new).column(...).build(in).read(...)` |

## Build & Test

```bash
./gradlew build
./gradlew test
```

## License

MIT License. See [LICENSE](./LICENSE) for details.
