# excel-kit Guide

Fluent API for streaming Excel (.xlsx) and CSV generation/parsing in Java 17+.
For installation and project overview, see [README](../../README.md).

Unlike annotation-based libraries, excel-kit defines column mappings in the adapter layer
using a fluent API — not on domain objects. Your DTOs and domain models stay free of
infrastructure concerns.

## Quick Navigation

| Document | Covers | Key APIs |
|----------|--------|----------|
| [Writing](writing.md) | Column DSL, types, formats, callbacks, auto-split | `ExcelWriter.create()`, `column()`, `write()`, `ExcelDataType` |
| [Reading](reading.md) | Setter/Mapping/Map mode, CellData, stream reading | `ExcelReader.setter()`, `mapping()`, `forMap()`, `CellData` |
| [Styling](styling.md) | Font, color, border, alignment, rotation | `fontColor()`, `border()`, `backgroundColor()`, `bold()` |
| [Headers](headers.md) | Header color/font, group headers, comments | `headerColor()`, `group()`, `headerComment()`, `groupComment()` |
| [Validation](validation.md) | Dropdown, data validation, conditional formatting | `dropdown()`, `validation()`, `conditionalFormatting()` |
| [Layout](layout.md) | Freeze, filter, outline, print, summary rows | `freezeRows()`, `autoFilter()`, `outline()`, `summary()` |
| [Content](content.md) | Formula, hyperlink, rich text, image, chart | `ExcelDataType.FORMULA`, `HYPERLINK`, `RICH_TEXT`, `IMAGE` |
| [Multi-Sheet](multi-sheet.md) | ExcelWorkbook, auto-rollover, tab color | `ExcelWorkbook.create()`, `maxRows()`, `tabColor()` |
| [Protection](protection.md) | Sheet/workbook protection, password encryption | `protectSheet()`, `protectWorkbook()`, `password()` |
| [CSV](csv.md) | CSV write/read, dialect, quoting | `CsvWriter.create()`, `CsvReader.setter()`, `dialect()` |
| [Spring](spring.md) | MVC, WebFlux integration | `DownloadResponse.excel()`, `StreamingResponseBody` |
| [Reference](reference.md) | Data types, formats, schema, exceptions, notes | `ExcelDataType`, `ExcelDataFormat`, `ExcelKitSchema` |

---

## Write Excel — Quick Example

```java
ExcelWriter.<Product>create().headerColor(ExcelColor.STEEL_BLUE)
    .column("Name", Product::name)
    .column("Price", Product::price, c -> c
        .type(ExcelDataType.INTEGER).format("#,##0")
        .alignment(HorizontalAlignment.RIGHT))
    .column("Released", Product::released, c -> c.type(ExcelDataType.DATE))
    .autoFilter(true).freezeRows(1)
    .write(productStream)
    .writeTo(Path.of("products.xlsx"));
```

> Details: [Writing](writing.md)

## Read Excel — Quick Example

**Mapping mode** (records / immutable objects):
```java
ExcelReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(),
        row.get("Age").asInt()))
    .build(inputStream)
    .read(result -> {
        if (result.success()) { Person p = result.data(); }
    });
```

**Setter mode** (mutable objects):
```java
ExcelReader.setter(User::new)
    .column("Name", (u, cell) -> u.setName(cell.asString()))
    .column("Age", (u, cell) -> u.setAge(cell.asInt())).required()
    .build(inputStream)
    .read(result -> { ... });
```

> Details: [Reading](reading.md)

## Column Styling — Quick Example

```java
.column("Price", Product::price, cfg -> cfg
    .type(ExcelDataType.INTEGER)
    .bold(true).fontColor(ExcelColor.RED)
    .backgroundColor(ExcelColor.LIGHT_YELLOW)
    .border(ExcelBorderStyle.THIN)
    .alignment(HorizontalAlignment.RIGHT))
```

> Details: [Styling](styling.md)

## Group Headers — Quick Example

```java
writer
    .column("Name", Row::name)
    .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
    .column("Q2", Row::q2, c -> c.group("Financial", "Revenue"))
    .column("Profit", Row::profit, c -> c.group("Financial"))
    .write(data);
// Produces:
// | Name |        Financial         |
// | Name |   Revenue    |  Profit   |
// | Name |  Q1  |  Q2   |  Profit   |
```

> Details: [Headers](headers.md)

## Validation — Quick Example

```java
.column("Status", p -> p.status(), c -> c.dropdown("Active", "Inactive", "Pending"))
.column("Age", p -> p.age(), c -> c.validation(ExcelValidation.integerBetween(0, 150)))
```

> Details: [Validation](validation.md)

## Multi-Sheet — Quick Example

```java
try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users").column("Name", User::name).write(userStream);
    wb.<Order>sheet("Orders").column("Amount", Order::amount).write(orderStream);
    wb.finish().writeTo(Path.of("report.xlsx"));
}
```

> Details: [Multi-Sheet](multi-sheet.md)

## CSV — Quick Example

```java
// Write
CsvWriter.<Product>create()
    .column("Name", Product::name)
    .column("Price", Product::price)
    .write(productStream).writeTo(Path.of("products.csv"));

// Read
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(), row.get("Age").asInt()))
    .build(inputStream).read(result -> { ... });
```

> Details: [CSV](csv.md)

## Spring MVC — Quick Example

```java
@GetMapping("/download")
public ResponseEntity<StreamingResponseBody> download() {
    ExcelHandler handler = writer.write(dataStream);
    return DownloadResponse.excel(handler, "report");
}
```

> Details: [Spring](spring.md)

## Document Properties — Quick Example

```java
ExcelWriter.<Product>create()
    .documentProperty("title", "Sales Report Q4")
    .documentProperty("author", "Finance Team")
    .column("Name", Product::name)
    .write(data);
```

> Details: [Writing](writing.md)

## Header Style — Quick Example

```java
writer
    .headerColor(ExcelColor.STEEL_BLUE)
    .headerStyle(h -> h.alignment(HorizontalAlignment.LEFT).border(ExcelBorderStyle.MEDIUM))
    .headerFontName("Arial").headerFontSize(14)
```

> Details: [Headers](headers.md)

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
