# Writing Excel & CSV

> [Back to Index](index.md)

## Basic Excel Writing

```java
record Person(long id, String name, int age) {}

var data = Stream.of(new Person(1, "Alice", 30), new Person(2, "Bob", 28));

ExcelHandler handler = ExcelWriter.<Person>create()
    .column("ID", p -> p.id(), c -> c.type(ExcelDataType.LONG).alignment(HorizontalAlignment.RIGHT))
    .column("Name", p -> p.name())
    .column("Age", p -> p.age(), c -> c.type(ExcelDataType.INTEGER))
    .write(data);

handler.writeTo(Path.of("people.xlsx"));
// or: handler.writeTo(outputStream);
```

## Basic CSV Writing

```java
CsvHandler ch = CsvWriter.<Row>create()
    .column("ID", r -> r.id())
    .column("Name", r -> r.name())
    .write(rows);

ch.writeTo(Path.of("rows.csv"));
```

## Data Types

Set via `.type(ExcelDataType.XXX)`:

| Type | Java Type | Default Format |
|------|-----------|----------------|
| `STRING` (default) | String | — |
| `INTEGER` | Number | `#,##0` |
| `LONG` | Number | `#,##0` |
| `DOUBLE` | Number | `#,##0.00` |
| `DOUBLE_PERCENT` | Number | `0.00%` |
| `DATE` | LocalDate | `yyyy-mm-dd` |
| `DATETIME` | LocalDateTime | `yyyy-mm-dd hh:mm:ss` |
| `FORMULA` | String (formula) | — |
| `HYPERLINK` | String or ExcelHyperlink | — |
| `RICH_TEXT` | ExcelRichText | — |
| `IMAGE` | ExcelImage | — |
| `BOOLEAN` | Boolean | — |

Custom format: `.format("#,##0.00")` or use `ExcelDataFormat` presets.

Full type table: [Reference](reference.md#excelDataType-reference)

## Column DSL

```java
.column("Price", Product::price, cfg -> cfg
    .type(ExcelDataType.INTEGER)
    .format("#,##0")
    .alignment(HorizontalAlignment.RIGHT)
    .backgroundColor(ExcelColor.LIGHT_YELLOW)
    .bold(true)
    .fontSize(12))
```

All column configuration methods are documented in [Styling](styling.md).

## Map-Based Writing

Write `Map<String, Object>` data without typed POJOs:

```java
ExcelWriter.forMap("Name", "Age", "City")
    .write(Stream.of(
        Map.of("Name", "Alice", "Age", 30, "City", "Seoul"),
        Map.of("Name", "Bob", "Age", 25, "City", "Tokyo")))
    .writeTo(outputStream);

// CSV equivalent
CsvWriter.forMap("Name", "Age").write(stream).writeTo(outputStream);
```

## Conditional Columns

```java
writer
    .column("Name", p -> p.name())
    .columnIf("Age", showAge, p -> p.age())  // only when showAge == true
    .columnIf("Score", showScore, p -> p.score(), c -> c.type(ExcelDataType.INTEGER))
    .column("Email", p -> p.email())
    .write(data);
```

## Constant Columns

```java
writer
    .constColumn("Source", "SYSTEM")  // same value for every row
    .constColumn("Version", "v2", c -> c.fontColor(ExcelColor.GRAY))
    .constColumnIf("Debug", isDebug, "true")  // conditional
    .write(data);
```

## Row Number Column (v0.16.11+)

```java
writer
    .rowNumberColumn("No.")  // 1-based sequential, works across auto-rollover sheets
    .column("Name", Product::name)
    .write(data);
```

Equivalent to: `.column("No.", (r, cursor) -> cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))`

## Cursor Access

`ExcelRowFunction` provides positional information during streaming:

```java
writer
    .column("No.", (row, cursor) -> cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))
    .column("Name", (row, cursor) -> row.name())
    .write(data);
```

- `cursor.getCurrentTotal()` — total rows written across all sheets (`long`)
- `cursor.getRowOfSheet()` — current row index within the sheet

## Null Value Defaults

```java
// Per-column
writer.column("Status", Item::getStatus, c -> c.nullValue("N/A"));

// Writer-level default (all columns inherit)
writer.defaultStyle(d -> d.nullValue("-"))
    .column("Name", Item::getName)              // null -> "-"
    .column("Status", Item::getStatus, c -> c.nullValue("N/A"))  // null -> "N/A"
    .write(data);
```

## Auto Width

Column widths are auto-calculated from the first N data rows:

```java
writer.autoWidthSampleRows(200)  // sample 200 rows (default: 100)
writer.autoWidthSampleRows(0)    // disable auto-width
```

## Row Height

```java
writer.rowHeight(25);  // data row height (default: 20pt)
```

## Lifecycle Callbacks

Callbacks allow inserting custom rows at specific points:

```
beforeHeader -> header -> data rows -> afterData -> afterAll (last sheet only)
```

On sheet rollover: current sheet gets `afterData` -> new sheet -> `beforeHeader` + header.

```java
writer
    .beforeHeader(ctx -> {
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
            .setCellValue("Generated: 2025-07-19");
        return ctx.getCurrentRow() + 1;
    })
    .afterData(ctx -> {
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
            .setCellValue("Subtotal");
        return ctx.getCurrentRow() + 1;
    })
    .afterAll(ctx -> {
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
            .setCellValue("Grand Total");
        return ctx.getCurrentRow() + 1;
    })
    .write(data);
```

`SheetContext` provides:
- `getSheet()`, `getWorkbook()` — POI objects
- `getCurrentRow()` — first available row index
- `getColumnCount()`, `getColumnNames()` — column metadata
- `columnLetter(int)` — static helper (0->"A", 26->"AA")
- `mergeCells(firstRow, lastRow, firstCol, lastCol)` or `mergeCells("A1:C3")`
- `groupRows(firstRow, lastRow)` / `groupRows(firstRow, lastRow, collapsed)`
- `namedRange(name, reference)` / `namedRange(name, colIdx, firstRow, lastRow)`

## Cell Merging

```java
writer.beforeHeader(ctx -> {
    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
        .setCellValue("Report Title");
    ctx.mergeCells(0, 0, 0, 2);  // merge first row across 3 columns
    // or: ctx.mergeCells("A1:C1");
    return ctx.getCurrentRow() + 1;
});
```

## Sheet Auto-Splitting

```java
ExcelWriter.create().maxRows(100_000);  // auto-split at 100K rows per sheet
```

## Progress Callback

```java
writer.onProgress(10_000, (count, cursor) -> log.info("Processed {} rows", count));
```

> The callback runs on the writing thread — keep it fast and non-blocking.

## Document Properties (v0.16.14+)

Set Excel document metadata (visible in File > Properties):

```java
ExcelWriter.<Product>create()
    .documentProperty("title", "Sales Report Q4")
    .documentProperty("author", "Finance Team")
    .documentProperty("keywords", "sales,revenue,2024")
    .documentProperty("department", "Engineering")  // custom property
    .column("Name", Product::name)
    .write(data);
```

Standard keys (`title`, `subject`, `author`/`creator`, `keywords`, `description`, `category`)
map to Excel core properties. Other keys become custom properties.

Also available on `ExcelWorkbook`.

## Named Ranges — Fluent API (v0.16.14+)

Register named ranges directly on the writer — no `afterData` callback needed:

```java
writer
    .column("Price", Product::price, c -> c.type(ExcelDataType.DOUBLE))
    .namedRange("PriceData", 0)  // column 0, covers all data rows
    .write(data);
// Other sheets can reference: =SUM(PriceData)
```

For manual control, use `SheetContext.namedRange()` in `afterData` callbacks.
