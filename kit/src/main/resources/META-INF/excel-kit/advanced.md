# excel-kit — Advanced Features

> Other topics: [Index](../AI.md) | [Quick Start](quick-start.md) | [Column Config](column-config.md) | [Reading](reading.md) | [CSV](csv.md)

## Multi-Sheet Workbook

```java
try (var wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.headerFontName("Arial").headerFontSize(12);

    wb.<User>sheet("Users")
        .column("Name", User::getName)
        .column("Age", User::getAge, cfg -> cfg.type(ExcelDataType.INTEGER))
        .autoFilter()
        .freezeRows(1)
        .write(userStream);

    wb.<Order>sheet("Orders")
        .column("ID", Order::getId)
        .tabColor(ExcelColor.GREEN)
        .write(orderStream);

    wb.finish().writeTo(out);
}
```

## Sheet Auto-Rollover

```java
// ExcelWriter: auto-split at 500K rows
ExcelWriter.<Product>create()
    .maxRows(500_000)
    .column("Name", Product::name)
    .write(millionRows);  // creates "Sheet", "Sheet (2)", etc.

// ExcelSheetWriter
wb.<Product>sheet("Data")
    .maxRows(100_000)
    .sheetName(i -> "Data-" + (i + 1))  // custom rollover naming
    .column("Name", Product::name)
    .write(stream);
```

## Group Headers (Merged)

```java
writer
    .column("Name", p -> p.name())
    .column("Price", p -> p.price()).group("Financial")
    .column("Qty", p -> p.qty()).group("Financial")
    .column("Notes", p -> p.notes())
    .write(data);
// Produces: | Name | Financial (merged) | Notes |
//           | Name | Price | Qty         | Notes |
```

### Multi-level groups

`group(String...)` accepts N levels, outermost first:

```java
.column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
.column("Q2", Row::q2, c -> c.group("Financial", "Revenue"))
.column("Profit", Row::profit, c -> c.group("Financial"))
.column("Name", Row::name)
// | Name |        Financial         |
// | Name |   Revenue    |  Profit   |
// | Name |  Q1  |  Q2   |  Profit   |
```

Columns with fewer levels merge vertically down into the column header cell.
Columns with no group span the full header depth.

### Group header comments (v0.16.11+)

Attach a comment to a merged group header cell, identified by its path (outermost first).
No-op if no column declares that path.

```java
writer
    .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
    .column("Q2", Row::q2, c -> c.group("Financial", "Revenue"))
    .groupComment("Quarterly revenue breakdown", "Financial", "Revenue")
    .groupComment(new ExcelCellComment("Auto-generated", "system"), "Financial");
```

## Row Number Column (v0.16.11+)

Convenience shortcut for a 1-based sequential row-number column, works across auto-rollover sheets:

```java
writer
    .rowNumberColumn("No.")           // equivalent to:
    //  .column("No.", (r, cur) -> cur.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))
    .column("Name", Product::name)
    .write(data);
```

## Formula Columns

```java
writer
    .column("Price", Product::price, c -> c.type(ExcelDataType.INTEGER))
    .column("Qty", Product::qty, c -> c.type(ExcelDataType.INTEGER))
    .column("Total", (row, cursor) ->
        "B" + (cursor.getRowOfSheet() + 1) + "*C" + (cursor.getRowOfSheet() + 1),
        c -> c.type(ExcelDataType.FORMULA))
    .write(data);
```

## Hyperlinks

```java
// Plain URL
.column("Website", Product::url, c -> c.type(ExcelDataType.HYPERLINK))

// Custom label
.column("Link", p -> new ExcelHyperlink(p.url(), "View"), c -> c.type(ExcelDataType.HYPERLINK))
```

## Rich Text

```java
.column("Desc", p -> new ExcelRichText()
    .text("Status: ").bold("APPROVED").text(" by ").styled("admin", s -> s.color(ExcelColor.BLUE).italic(true)))
    .type(ExcelDataType.RICH_TEXT)
```

## Image Embedding

```java
.column("Photo", p -> ExcelImage.png(imageBytes)).type(ExcelDataType.IMAGE)
// Also: ExcelImage.jpeg(bytes)
```

## Conditional Formatting

```java
writer
    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .conditionalFormatting(cf -> cf
        .columns(1)
        .greaterThan("10000", ExcelColor.LIGHT_RED)
        .lessThan("1000", ExcelColor.LIGHT_GREEN)
        .between("5000", "10000", ExcelColor.LIGHT_YELLOW))
    .write(data);
```

Operators: `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equalTo`, `notEqualTo`, `between`, `notBetween`

Also: `.dataBar()` for gradient bars, `.iconSet()` for arrows/traffic lights.

## Data Validation

```java
// Dropdown
.column("Status", p -> p.status(), c -> c.dropdown("Active", "Inactive", "Pending"))

// Advanced validations
.validation(ExcelValidation.integerBetween(0, 150))
.validation(ExcelValidation.decimalBetween(0.0, 4.0))
.validation(ExcelValidation.textLength(1, 100))
.validation(ExcelValidation.dateRange(LocalDate.of(2024, 1, 1), LocalDate.of(2024, 12, 31)))
.validation(ExcelValidation.formula("AND(A2>0,A2<100)"))
.validation(ExcelValidation.listFromRange("Options!$A$1:$A$5"))

// Error messages
ExcelValidation.integerBetween(1, 100).errorTitle("Invalid").errorMessage("Enter 1-100")
```

## Freeze Panes

```java
writer.freezeRows(1);          // freeze 1 row below the header
writer.freezeCols(2);          // freeze first 2 columns from the left
writer.freezePane(2, 1);       // freeze both axes (2 cols + 1 row)
```

Same three methods are available on `ExcelSheetWriter` (multi-sheet workbooks).
Negative values throw `IllegalArgumentException`.

## Sheet Protection

```java
writer
    .column("Name", p -> p.name(), c -> c.locked(false))   // editable
    .column("Price", p -> p.price(), c -> c.locked(true))   // read-only
    .protectSheet("password123")
    .write(data);
```

## Workbook Protection

```java
writer.protectWorkbook("password123");  // prevent add/delete/rename sheets
```

## Password Encryption

```java
// Option 1 — bind at writer level
ExcelWriter.create().password("secret").column(...).write(data).writeTo(out);

// Option 2 — late-binding (service builds handler, presentation layer applies password)
ExcelHandler h = ExcelWriter.create().column(...).write(data);
h.writeTo(out, "secret");                 // OutputStream
h.writeTo(path, "secret");                // Path
h.writeTo(out, pwChars);                  // char[] (zeroed after use)
h.writeTo(path, pwChars);                 // Path + char[]
```

## Charts

```java
writer
    .column("Name", Product::name)
    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.BAR)  // BAR, LINE, PIE, SCATTER, AREA, DOUGHNUT
        .title("Sales Report")
        .categoryColumn(0)
        .valueColumn(1, "Sales")
        .position(3, 0, 12, 20))
    .write(data);
```

## Summary/Footer Rows

```java
writer
    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .column("Qty", p -> p.qty(), c -> c.type(ExcelDataType.INTEGER))
    .summary(s -> s.label("Total").sum("Price").sum("Qty").average("Price"))
    .write(data);
```

Operations: `sum`, `average`, `count`, `min`, `max`

## Named Ranges

```java
writer.afterData(ctx -> {
    ctx.namedRange("Categories", 0, 1, ctx.getCurrentRow() - 1);
    return ctx.getCurrentRow();
})
```

## Print Setup

```java
writer.printSetup(p -> p
    .landscape()
    .paperSize(ExcelPrintSetup.PaperSize.A4)
    .fitToPage(1, 0)
    .repeatRows(0, 0))
```

## Callbacks

```java
writer
    .beforeHeader(ctx -> { /* write rows before header */ return ctx.getCurrentRow(); })
    .afterData(ctx -> { /* write rows after data */ return ctx.getCurrentRow(); })
    .afterAll(ctx -> { /* after everything */ return ctx.getCurrentRow(); })
    .onProgress(10_000, (count, cursor) -> log.info("Processed {}", count))
```

## Sheet Tab Color

```java
writer.tabColor(ExcelColor.STEEL_BLUE);
// or: writer.tabColor(255, 0, 0);
```

## Row Grouping

```java
writer.afterData(ctx -> {
    ctx.groupRows(1, 5);              // collapsible group
    ctx.groupRows(7, 10, true);       // collapsed by default
    return ctx.getCurrentRow();
})
```

## Template Writing

```java
try (var tw = new ExcelTemplateWriter(templateInputStream)) {
    tw.cell("B2", "Report Title");
    tw.cell("B3", LocalDate.now());
    tw.<Product>list("Name", Product::name)
        .list("Price", Product::price)
        .write(productStream, outputStream);
}
```

## Unified Schema

```java
var schema = ExcelKitSchema.<Product>builder()
    .column("Name", Product::getName, (p, cell) -> p.setName(cell.asString()))
    .column("Price", Product::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
            c -> c.type(ExcelDataType.INTEGER))
    .build();

// Write
schema.excelWriter().write(data).writeTo(out);

// Read
schema.excelReader(Product::new, null).read(in, consumer);
```
