# Advanced Content

> [Back to Index](index.md)

## Formula Columns

```java
writer
    .column("Price", Product::price, c -> c.type(ExcelDataType.INTEGER))
    .column("Qty", Product::qty, c -> c.type(ExcelDataType.INTEGER))
    .column("Subtotal", (row, cursor) ->
            "B" + (cursor.getRowOfSheet() + 1) + "*C" + (cursor.getRowOfSheet() + 1),
        c -> c.type(ExcelDataType.FORMULA).format(ExcelDataFormat.CURRENCY_KRW.getFormat()))
    .write(data);
```

Use `SheetContext.columnLetter()` in callbacks:
```java
writer.afterData(ctx -> {
    String col = SheetContext.columnLetter(0); // "A"
    var row = ctx.getSheet().createRow(ctx.getCurrentRow());
    row.createCell(0).setCellFormula("SUM(%s2:%s%d)".formatted(col, col, ctx.getCurrentRow()));
    return ctx.getCurrentRow() + 1;
});
```

> Do not pass untrusted user input as formula values — use `STRING` type for user-supplied data.

## Hyperlink Columns

```java
// Plain URL
.column("Website", Product::url, c -> c.type(ExcelDataType.HYPERLINK))

// Custom label
.column("Link", p -> new ExcelHyperlink(p.url(), "View Details"), c -> c.type(ExcelDataType.HYPERLINK))
```

## Rich Text

Mixed formatting within a single cell:

```java
.column("Desc", p -> new ExcelRichText()
        .text("Status: ")
        .bold("APPROVED")
        .text(" — reviewed by ")
        .styled("admin", s -> s.color(ExcelColor.BLUE).italic(true)),
    c -> c.type(ExcelDataType.RICH_TEXT))
```

FontStyle options: `bold()`, `italic()`, `underline()`, `strikethrough()`, `color()`, `fontSize()` (1–409)

**Rich text needs shared strings.** Streaming writes a string cell as its plain text, so the formatting is kept only
when the workbook is created with shared strings; without them the write fails (or, with
`writeErrorPolicy(LENIENT)`, writes the plain text):

```java
ExcelWriter.<Item>create(o -> o.useSharedStrings(true))
    .column("Desc", p -> new ExcelRichText().text("Status: ").bold("APPROVED"),
        c -> c.type(ExcelDataType.RICH_TEXT))
```

The shared strings table is kept in memory, so with millions of distinct strings this costs heap. Unformatted
text and every formatted segment start from the column's font (name, size, colour, bold), and a segment's own
options are laid over it; a cell with no formatted segment at all is a plain string in the column's font.
`ExcelTemplateWriter.cell(...)` takes an `ExcelRichText` too and keeps its formatting.

## Image Embedding

```java
byte[] imageBytes = Files.readAllBytes(Path.of("logo.png"));

.column("Photo", p -> ExcelImage.png(imageBytes), c -> c.type(ExcelDataType.IMAGE))
```

**Custom size** (columns × rows):

```java
.column("Photo", p -> ExcelImage.png(imageBytes).size(3, 4), c -> c.type(ExcelDataType.IMAGE))
```

**From URL** (auto-detects PNG/JPEG):

```java
.column("Photo", p -> ExcelImage.fromUrl(p.getPhotoUrl()), c -> c.type(ExcelDataType.IMAGE))
```

`fromUrl` accepts only HTTP(S), uses 10-second timeouts, and limits downloads
to 10 MiB by default. Use `ExcelImage.fromUrl(url, maxBytes)` for a stricter
limit. Validate hosts yourself when URLs come from untrusted users.

Factory methods: `ExcelImage.png(byte[])`, `ExcelImage.jpeg(byte[])`, `ExcelImage.fromUrl(String)`,
`ExcelImage.fromUrl(String, int)`

## Cell Comments (Notes)

Conditional per-cell comments:

```java
.column("Score", p -> p.score(), cfg -> cfg
    .type(ExcelDataType.INTEGER)
    .comment(p -> p.score() < 50 ? "Low score - needs review" : null))
```

Returns `null` to skip. Comments appear as yellow note icons.

## Chart Generation

```java
writer
    .column("Name", Product::name)
    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
    .column("Profit", p -> p.profit(), c -> c.type(ExcelDataType.INTEGER))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.BAR)
        .title("Sales vs Profit")
        .categoryColumn(0)
        .valueColumn(1, "Sales")
        .valueColumn(2, "Profit")
        .categoryAxisTitle("Product")
        .valueAxisTitle("Amount")
        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM)
        .showDataLabels(true)
        .position(3, 0, 12, 20))
    .write(data);
```

**Chart types:** `BAR`, `LINE`, `PIE`, `SCATTER`, `AREA`, `DOUGHNUT`

With `maxRows()` rollover each sheet gets its own chart over that sheet's rows. `position(...)` refuses a negative or
backwards column pair and two given rows that run upwards; a category or value column that is not among the columns fails
with `ExcelWriteException`.

**Combination chart:** a line over the main series, on the same value axis (`lineColumn`) or on a second axis on the
right (`secondaryLineColumn`), for a quantity with another unit (columns of revenue, a line of the growth rate):

```java
writer
    .column("Month", Sale::month)
    .column("Revenue", Sale::revenue, c -> c.type(ExcelDataType.INTEGER))
    .column("Target", Sale::target, c -> c.type(ExcelDataType.INTEGER))
    .column("Growth", Sale::growth, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.BAR)
        .categoryColumn(0)
        .valueColumn(1, "Revenue")              // the columns
        .lineColumn(2, "Target")                // a line on the same axis
        .secondaryLineColumn(3, "Growth")       // a line on the right axis
        .valueAxisTitle("Revenue")
        .secondaryValueAxisTitle("Growth"))
    .write(data);
```

The lines go over a vertical `BAR`, a `LINE` or an `AREA` chart that has at least one `valueColumn`; a pie, doughnut,
scatter or horizontal-bar chart refuses them with `ExcelWriteException`. Lines are drawn straight (not smoothed), and
columns sit between the tick marks, so the first and the last one are not cut in half.

**Bar options:** `barDirection(VERTICAL | HORIZONTAL)`, `barGrouping(STANDARD | STACKED | PERCENT_STACKED)`

**Legend positions:** `BOTTOM`, `LEFT`, `RIGHT`, `TOP`, `TOP_RIGHT`

### Scatter Chart

Both axes are numeric:
```java
.chart(chart -> chart
    .type(ExcelChartConfig.ChartType.SCATTER)
    .categoryColumn(0).valueColumn(1, "Y values")
    .categoryAxisTitle("X").valueAxisTitle("Y"))
```

### Doughnut Chart

```java
.chart(chart -> chart
    .type(ExcelChartConfig.ChartType.DOUGHNUT)
    .categoryColumn(0).valueColumn(1, "Share")
    .showDataLabels(true))
```
