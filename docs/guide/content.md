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

FontStyle options: `bold()`, `italic()`, `underline()`, `strikethrough()`, `color()`, `fontSize()`

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
