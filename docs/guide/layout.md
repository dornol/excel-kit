# Layout & Structure

> [Back to Index](index.md)

## Auto-Filter and Freeze Panes

```java
writer
    .autoFilter(true)    // dropdown filter over the header and the data rows
    .freezeRows(1)       // freeze 1 row below the header
    .column("Name", p -> p.name())
    .write(data);

// Columns-only
writer.freezeCols(2);        // freeze 2 columns from the left

// Both axes
writer.freezePane(2, 1);     // 2 columns + 1 row
```

| Method | Vertical split (cols) | Horizontal split (rows) |
|--------|----------------------|------------------------|
| `freezeRows(n)` | 0 | n |
| `freezeCols(n)` | n | 0 |
| `freezePane(c, r)` | c | r |

All three available on `ExcelSheetWriter` too.

The filter written to the file covers the header row (the row of the column names, below any group header) and every data
row of the sheet. It stops above the summary rows and anything written by `afterData`, and with `maxRows()` each
rolled-over sheet filters its own rows. Excel extends a filter over the rows that follow it directly and hold data, so
summary rows written right below the data are filtered along with it once the filter is applied (checked on Excel). An
empty row above them keeps them out: `afterData(ctx -> ctx.getCurrentRow() + 1)` leaves one before the summary (checked on
Excel: the Sum and Average rows stay visible while a filter is applied). A sheet with a `table(...)` gets
its filter buttons from the table itself, since a second filter over the same cells would overlap it: `autoFilter(true)`
turns them on, and without it the table has none.

## Column Outline (Grouping)

Group columns for collapse/expand:

```java
writer
    .column("Name", p -> p.name())
    .column("Detail1", p -> p.detail1(), c -> c.outline(1))
    .column("Detail2", p -> p.detail2(), c -> c.outline(1))
    .column("Summary", p -> p.summary())
    .write(data);
```

Adjacent columns with the same outline level are grouped. Supports levels 1-7.

## Column Hiding

```java
.column("Internal Code", p -> p.code(), cfg -> cfg.hidden())
```

Hidden in Excel but data is still written.

## Summary / Footer Rows

```java
writer
    .column("Name", Product::name)
    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .column("Qty", p -> p.qty(), c -> c.type(ExcelDataType.INTEGER))
    .summary(s -> s
        .label("Total")
        .sum("Price").sum("Qty"))
    .write(data);
```

Multiple operations create multiple rows:
```java
.summary(s -> s
    .sum("Amount")       // row 1: SUM
    .average("Amount")   // row 2: AVERAGE
    .count("Amount")     // row 3: COUNT
    .min("Score")        // row 4: MIN
    .max("Score"))       // row 5: MAX
```

Custom label column: `.label("Total", "Grand Total")`

With several operations each row needs its own name, so a label is followed by the operation: `.label("Total")` with
`sum` and `average` gives the rows `Total - Sum` and `Total - Average`. With one operation the row carries the label as
it is, and without a label a row is named after its operation (`Sum`, `Average`).

Works with sheet rollover — formulas generated per-sheet.

A summary or label column name that is not among the sheet's columns fails with
`ExcelWriteException` instead of being skipped. When a sheet has no data rows, the label row is
written without formulas.

## Named Ranges

```java
writer.afterData(ctx -> {
    ctx.namedRange("Categories", 0, 1, ctx.getCurrentRow() - 1);
    // or: ctx.namedRange("Categories", "Sheet1!$A$2:$A$100");
    return ctx.getCurrentRow();
});
```

## Row Grouping / Outlining

Group rows for collapse/expand in callbacks:

```java
writer.afterData(ctx -> {
    ctx.groupRows(1, 5);              // expanded
    ctx.groupRows(7, 10, true);       // collapsed
    return ctx.getCurrentRow();
});
```

## Print Setup

```java
writer.printSetup(ps -> ps
    .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
    .paperSize(ExcelPrintSetup.PaperSize.A4)
    .margins(0.5, 0.5, 0.75, 0.75)   // left, right, top, bottom (inches)
    .headerCenter("Invoice Report")
    .footerCenter("Page &P of &N")
    .repeatHeaderRows()
    .fitToPageWidth());
```

**Options:**
- Orientation: `PORTRAIT`, `LANDSCAPE`
- Paper: `LETTER`, `LEGAL`, `A3`, `A4`, `A5`, `B4`, `B5`
- Margins: individual `leftMargin()`, `rightMargin()`, `topMargin()`, `bottomMargin()`
- Headers/Footers: `headerLeft/Center/Right()`, `footerLeft/Center/Right()`
- Special codes: `&P` (page), `&N` (total pages), `&D` (date), `&T` (time), `&F` (filename)
- Repeat: `repeatHeaderRows()` or `repeatRows(first, last)`
- Fit: `fitToPageWidth()` or `fitToPage(width, height)`
