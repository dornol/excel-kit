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

## Merging Equal Values

`mergeSameValues()` merges runs of equal values of a column vertically, as reports do for a department over its rows.
The rows must be sorted so that equal values are neighbours.

```java
writer.column("Department", Emp::department, c -> c.mergeSameValues())
      .column("Team", Emp::team, c -> c.mergeSameValues())    // merges only inside a department
      .column("Name", Emp::name);
```

- A run ends where the value changes or is empty (empty values are not merged), and a run of one row is not merged.
- A column merges only within the merged regions of the merge columns before it: a team never spans two departments.
- A rolled-over sheet starts its runs again.
- The cells below the first one of a region keep their values: unmerging in Excel shows them, and a reader without
  `fillMergedCells()` still reads every row. With `fillMergedCells()` the first cell's value is read for all of them.
- Merged cells do not combine with `table(...)` (Excel repairs such a file): a writer with both fails when written.
  `TemplateListWriter` does not merge. Excel cannot sort or filter a range with merged cells of different sizes.

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

A summary covers the rows of its own sheet: when the data rolls over (`maxRows`, or a full sheet), every sheet ends with a
summary of its own rows, and there is no total over all the sheets. The same holds for `groupSubtotal(...)`.

## Empty Sheet Message

```java
ExcelWriter.<Order>create()
    .column("Name", Order::name)
    .column("Qty", Order::qty)
    .emptyMessage("No data")
    .write(orders);      // no orders: the header and a line that says "No data"
```

When no row is written, `emptyMessage(text)` puts the text below the header, centered over all the columns in gray
italics (one merged cell), so an empty export says so instead of looking like a broken file. A sheet with rows is not
touched, and every sheet of an `ExcelWorkbook` decides for itself.

- The line is not a data row: the auto filter covers the header only, a chart and a named range have nothing to cover,
  and a `summary(...)` below it counts nothing (its rows keep their labels, without formulas).
- Cannot be combined with `table(...)` (a merged cell has no place in an Excel table): the write fails with
  `ExcelWriteException`.
- A file with the line reads back with it as a data row: leave it off a sheet that people fill in and upload.

## Group Subtotals

```java
ExcelWriter.<Employee>create()
    .column("Department", Employee::department)
    .column("Name", Employee::name)
    .column("Salary", Employee::salary)
    .groupSubtotal(Employee::department, s -> s
        .label("Name", dept -> dept + " total")
        .keyColumn("Department")
        .sum("Salary").average("Salary"))
    .summary(s -> s.label("Total").sum("Salary"))
    .write(sortedByDepartment);
```

A subtotal row follows every run of rows with the same key (compared with `equals`: `1` and `1L`, or two arrays, are
different keys, so return a `String` or another value that compares as you mean). It is an Excel `SUBTOTAL` formula over the rows of its
group, so it follows the data when a person edits the file, and a `summary(...)` over the sheet counts every row once
(it switches to `SUBTOTAL` too). `sum`, `average`, `count`, `countA`, `min` and `max` are available; `style(...)` changes the
default bold-on-gray look; `label(column, key -> text)` sets the column of the text (the first column by default).

- `keyColumn("Dept")` writes the group's value into that column of the subtotal row (same type and format as the data
  cells), so an auto filter on `Dept` keeps the subtotal with its rows: filtering `Dept = A` shows A's rows and its
  subtotal. Without it the subtotal row is empty in that column and a filter on it hides the subtotal. The column must
  not be the label column or a subtotal column.
- `count` counts the cells that hold a number; for text columns use `countA` (counts every non-empty cell). Both are
  available in `summary(...)` too.
- Sort the rows by the key: a key that returns after another key gets a subtotal of its own.
- Subtotal rows are not data rows: they do not count for `rowNumberColumn`, `stripedRows` or `maxRows`. A rolled-over
  sheet closes its last group and starts a new one on the next sheet. Because `maxRows` does not count them, a sheet
  also turns over early when the subtotal rows (and the summary) would take the room Excel's 1,048,576 rows leave, so
  the default `maxRows` works with subtotals; a group that straddles that point gets a subtotal on each sheet.
- `ExcelPrintSetup.rowBreakEvery(n)` counts sheet rows, so the subtotal rows take part of each page.
- A subtotal row is always locked under `protectSheet(...)`, even in a column set to `locked(false)`, and its label counts
  for the column's automatic width (within `autoWidthSampleRows`). It takes the number format and alignment of its column;
  `cellColor`, `rowColor` and `rowStyle` do not apply to it (use `style(...)`).
- Conditional formatting rules (comparisons, `dataBar`, `iconSet`) skip the subtotal rows: a subtotal is a sum, so a threshold
  rule would colour it for its size and a data bar would scale every data row against it. The rule covers one range per run of
  data rows (`C2:C3 C5:C6`); `includeSubtotalRows()` on the rule puts the subtotal rows back under it. Column validations
  (`dropdown`, `validation`) still run over the subtotal rows, as they run over every row of the data block.
- Cannot be combined with `table(...)`, `chart(...)` or `namedRange(...)` (the write fails with `ExcelWriteException`).
- A subtotal and a `summary(...)` formula carry their result next to the formula (worked out while the rows are written
  from the cells as they are stored: text does not count for `sum`, a date is a number), so Protected View, previews and
  libraries that do not calculate show the number; Excel recalculates when it opens the file. Where the result cannot be
  known for sure (the column holds formulas of its own, an `average` over no numbers, `countA` over empty text) the
  formula is written without one and shows empty until Excel calculates. LibreOffice shows the stored result without
  recalculating (and calculates a formula that has none, skipping the subtotal cells like Excel).
- Reading the file back returns the subtotal rows as data rows.

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
- Page break every N data rows: `rowBreakEvery(n)` (each rollover sheet counts its own rows)

### Page breaks by data

A new printed page where the data changes, for example one page per department of a list sorted by department:

```java
writer.pageBreakOnChange(Employee::department)                        // key changed from the row before
      .printSetup(ps -> ps.repeatHeaderRows().fitToPageWidth());
writer.pageBreakWhen((previous, current) -> current.level() < previous.level())   // any condition
```

The breaks add to `rowBreakEvery(n)`; a rolled-over sheet breaks within itself only. Excel keeps at most 1026 page
breaks per sheet: the write fails with `ExcelWriteException` beyond that. `pageBreakWhen`/`pageBreakOnChange` are
on `ExcelWriter` and `ExcelSheetWriter`, not on `TemplateListWriter`.
