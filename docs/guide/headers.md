# Headers & Group Headers

> [Back to Index](index.md)

## Header Color (Global)

```java
ExcelWriter.create().headerColor(ExcelColor.STEEL_BLUE);
```

## Header Font

```java
writer.headerFontName("Arial").headerFontSize(14);
```

## Header Style (v0.16.14+)

Configure header cell alignment, bold, border, and wrap text:

```java
writer
    .headerColor(ExcelColor.STEEL_BLUE)
    .headerStyle(h -> h
        .bold(true)
        .alignment(HorizontalAlignment.LEFT)
        .border(ExcelBorderStyle.MEDIUM)
        .wrapText(true))
    .headerFontName("Arial").headerFontSize(14)
```

Font name, font size, and background color are set separately via their dedicated methods.
Available on both `ExcelWriter` and `ExcelWorkbook`.

## Header Row Height (v0.16.11+)

```java
writer.headerRowHeight(24f);  // points; 0 = default. Applies to all header rows including groups.
```

## Per-Column Header Font Color

Conditionally highlight specific column headers:

```java
boolean hasError = checkSomething();

writer.column("Amount", Product::amount, cfg -> cfg
    .headerFontColor(hasError ? ExcelColor.RED : null))
```

Accepts `ExcelColor`, `null` (default), or `headerFontColor(int r, int g, int b)`.

## Per-Column Header Background (v0.16.11+)

```java
writer
    .headerColor(ExcelColor.STEEL_BLUE)                    // default for all
    .column("Amount", Product::amount, c -> c
        .headerBackgroundColor(ExcelColor.LIGHT_RED))      // this header only
```

`.headerBackgroundColor(null)` restores default.

## Header Comments

Attach a static comment to a column's header cell:

```java
writer
    .column("Birth Date", User::birth, cfg -> cfg
        .type(ExcelDataType.DATE)
        .headerComment("Enter in YYYY-MM-DD format"))
    .column("Amount", User::amount, cfg -> cfg
        .headerComment(new ExcelCellComment("In KRW", "admin").size(3, 4)))
```

When combined with group headers, the comment attaches to the column header row (not the group row).

## Group Headers

Create multi-row merged headers using `.group()`:

```java
writer
    .column("Name", p -> p.name())
    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER).group("Financial"))
    .column("Qty", p -> p.qty(), c -> c.type(ExcelDataType.INTEGER).group("Financial"))
    .column("Notes", p -> p.notes())
    .write(data);
```

Produces:

| Name | Financial (merged) | Notes |
|------|---------------------|-------|
| Name | Price \| Qty | Notes |

Adjacent columns with the same group name are merged horizontally. Ungrouped columns are merged vertically across all header rows.

### Multi-Level Groups (v0.16.9+)

Pass multiple level names to `group(String...)`, outermost first:

```java
writer
    .column("Name", Row::name)
    .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
    .column("Q2", Row::q2, c -> c.group("Financial", "Revenue"))
    .column("Profit", Row::profit, c -> c.group("Financial"))
    .write(data);
```

Produces 3 header rows (2 group + 1 column):

| Name | Financial |
|------|-----------|
| Name | Revenue \| Profit |
| Name | Q1 \| Q2 \| Profit |

- Adjacent columns with equal values on the same row merge horizontally.
- Columns with fewer levels merge vertically into the column header cell.
- Columns with no group span the full header depth.

### Group Header Comments (v0.16.11+)

Attach a comment to a merged group header cell by path (outermost first):

```java
writer
    .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
    .column("Q2", Row::q2, c -> c.group("Financial", "Revenue"))
    .groupComment("Quarterly revenue", "Financial", "Revenue")
    .groupComment(new ExcelCellComment("Owner: Finance team", "system"), "Financial")
    .write(data);
```

No-op if no column declares that path.

## Row Number Column (v0.16.11+)

```java
writer.rowNumberColumn("No.")  // 1-based sequential, correct across auto-rollover
```

## Reading Multi-Row Headers (v0.16.13+)

See [Reading — Multi-row headers](reading.md#advanced-options).
