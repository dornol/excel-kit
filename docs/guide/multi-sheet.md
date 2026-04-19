# Multi-Sheet Workbook

> [Back to Index](index.md)

## ExcelWorkbook

Write different data types to separate sheets:

```java
try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.headerFontName("Arial").headerFontSize(12);

    wb.<User>sheet("Users")
        .column("Name", User::getName)
        .column("Age", User::getAge, c -> c.type(ExcelDataType.INTEGER))
        .autoFilter().freezeRows(1)
        .write(userStream);

    wb.<Order>sheet("Orders")
        .column("ID", Order::getId)
        .column("Amount", Order::getAmount, c -> c.type(ExcelDataType.DOUBLE))
        .tabColor(ExcelColor.GREEN)
        .write(orderStream);

    wb.finish().writeTo(outputStream);
}
```

Each `ExcelSheetWriter` supports the same features as `ExcelWriter`: column config, callbacks, auto-filter, freeze panes, row/cell color, protection, charts, print setup, summary, etc.

## Sheet Auto-Rollover

Auto-split when row limit is reached:

```java
// ExcelWriter
ExcelWriter.<Product>create()
    .maxRows(500_000)
    .column("Name", Product::name)
    .write(millionRows);  // creates "Sheet", "Sheet (2)", etc.

// ExcelSheetWriter with custom naming
wb.<Product>sheet("Data")
    .maxRows(100_000)
    .sheetName(i -> "Data-" + (i + 1))
    .column("Name", Product::name)
    .write(stream);
```

If `sheetName()` is not set, rollover sheets are named `"baseName (2)"`, `"baseName (3)"`, etc.

## Sheet Tab Color

```java
// ExcelWriter — applies to all sheets (including rollover)
writer.tabColor(ExcelColor.STEEL_BLUE);
writer.tabColor(255, 0, 0);  // RGB

// ExcelSheetWriter — per sheet
wb.<User>sheet("Users").tabColor(ExcelColor.BLUE);
wb.<Order>sheet("Orders").tabColor(ExcelColor.GREEN);
```
