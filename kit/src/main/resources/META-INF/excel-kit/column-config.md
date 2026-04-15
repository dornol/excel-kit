# excel-kit — Column Configuration

> Other topics: [Index](../AI.md) | [Quick Start](quick-start.md) | [Reading](reading.md) | [Advanced](advanced.md) | [CSV](csv.md)

## Data Types

Set via `.type(ExcelDataType.XXX)`:

| Type | Java Type | Default Format |
|------|-----------|----------------|
| `STRING` (default) | String | — |
| `INTEGER` | Number | `#,##0` |
| `LONG` | Number | `#,##0` |
| `DOUBLE` | Number | `#,##0.00` |
| `DOUBLE_PERCENT` | Number | `0.00%` |
| `DATE` | LocalDate/Date | `yyyy-mm-dd` |
| `DATETIME` | LocalDateTime | `yyyy-mm-dd hh:mm:ss` |
| `FORMULA` | String (formula) | — |
| `HYPERLINK` | String or ExcelHyperlink | — |
| `RICH_TEXT` | ExcelRichText | — |
| `IMAGE` | ExcelImage | — |
| `BOOLEAN` | Boolean | — |

Custom format: `.format("#,##0.00")` or use `ExcelDataFormat` presets.

## Column Styling Methods

All methods available via lambda configurer on `ExcelWriter`, `ExcelSheetWriter`, and `CsvWriter`:

### Layout
| Method | Description |
|--------|-------------|
| `.width(int)` | Fixed width (disables auto-resize) |
| `.minWidth(int)` | Minimum auto-resize width |
| `.maxWidth(int)` | Maximum auto-resize width |
| `.alignment(HorizontalAlignment)` | LEFT, CENTER (default), RIGHT |
| `.verticalAlignment(VerticalAlignment)` | TOP, CENTER (default), BOTTOM, JUSTIFY |
| `.wrapText()` / `.wrapText(false)` | Text wrapping (default: enabled) |
| `.rotation(int)` | Text angle -90 to 90 degrees |
| `.indentation(int)` | Indent level 0-250 |

### Font
| Method | Description |
|--------|-------------|
| `.bold(boolean)` | Bold text |
| `.fontSize(int)` | Font size in points |
| `.fontName(String)` | Font family ("Arial", "맑은 고딕") |
| `.fontColor(ExcelColor)` | Font color (preset or RGB) |
| `.fontColor(int r, int g, int b)` | Font color (RGB) |
| `.strikethrough()` | Strikethrough text |
| `.underline()` | Underline text |

### Background & Borders
| Method | Description |
|--------|-------------|
| `.backgroundColor(ExcelColor)` | Cell background color |
| `.border(ExcelBorderStyle)` | Uniform border (THIN, MEDIUM, THICK, DASHED, DOTTED, DOUBLE, NONE, ...) |
| `.borderTop/Bottom/Left/Right(ExcelBorderStyle)` | Per-side border override |

### Header
| Method | Description |
|--------|-------------|
| `.headerFontColor(ExcelColor)` | Override header font color for this column |
| `.headerFontColor(int r, int g, int b)` | Override header font color (RGB) |
| `.headerFontColor(null)` | Use default header style |
| `.headerBackgroundColor(ExcelColor)` | Override header background for this column only (v0.16.11+) |
| `.headerBackgroundColor(int r, int g, int b)` | Override header background (RGB) |
| `.headerBackgroundColor(null)` | Fall back to writer-level `headerColor` |
| `.headerComment(String)` | Static comment on the column's header cell |
| `.group(String... levels)` | Group header (merged rows above). N levels = N group rows (v0.16.9+) |

### Data & Behavior
| Method | Description |
|--------|-------------|
| `.type(ExcelDataType)` | Cell data type |
| `.format(String)` | Excel number/date format |
| `.dropdown(String...)` | Dropdown validation options |
| `.validation(ExcelValidation)` | Advanced data validation |
| `.cellColor(CellColorFunction)` | Per-cell conditional background |
| `.comment(Function<T, String>)` | Per-cell comment/note |
| `.outline(int)` | Column outline level 1-7 |
| `.hidden()` | Hide column |
| `.locked(boolean)` | Lock/unlock for sheet protection |

## Usage Examples

### ExcelWriter
```java
ExcelWriter.<Product>create().headerColor(ExcelColor.STEEL_BLUE)
    .column("Name", Product::name)
    .column("Price", Product::price, cfg -> cfg
        .type(ExcelDataType.INTEGER)
        .format("#,##0")
        .alignment(HorizontalAlignment.RIGHT)
        .backgroundColor(ExcelColor.LIGHT_YELLOW))
    .column("Status", Product::status, cfg -> cfg
        .dropdown("Active", "Inactive")
        .fontColor(ExcelColor.RED))
    .write(data);
```

### ExcelSheetWriter (multi-sheet)
```java
workbook.<Product>sheet("Products")
    .column("Name", Product::name)
    .column("Price", Product::price, cfg -> cfg
        .type(ExcelDataType.INTEGER)
        .format("#,##0")
        .alignment(HorizontalAlignment.RIGHT))
    .column("Status", Product::status, cfg -> cfg
        .dropdown("Active", "Inactive")
        .fontColor(ExcelColor.RED))
    .write(data);
```

### Conditional Header Font Color
```java
boolean hasError = checkSomething();

writer.column("Amount", Product::amount, cfg -> cfg
    .headerFontColor(hasError ? ExcelColor.RED : null))
```

### Default Column Style
```java
ExcelWriter.<Product>create()
    .defaultStyle(d -> d.fontName("Arial").fontSize(10).alignment(HorizontalAlignment.LEFT))
    .column("Name", Product::name)           // inherits defaults
    .column("Price", Product::price, cfg -> cfg
        .alignment(HorizontalAlignment.RIGHT)) // overrides default
    .write(data);
```

### Row-Level & Cell-Level Styling
```java
writer
    .rowColor(p -> p.isError() ? ExcelColor.LIGHT_RED : null)
    .column("Amount", Product::amount, cfg -> cfg
        .cellColor((value, row) -> {
            double amt = ((Number) value).doubleValue();
            if (amt < 0) return ExcelColor.LIGHT_RED;
            return null;
        }))
    .write(data);
```

Priority: `cellColor` > `rowColor` > `backgroundColor`

## Header Color (Global)

```java
// Background color for all headers
ExcelWriter.create().headerColor(ExcelColor.STEEL_BLUE);

// Font name and size for all headers
writer.headerFontName("Arial").headerFontSize(14);

// Explicit height for every header row (including group rows, v0.16.11+)
writer.headerRowHeight(24f);  // points; 0 = default
```

### Per-column background override (v0.16.11+)

```java
writer
    .headerColor(ExcelColor.STEEL_BLUE)                             // default for all headers
    .column("Name", Product::name)
    .column("Amount", Product::amount, c -> c.type(ExcelDataType.INTEGER)
        .headerBackgroundColor(ExcelColor.LIGHT_RED))               // this header only: red
    .write(data);
```

Useful for highlighting required / alert columns against the baseline header color.

Presets: `WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`, etc.

## Conditional Columns

```java
writer
    .column("Name", p -> p.name())
    .columnIf("Age", showAge, p -> p.age())  // only when showAge == true
    .column("Email", p -> p.email())
    .write(data);
```
