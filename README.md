# excel-kit

A lightweight Java library for streaming Excel (.xlsx) and CSV generation/parsing built on Apache POI.
Designed for large datasets with minimal memory footprint, fluent DSL-style column definitions,
password-encrypted Excel export, and optional Bean Validation support.

- **Group/Artifact:** `io.github.dornol:excel-kit`
- **License:** MIT

## Features

**Excel Writing** (SXSSFWorkbook streaming)
- Fluent column DSL with type, format, alignment, and style options
- Automatic sheet splitting when row limit is reached
- Customizable header color, row height, auto-filter, and freeze panes
- Lifecycle callbacks with `SheetContext`: `beforeHeader`, `afterData`, `afterAll`
- Dropdown data validation (select list) per column
- Row-level conditional styling (background color)
- Cell-level conditional styling via `CellColorFunction` (per-cell background based on value)
- Formula columns (`ExcelDataType.FORMULA`) for computed values
- Hyperlink columns (`ExcelDataType.HYPERLINK`) for clickable URLs
- Group headers — merged multi-row headers for column grouping
- Column outline — expand/collapse column groups via `outline()`
- Column hiding via `hidden()` — hide columns in the output while still writing data
- Rich text via `ExcelDataType.RICH_TEXT` — mixed formatting (bold, italic, colors) within a single cell
- Print setup via `printSetup()` — page orientation, paper size, margins, headers/footers, repeat rows, fit-to-page
- Explicit multi-sheet workbook with different data types per sheet (`ExcelWorkbook`)
- Auto-rollover for `ExcelSheetWriter` via `maxRows()`
- Progress callback via `onProgress()` for large dataset monitoring
- Password-encrypted Excel output
- Consume-once output via `ExcelHandler`
- Cell border styles via `border()` — THIN, MEDIUM, THICK, DASHED, DOTTED, DOUBLE, etc.
- Per-side border control via `borderTop()`, `borderBottom()`, `borderLeft()`, `borderRight()`
- Cell comments (notes) via `comment()` — conditional per-cell comments
- Conditional formatting rules via `conditionalFormatting()` — greaterThan, lessThan, between, etc.
- Sheet protection via `protectSheet()` with per-column `locked()` control
- Image embedding via `ExcelDataType.IMAGE` with `ExcelImage.png()` / `ExcelImage.jpeg()`
- Chart generation via `chart()` — BAR, LINE, PIE, SCATTER, AREA, DOUGHNUT charts with XDDF API
- Map-based writing via `ExcelMapWriter` — write `Map<String, Object>` without typed POJOs
- Text rotation via `rotation()` — rotate cell text from -90 to 90 degrees
- Font color via `fontColor()` — RGB or preset color for cell text
- Strikethrough via `strikethrough()` and underline via `underline()` for font styling
- Advanced data validation via `validation()` — integer/decimal ranges, text length, date ranges, custom formulas
- Row grouping via `SheetContext.groupRows()` — collapsible row groups in callbacks
- Sheet tab color via `tabColor()` — colorize sheet tabs
- Vertical alignment via `verticalAlignment()` — TOP, CENTER, BOTTOM, JUSTIFY
- Text wrapping via `wrapText()` — configurable per-column text wrapping (default: enabled)
- Font name via `fontName()` — custom font family (e.g., "Arial", "맑은 고딕")
- Cell indentation via `indentation()` — indent cell content by level (0–250)
- Workbook protection via `protectWorkbook()` — prevent sheet add/delete/rename/reorder
- Header font customization via `headerFontName()` and `headerFontSize()`
- Default column style via `defaultStyle()` — writer-level style defaults inherited by all columns
- Summary/footer rows via `summary()` — fluent DSL for SUM, AVERAGE, COUNT, MIN, MAX formulas
- Named ranges via `SheetContext.namedRange()` — create workbook-scoped named ranges in callbacks
- List validation from cell range via `ExcelValidation.listFromRange()` — dropdown options from sheet/range reference
- Custom cell conversion via `CellData.as(Function)` — ad-hoc type conversion (e.g., `UUID::fromString`)
- Default value overloads — `asInt(defaultValue)`, `asLong(defaultValue)`, `asDouble(defaultValue)`, `asString(defaultValue)`, `as(Function, defaultValue)`

**Excel Reading** (SAX-based streaming)
- Header name-based column mapping — columns matched by header name, order-independent
- Index-based column mapping via `columnAt()` — read specific columns by index
- Positional column mapping with skip support
- **Mapping mode** via `ExcelReader.mapping()` — immutable object / Java record support with `RowData`
- Configurable header row index and sheet index
- Optional Bean Validation integration with per-row results
- Stream-based reading via `readAsStream()`
- Read progress callback via `onProgress()`
- Large file support configuration
- Multi-sheet discovery via `getSheetNames()` and `getSheetHeaders()`
- Map-based reading via `ExcelMapReader` — read into `Map<String, String>` without typed POJOs

**CSV Writing**
- Streaming write to temp file, then flush to `OutputStream`
- Proper escaping (quotes, commas, newlines)
- UTF-8 BOM for Excel compatibility
- Configurable delimiter and charset
- Progress callback via `onProgress()`
- Configurable CSV injection defense via `csvInjectionDefense()` — toggle formula character prefixing

**CSV Reading** (OpenCSV-based)
- Header name-based column mapping — columns matched by header name, order-independent
- Index-based column mapping via `columnAt()`
- **Mapping mode** via `CsvReader.mapping()` — immutable object / Java record support with `RowData`
- Header auto-detection with BOM removal
- Column mapping DSL with Bean Validation support
- Configurable delimiter, charset, and header row index
- Map-based writing via `CsvMapWriter`

**Unified Schema**
- `ExcelKitSchema` — define columns once for both reading and writing
- Write configuration (type, format, style) embedded in schema
- Schema-based readers automatically use name-based column matching

## Installation

**Gradle (Kotlin DSL)**
```kotlin
dependencies {
    implementation("io.github.dornol:excel-kit:0.8.2")
}
```

**Maven**
```xml
<dependency>
  <groupId>io.github.dornol</groupId>
  <artifactId>excel-kit</artifactId>
  <version>0.8.0</version>
</dependency>
```

### Runtime Dependencies

The library declares the following as `compileOnly`. You must provide them at runtime:

| Dependency | Required For |
|------------|-------------|
| `org.apache.poi:poi-ooxml` | Excel read/write |
| `org.slf4j:slf4j-api` | Logging |
| `jakarta.validation:jakarta.validation-api` | Bean Validation (optional) |
| `com.opencsv:opencsv` | CSV reading |

You also need runtime implementations:
- **SLF4J:** e.g. `org.slf4j:slf4j-simple` or Logback
- **Bean Validation:** e.g. `org.hibernate:hibernate-validator` (if using validation)

## Quick Start

### Excel Writing

```java
record Person(long id, String name, int age) {}

var data = Stream.of(new Person(1, "Alice", 30), new Person(2, "Bob", 28));

ExcelHandler handler = new ExcelWriter<Person>()
        .column("ID", p -> p.id())
            .type(ExcelDataType.LONG)
            .alignment(HorizontalAlignment.RIGHT)
        .column("Name", p -> p.name())
        .column("Age", p -> p.age())
            .type(ExcelDataType.INTEGER)
        .write(data);

try (var os = Files.newOutputStream(Path.of("people.xlsx"))) {
    handler.consumeOutputStream(os);
}
```

### Excel Reading

```java
class User {
    String name;
    Integer age;
}

ExcelReader<User> reader = new ExcelReader<>(User::new, null);

ExcelReadHandler<User> rh = reader
        .column((u, cell) -> u.name = cell.asString())
        .column((u, cell) -> u.age = cell.asInt())
        .build(Files.newInputStream(Path.of("users.xlsx")));

rh.read(result -> {
    if (result.success()) {
        User u = result.data();
        // process user
    } else {
        System.out.println(result.messages());
    }
});
```

### CSV Writing

```java
record Row(long id, String name) {}

var rows = Stream.of(new Row(1, "Alice"), new Row(2, "Bob"));

CsvWriter<Row> csv = new CsvWriter<>();
CsvHandler ch = csv
        .column("ID", r -> r.id())
        .column("Name", r -> r.name())
        .write(rows);

try (var os = Files.newOutputStream(Path.of("rows.csv"))) {
    ch.consumeOutputStream(os);
}
```

### CSV Reading

```java
class Product {
    String name;
    Integer price;
}

CsvReader<Product> csvReader = new CsvReader<>(Product::new, null);

CsvReadHandler<Product> crh = csvReader
        .column((p, cell) -> p.name = cell.asString())
        .column((p, cell) -> p.price = cell.asInt())
        .build(Files.newInputStream(Path.of("products.csv")));

crh.read(result -> {
    if (result.success()) {
        Product p = result.data();
    }
});
```

## Advanced Usage

### Name-Based Column Mapping (Reading)

Match columns by header name instead of positional index. Column order in the file doesn't matter:

```java
// Excel — matched by header name "Name" and "Age", regardless of column order
new ExcelReader<>(User::new, null)
        .column("Name", (u, cell) -> u.name = cell.asString())
        .column("Age", (u, cell) -> u.age = cell.asInt())
        .build(inputStream)
        .read(result -> { ... });

// CSV — same API
new CsvReader<>(User::new, null)
        .column("Name", (u, cell) -> u.name = cell.asString())
        .column("Age", (u, cell) -> u.age = cell.asInt())
        .build(inputStream)
        .read(result -> { ... });
```

You can also read a subset of columns — only the named columns are mapped, others are ignored:

```java
// File has columns: Name, Age, City, Email, Phone
// Only read Name and City
new ExcelReader<>(User::new, null)
        .column("Name", (u, cell) -> u.name = cell.asString())
        .column("City", (u, cell) -> u.city = cell.asString())
        .build(inputStream);
```

The `addColumn(String headerName, BiConsumer)` method also supports name-based mapping:

```java
new ExcelReader<>(User::new, null)
        .addColumn("Name", (u, cell) -> u.name = cell.asString())
        .addColumn("Age", (u, cell) -> u.age = cell.asInt())
        .build(inputStream);
```

> **Positional mapping** (existing behavior) — use `column(BiConsumer)` without a header name.

### Index-Based Column Mapping (Reading)

Map columns by explicit 0-based index — no need for `skipColumn()` chains:

```java
// Read only columns 0, 2, and 4 directly
new ExcelReader<>(User::new, null)
        .columnAt(0, (u, cell) -> u.name = cell.asString())
        .columnAt(2, (u, cell) -> u.city = cell.asString())
        .columnAt(4, (u, cell) -> u.phone = cell.asString())
        .build(inputStream);
```

Can be mixed with name-based and positional mapping:

```java
new ExcelReader<>(User::new, null)
        .column("Name", (u, cell) -> u.name = cell.asString())  // by name
        .columnAt(3, (u, cell) -> u.age = cell.asInt())          // by index
        .build(inputStream);
```

Also available for CSV:
```java
new CsvReader<>(User::new, null)
        .columnAt(0, (u, cell) -> u.name = cell.asString())
        .columnAt(2, (u, cell) -> u.city = cell.asString())
        .build(inputStream);
```

### Mapping Mode — Immutable Objects / Java Records (Reading)

Use `ExcelReader.mapping()` or `CsvReader.mapping()` to read into immutable objects (Java records, final-field classes) without setters. Instead of defining columns with `BiConsumer` setters, provide a single `Function<RowData, T>` that creates the object in one step:

```java
record PersonRecord(String name, Integer age, String city) {}

// Excel
ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
        row.get("Name").asString(),
        row.get("Age").asInt(),
        row.get("City").asString()
)).build(inputStream).read(result -> {
    if (result.success()) {
        PersonRecord p = result.data();
    }
});

// CSV — same API
CsvReader.<PersonRecord>mapping(row -> new PersonRecord(
        row.get("Name").asString(),
        row.get("Age").asInt(),
        row.get("City").asString()
)).build(inputStream).read(result -> { ... });
```

**`RowData` access methods:**

| Method | Description |
|--------|-------------|
| `get(String headerName)` | Get cell by header name (throws if not found) |
| `get(int columnIndex)` | Get cell by 0-based index (empty if out of bounds) |
| `has(String headerName)` | Check if header exists |
| `size()` | Number of cells in this row |
| `headerNames()` | List of header names |

Column order in the file doesn't matter — columns are matched by header name. No column definitions are needed:

```java
// File has columns in any order: City, Age, Name, Email
// Only read Name and City — other columns are ignored
ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
        row.get("Name").asString(),
        null,  // skip Age
        row.get("City").asString()
)).build(inputStream);
```

All read modes work with mapping: `read()`, `readStrict()`, `readAsStream()`. Configuration options (`sheetIndex`, `headerRowIndex`, `onProgress`) are also supported:

```java
ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
        row.get("Name").asString(),
        row.get("Age").asInt(),
        null
)).sheetIndex(1)
  .headerRowIndex(2)
  .onProgress(10_000, (count, cursor) -> log.info("Read {} rows", count))
  .build(inputStream)
  .read(consumer);
```

Bean Validation is supported via the second argument:

```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
ExcelReader.mapping(row -> { ... }, validator).build(inputStream);
```

**Error handling:** If the mapping function throws an exception (e.g., missing header, type conversion error), the row is reported as failed in `ReadResult` with `success() == false` and the error message. Other rows continue processing normally.

Also available via `ExcelKitSchema`:

```java
schema.excelReader(row -> {
    Person p = new Person();
    p.setName(row.get("Name").asString());
    p.setAge(row.get("Age").asInt());
    return p;
}, validator).build(inputStream);
```

### Formula Columns

Use `ExcelDataType.FORMULA` to write Excel formula cells:

```java
writer
    .column("Price", Product::price).type(ExcelDataType.INTEGER)
    .column("Quantity", Product::quantity).type(ExcelDataType.INTEGER)
    // Formula: Price * Quantity (uses cursor to compute the correct row reference)
    .column("Subtotal", (row, cursor) ->
            "D" + (cursor.getRowOfSheet() + 1) + "*E" + (cursor.getRowOfSheet() + 1))
        .type(ExcelDataType.FORMULA)
        .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
    .write(data);
```

Use `SheetContext.columnLetter()` to build formula strings in callbacks:

```java
writer
    .column("Price", Product::price).type(ExcelDataType.INTEGER)
    .afterData(ctx -> {
        var sheet = ctx.getSheet();
        int row = ctx.getCurrentRow();
        String col = SheetContext.columnLetter(0); // "A"

        var sumRow = sheet.createRow(row);
        sumRow.createCell(0).setCellFormula("SUM(%s2:%s%d)".formatted(col, col, row));

        var avgRow = sheet.createRow(row + 1);
        avgRow.createCell(0).setCellFormula("AVERAGE(%s2:%s%d)".formatted(col, col, row));

        return row + 2;
    })
    .write(data);
```

### Hyperlink Columns

Use `ExcelDataType.HYPERLINK` to create clickable URL links:

```java
// Plain URL — displayed text is the URL itself
writer
    .column("Website", Product::url)
        .type(ExcelDataType.HYPERLINK)
    .write(data);

// Custom label — use ExcelHyperlink to separate display text from URL
writer
    .column("Link", p -> new ExcelHyperlink(p.url(), "View Details"))
        .type(ExcelDataType.HYPERLINK)
    .write(data);
```

### Rich Text

Use `ExcelDataType.RICH_TEXT` to write mixed formatting within a single cell:

```java
writer
    .column("Description", p -> new ExcelRichText()
            .text("Status: ")
            .bold("APPROVED")
            .text(" — reviewed by ")
            .styled("admin", s -> s.color(ExcelColor.BLUE).italic(true)))
        .type(ExcelDataType.RICH_TEXT)
    .write(data);
```

Available FontStyle options: `bold(boolean)`, `italic(boolean)`, `underline(boolean)`, `strikethrough(boolean)`, `color(int r, int g, int b)`, `color(ExcelColor)`, `fontSize(int)`

### Auto Width Sample Rows

Column widths are auto-calculated from the first N data rows. Configurable via `autoWidthSampleRows()`:

```java
new ExcelWriter<Person>()
        .autoWidthSampleRows(200)           // sample 200 rows (default: 100)
        .column("Name", p -> p.name())
        .write(data);

new ExcelWriter<Person>()
        .autoWidthSampleRows(0)             // disable auto-width
        .column("Name", p -> p.name())
        .write(data);
```

### Row Height

```java
new ExcelWriter<Person>()
        .rowHeight(25)                      // data row height (default: 20pt)
        .column("Name", p -> p.name())
        .write(data);
```

### Header Color

```java
// RGB values
new ExcelWriter<>(91, 155, 213, 1_000_000);

// Preset color
new ExcelWriter<>(ExcelColor.STEEL_BLUE);
```

Available presets: `WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `DARK_GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `LIGHT_ORANGE`, `LIGHT_PURPLE`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`

### Column Styling

```java
writer
    .column("Amount", p -> p.amount())
        .type(ExcelDataType.DOUBLE)
        .format("#,##0.00")
        .alignment(HorizontalAlignment.RIGHT)
        .backgroundColor(ExcelColor.LIGHT_YELLOW)
        .bold(true)
        .fontSize(12)
    .column("Status", p -> p.status())
    .write(data);
```

### Dropdown Validation

Add a dropdown (data validation) to a column so users can only select from predefined options:

```java
writer
    .column("Name", p -> p.name())
    .column("Status", p -> p.status())
        .dropdown("Active", "Inactive", "Pending")
    .write(data);
```

The dropdown is applied to all data rows across all sheets (including rollover sheets).

### Row-Level Styling

Apply conditional background colors to entire rows based on data:

```java
writer
    .rowColor(p -> p.isError() ? ExcelColor.LIGHT_RED : null)  // null = no override
    .column("Name", p -> p.name())
    .column("Status", p -> p.status())
    .write(data);
```

When a row color is set, it overrides any column-level `backgroundColor`.

### Cell-Level Conditional Styling

Apply per-cell background colors based on cell value and row data using `CellColorFunction`:

```java
writer
    .column("Amount", p -> p.amount())
        .type(ExcelDataType.DOUBLE)
        .cellColor((value, row) -> {
            double amt = ((Number) value).doubleValue();
            if (amt < 0) return ExcelColor.LIGHT_RED;
            if (amt > 10000) return ExcelColor.LIGHT_GREEN;
            return null;  // no override
        })
    .write(data);
```

**Priority order:** `cellColor` > `rowColor` > column `backgroundColor`.

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Amount", Item::getAmount, c -> c
        .type(ExcelDataType.DOUBLE)
        .cellColor((value, row) ->
            ((Number) value).doubleValue() < 0 ? ExcelColor.LIGHT_RED : null))
    .write(stream);
```

### Group Headers

Create multi-row headers with merged group labels using `.group()`:

```java
writer
    .column("Name", p -> p.name())
    .column("Price", p -> p.price()).type(ExcelDataType.INTEGER).group("Financial")
    .column("Quantity", p -> p.qty()).type(ExcelDataType.INTEGER).group("Financial")
    .column("Total", p -> p.total()).type(ExcelDataType.INTEGER).group("Financial")
    .column("Notes", p -> p.notes())
    .write(data);
```

This produces:

| Name | Financial | | | Notes |
|------|-----------|----------|-------|-------|
| Name | Price | Quantity | Total | Notes |

Adjacent columns with the same group name are merged horizontally. Ungrouped columns are merged vertically across both header rows.

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Report")
    .column("Price", Item::getPrice, c -> c.type(ExcelDataType.INTEGER).group("Financial"))
    .column("Qty", Item::getQty, c -> c.type(ExcelDataType.INTEGER).group("Financial"))
    .write(stream);
```

### Column Outline (Grouping)

Group columns so they can be collapsed/expanded in Excel:

```java
writer
    .column("Name", p -> p.name())
    .column("Detail1", p -> p.detail1()).outline(1)
    .column("Detail2", p -> p.detail2()).outline(1)
    .column("Detail3", p -> p.detail3()).outline(1)
    .column("Summary", p -> p.summary())
    .write(data);
```

Adjacent columns with the same outline level are grouped together. Supports levels 1–7.

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Report")
    .column("Detail", Item::getDetail, c -> c.outline(1))
    .column("Detail2", Item::getDetail2, c -> c.outline(1))
    .write(stream);
```

### Column Hiding

Hide columns in the Excel output while still writing data:

```java
writer
    .column("ID", p -> p.id())
        .type(ExcelDataType.LONG)
    .column("Internal Code", p -> p.code())
        .hidden()                              // hidden in Excel but data is still written
    .column("Name", p -> p.name())
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Internal", Item::getCode, c -> c.hidden())
    .column("Name", Item::getName)
    .write(stream);
```

### Cell Border Style

Customize cell border styles per column (default: `THIN`):

```java
writer
    .column("Amount", p -> p.amount())
        .type(ExcelDataType.DOUBLE)
        .border(ExcelBorderStyle.MEDIUM)
    .column("Notes", p -> p.notes())
        .border(ExcelBorderStyle.DASHED)
    .column("Raw", p -> p.raw())
        .border(ExcelBorderStyle.NONE)        // no borders
    .write(data);
```

Available styles: `NONE`, `THIN`, `MEDIUM`, `THICK`, `DASHED`, `DOTTED`, `DOUBLE`, `HAIR`, `MEDIUM_DASHED`, `DASH_DOT`

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Report")
    .column("Amount", Item::getAmount, c -> c.border(ExcelBorderStyle.THICK))
    .write(stream);
```

### Per-Side Border Control

Set different border styles for each side individually. Per-side borders override the uniform `border()` setting; unset sides fall back to `border()` or default `THIN`:

```java
writer
    .column("Mixed", p -> p.value())
        .borderTop(ExcelBorderStyle.THICK)
        .borderBottom(ExcelBorderStyle.THIN)
        .borderLeft(ExcelBorderStyle.DASHED)
        .borderRight(ExcelBorderStyle.DOTTED)
    .write(data);

// Partial override: top=THICK, rest=MEDIUM
writer
    .column("Partial", p -> p.value())
        .border(ExcelBorderStyle.MEDIUM)
        .borderTop(ExcelBorderStyle.THICK)
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Report")
    .column("Amount", Item::getAmount, c -> c
        .borderTop(ExcelBorderStyle.DOUBLE)
        .borderBottom(ExcelBorderStyle.HAIR))
    .write(stream);
```

### Text Rotation

Rotate cell text counter-clockwise (positive) or clockwise (negative), from -90 to 90 degrees:

```java
writer
    .column("Rotated", p -> p.label())
        .rotation(45)       // 45° counter-clockwise
    .column("Vertical", p -> p.code())
        .rotation(90)       // straight up
    .column("Clock", p -> p.note())
        .rotation(-30)      // 30° clockwise
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Header", Item::getLabel, c -> c.rotation(45))
    .write(stream);
```

### Font Color, Strikethrough, Underline

Customize font appearance per column:

```java
writer
    .column("Warning", p -> p.message())
        .fontColor(255, 0, 0)                  // RGB red
    .column("Info", p -> p.info())
        .fontColor(ExcelColor.BLUE)            // preset color
    .column("Deleted", p -> p.oldValue())
        .strikethrough()                        // strike-through text
    .column("Important", p -> p.key())
        .underline()                            // single underline
    .column("All Styles", p -> p.summary())
        .fontColor(ExcelColor.RED)
        .bold(true)
        .underline()
        .strikethrough()
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Report")
    .column("Amount", Item::getAmount, c -> c
        .fontColor(ExcelColor.RED)
        .strikethrough()
        .underline())
    .write(stream);
```

### Vertical Alignment

Set the vertical text alignment within cells (default: `CENTER`):

```java
writer
    .column("Top", p -> p.value())
        .verticalAlignment(VerticalAlignment.TOP)
    .column("Bottom", p -> p.other())
        .verticalAlignment(VerticalAlignment.BOTTOM)
    .column("Justify", p -> p.text())
        .verticalAlignment(VerticalAlignment.JUSTIFY)
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Notes", Item::getNotes, c -> c.verticalAlignment(VerticalAlignment.TOP))
    .write(stream);
```

### Text Wrapping

Control per-column text wrapping (enabled by default). Disable to clip content at column width:

```java
writer
    .column("Description", p -> p.desc())
        .wrapText()                                 // explicitly enable (default)
    .column("Code", p -> p.code())
        .wrapText(false)                            // disable wrapping
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Code", Item::getCode, c -> c.wrapText(false))
    .write(stream);
```

### Font Name

Specify the font family for a column's cells:

```java
writer
    .column("Title", p -> p.title())
        .fontName("Arial")
    .column("한국어", p -> p.korean())
        .fontName("맑은 고딕")
    .column("Serif", p -> p.content())
        .fontName("Times New Roman")
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Name", Item::getName, c -> c.fontName("Arial"))
    .write(stream);
```

### Cell Indentation

Indent cell content by a specified level (0–250):

```java
writer
    .column("Category", p -> p.category())
    .column("Sub-item", p -> p.item())
        .indentation(2)
        .alignment(HorizontalAlignment.LEFT)
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .column("Detail", Item::getDetail, c -> c.indentation(3))
    .write(stream);
```

### Workbook Protection

Protect the workbook structure to prevent adding, deleting, renaming, or reordering sheets:

```java
new ExcelWriter<Product>()
    .protectWorkbook("password123")
    .addColumn("Name", Product::name)
    .write(data);
```

For `ExcelWorkbook`:
```java
try (var workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    workbook.protectWorkbook("password123");
    workbook.<User>sheet("Users").column("Name", User::getName).write(userStream);
    workbook.finish().consumeOutputStream(outputStream);
}
```

Can be combined with `protectSheet()` — workbook protection prevents structural changes, sheet protection prevents cell editing.

### Header Font Customization

Customize the header row font name and size (default: bold, 11pt):

```java
new ExcelWriter<Product>()
    .headerFontName("Arial")
    .headerFontSize(14)
    .addColumn("Name", Product::name)
    .write(data);
```

For `ExcelWorkbook`:
```java
try (var workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    workbook.headerFontName("맑은 고딕").headerFontSize(12);
    workbook.<User>sheet("Users").column("Name", User::getName).write(userStream);
    workbook.finish().consumeOutputStream(outputStream);
}
```

### Default Column Style

Set writer-level default styles that all columns inherit unless overridden per-column:

```java
new ExcelWriter<Product>()
    .defaultStyle(d -> d
        .fontName("Arial")
        .fontSize(10)
        .alignment(HorizontalAlignment.LEFT)
        .bold(true))
    .column("Name", p -> p.name())                       // inherits all defaults
    .column("Price", p -> p.price())
        .type(ExcelDataType.INTEGER)
        .bold(false)                                       // override: not bold
        .alignment(HorizontalAlignment.RIGHT)              // override: right-aligned
    .write(data);
```

For `ExcelSheetWriter`:
```java
workbook.<Item>sheet("Data")
    .defaultStyle(d -> d.fontName("Courier New").fontSize(9))
    .column("Code", Item::getCode)
    .column("Value", Item::getValue)
    .write(stream);
```

### Summary/Footer Rows

Add summary rows with formulas using a fluent DSL:

```java
new ExcelWriter<Product>()
    .addColumn("Name", Product::name)
    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .addColumn("Qty", p -> p.qty(), c -> c.type(ExcelDataType.INTEGER))
    .summary(s -> s
        .label("Total")
        .sum("Price")
        .sum("Qty"))
    .write(data);
```

Multiple operations create multiple summary rows:
```java
.summary(s -> s
    .sum("Amount")          // row 1: "Sum" label + SUM formula
    .average("Amount")      // row 2: "Average" label + AVERAGE formula
    .count("Amount")        // row 3: "Count" label + COUNT formula
    .min("Score")           // row 4: "Min" label + MIN formula
    .max("Score"))          // row 5: "Max" label + MAX formula
```

Place the label in a specific column:
```java
.summary(s -> s.label("Total", "Grand Total").sum("Amount"))
```

Summary rows work with sheet rollover — formulas are generated per-sheet.

### Named Ranges

Create workbook-scoped named ranges in lifecycle callbacks:

```java
writer
    .addColumn("Category", p -> p.category())
    .afterData(ctx -> {
        // By reference string
        ctx.namedRange("Categories", "Sheet1!$A$2:$A$100");

        // By column index and row range (auto-generates reference for current sheet)
        ctx.namedRange("CategoryList", 0, 1, ctx.getCurrentRow() - 1);

        return ctx.getCurrentRow();
    })
    .write(data);
```

Named ranges can be used as validation sources with `ExcelValidation.listFromRange()`.

### List Validation from Cell Range

Create dropdown validations that reference a cell range instead of inline string arrays:

```java
writer
    .column("Status", p -> p.status())
        .validation(ExcelValidation.listFromRange("Options!$A$1:$A$5"))
    .write(data);
```

Useful for large option lists stored on a separate sheet, or dynamic lists that can be updated independently.

### Advanced Data Validation

Apply advanced data validation rules beyond dropdowns:

```java
writer
    .column("Age", p -> p.age())
        .type(ExcelDataType.INTEGER)
        .validation(ExcelValidation.integerBetween(0, 150))
    .column("GPA", p -> p.gpa())
        .type(ExcelDataType.DOUBLE)
        .validation(ExcelValidation.decimalBetween(0.0, 4.0))
    .column("Name", p -> p.name())
        .validation(ExcelValidation.textLength(1, 100))
    .column("Date", p -> p.date())
        .type(ExcelDataType.DATE)
        .validation(ExcelValidation.dateRange(
            LocalDate.of(2024, 1, 1),
            LocalDate.of(2024, 12, 31)))
    .column("Custom", p -> p.value())
        .validation(ExcelValidation.formula("AND(A2>0,A2<100)"))
    .write(data);
```

Available factory methods:
- `ExcelValidation.integerBetween(min, max)` — integer range (inclusive)
- `ExcelValidation.integerGreaterThan(min)` — integer greater than
- `ExcelValidation.integerLessThan(max)` — integer less than
- `ExcelValidation.decimalBetween(min, max)` — decimal range (inclusive)
- `ExcelValidation.textLength(min, max)` — text length constraint
- `ExcelValidation.dateRange(start, end)` — date range constraint
- `ExcelValidation.formula(formula)` — custom Excel formula
- `ExcelValidation.listFromRange(range)` — dropdown from cell range (e.g., `"Sheet2!$A$1:$A$10"`)

Fluent error configuration:
```java
ExcelValidation.integerBetween(1, 100)
    .errorTitle("Invalid Value")
    .errorMessage("Please enter a number between 1 and 100")
    .showError(true)    // default: true
```

### Row Grouping/Outlining

Group rows so they can be collapsed/expanded in Excel. Use in `afterData` or `afterAll` callbacks via `SheetContext`:

```java
writer
    .addColumn("Data", p -> p.value())
    .afterData(ctx -> {
        ctx.groupRows(1, 5);                    // group rows 1-5
        ctx.groupRows(7, 10, true);             // group rows 7-10 and collapse
        return ctx.getCurrentRow();
    })
    .write(data);
```

Methods:
- `groupRows(int firstRow, int lastRow)` — group rows (expanded)
- `groupRows(int firstRow, int lastRow, boolean collapsed)` — group rows with optional collapse

### Sheet Tab Color

Colorize the sheet tab in the workbook:

```java
// ExcelWriter — applies to all sheets (including rollover)
new ExcelWriter<Product>()
    .tabColor(255, 0, 0)                       // RGB red
    .addColumn("Name", Product::name)
    .write(data);

new ExcelWriter<Product>()
    .tabColor(ExcelColor.STEEL_BLUE)           // preset color
    .addColumn("Name", Product::name)
    .write(data);
```

For `ExcelSheetWriter` — set per sheet:
```java
workbook.<User>sheet("Users")
    .tabColor(ExcelColor.BLUE)
    .column("Name", User::getName)
    .write(userStream);

workbook.<Order>sheet("Orders")
    .tabColor(ExcelColor.GREEN)
    .column("ID", Order::getId)
    .write(orderStream);
```

### Cell Comments (Notes)

Add conditional cell comments (notes) to specific columns:

```java
writer
    .column("Score", p -> p.score())
        .type(ExcelDataType.INTEGER)
        .comment(p -> p.score() < 50 ? "Low score - needs review" : null)
    .write(data);
```

The comment function receives the row data and returns a comment string, or `null` to skip. Comments appear as yellow note icons in Excel.

### Conditional Formatting

Apply Excel conditional formatting rules:

```java
new ExcelWriter<Product>()
    .addColumn("Name", Product::name)
    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .conditionalFormatting(cf -> cf
        .columns(1)                                    // apply to column 1 only
        .greaterThan("10000", ExcelColor.LIGHT_RED)    // highlight expensive
        .lessThan("1000", ExcelColor.LIGHT_GREEN)      // highlight cheap
        .between("5000", "10000", ExcelColor.LIGHT_YELLOW))
    .write(data);
```

Available operators: `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equalTo`, `notEqualTo`, `between`, `notBetween`

If `columns()` is not set, rules apply to all columns.

### Sheet Protection

Protect sheets with a password and selectively unlock columns:

```java
new ExcelWriter<Product>()
    .addColumn("Name", Product::name, c -> c.locked(false))   // editable
    .addColumn("Price", p -> p.price(), c -> c.locked(true))  // read-only
    .protectSheet("password123")
    .write(data);
```

When sheet protection is enabled, all cells are locked by default. Use `.locked(false)` on specific columns to allow editing.

### Image Embedding

Embed images in Excel cells using `ExcelDataType.IMAGE`:

```java
// From byte array
byte[] imageBytes = Files.readAllBytes(Path.of("logo.png"));

new ExcelWriter<Product>()
    .addColumn("Name", Product::name)
    .addColumn("Photo", p -> ExcelImage.png(imageBytes))
        .type(ExcelDataType.IMAGE)
    .write(data);
```

Factory methods: `ExcelImage.png(byte[])`, `ExcelImage.jpeg(byte[])`, or `new ExcelImage(byte[], pictureType)` for other types.

### Chart Generation

Add charts after data is written:

```java
new ExcelWriter<Product>()
    .addColumn("Name", Product::name)
    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
    .addColumn("Profit", p -> p.profit(), c -> c.type(ExcelDataType.INTEGER))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.BAR)       // BAR, LINE, PIE, SCATTER, AREA, or DOUGHNUT
        .title("Sales vs Profit")
        .categoryColumn(0)                           // X-axis: Name column
        .valueColumn(1, "Sales")                     // Y-axis series 1
        .valueColumn(2, "Profit")                    // Y-axis series 2
        .categoryAxisTitle("Product")                // X-axis label
        .valueAxisTitle("Amount")                    // Y-axis label
        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM)
        .barDirection(ExcelChartConfig.BarDirection.HORIZONTAL)
        .barGrouping(ExcelChartConfig.BarGrouping.STACKED)
        .showDataLabels(true)
        .position(3, 0, 12, 20))                     // chart position (col1, row1, col2, row2)
    .write(data);
```

Charts are created using Apache POI's XDDF chart API and reference data cell ranges.

Available chart options:
- Chart types: `BAR`, `LINE`, `PIE`, `SCATTER`, `AREA`, `DOUGHNUT`
- Legend positions: `BOTTOM`, `LEFT`, `RIGHT`, `TOP`, `TOP_RIGHT`
- Bar directions: `VERTICAL` (default), `HORIZONTAL`
- Bar groupings: `STANDARD` (default), `STACKED`, `PERCENT_STACKED`

**Scatter chart** — both axes are numeric (no category axis):
```java
writer
    .addColumn("X", p -> p.x(), c -> c.type(ExcelDataType.DOUBLE))
    .addColumn("Y", p -> p.y(), c -> c.type(ExcelDataType.DOUBLE))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.SCATTER)
        .categoryColumn(0)          // X-axis data
        .valueColumn(1, "Y values")
        .categoryAxisTitle("X Axis")
        .valueAxisTitle("Y Axis"))
    .write(data);
```

**Area chart** — like line chart but with filled regions:
```java
writer
    .addColumn("Month", Product::month)
    .addColumn("Revenue", p -> p.revenue(), c -> c.type(ExcelDataType.DOUBLE))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.AREA)
        .categoryColumn(0)
        .valueColumn(1, "Revenue"))
    .write(data);
```

**Doughnut chart** — like pie chart with a hollow center:
```java
writer
    .addColumn("Category", Product::category)
    .addColumn("Share", p -> p.share(), c -> c.type(ExcelDataType.INTEGER))
    .chart(chart -> chart
        .type(ExcelChartConfig.ChartType.DOUGHNUT)
        .categoryColumn(0)
        .valueColumn(1, "Share")
        .showDataLabels(true))
    .write(data);
```

### Map-Based Writing

Write `Map<String, Object>` data without defining typed POJOs:

```java
ExcelMapWriter writer = new ExcelMapWriter("Name", "Age", "City");
writer.write(Stream.of(
    Map.of("Name", "Alice", "Age", 30, "City", "Seoul"),
    Map.of("Name", "Bob", "Age", 25, "City", "Tokyo")
)).consumeOutputStream(outputStream);
```

CSV equivalent:
```java
CsvMapWriter csvWriter = new CsvMapWriter("Name", "Age");
csvWriter.write(stream).consumeOutputStream(outputStream);
```

### Map-Based Reading

Read Excel files into `Map<String, String>` without defining typed POJOs:

```java
new ExcelMapReader()
    .sheetIndex(0)         // optional, defaults to 0
    .headerRowIndex(0)     // optional, defaults to 0
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
        String name = row.get("Name");
        String age = row.get("Age");
    });
```

All columns from the header row are automatically mapped.

### Multi-Sheet Discovery

Discover sheet names and read specific sheets:

```java
// Get all sheet names and indices
List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(inputStream);
sheets.forEach(s -> System.out.println(s.index() + ": " + s.name()));

// Get header names from a specific sheet
List<String> headers = ExcelReader.getSheetHeaders(inputStream, 0, 0);
```

Combined with `sheetIndex()` to iterate over all sheets:
```java
for (ExcelSheetInfo sheet : ExcelReader.getSheetNames(new FileInputStream(file))) {
    new ExcelReader<>(Row::new, null)
        .sheetIndex(sheet.index())
        .column("Name", (r, cell) -> r.name = cell.asString())
        .build(new FileInputStream(file))
        .read(consumer);
}
```

### Progress Callback

Monitor progress during large dataset writes:

```java
writer
    .column("Name", p -> p.name())
    .onProgress(10_000, (count, cursor) ->
        log.info("Processed {} rows", count))
    .write(data);
```

The callback fires every `interval` rows. Works with both `ExcelWriter` and `ExcelSheetWriter`, including across sheet rollovers.

### Conditional Columns

```java
boolean showAge = true;

writer
    .column("Name", p -> p.name())
    .columnIf("Age", showAge, p -> p.age())  // only added when condition is true
    .column("Email", p -> p.email())
    .write(data);
```

### Constant Columns

```java
writer
    .column("Name", p -> p.name())
    .constColumn("Source", "SYSTEM")  // same value for every row
    .write(data);
```

### Auto-Filter and Freeze Panes

```java
writer
    .autoFilter(true)    // dropdown filter on header row
    .freezePane(1)       // freeze 1 row below the header
    .column("Name", p -> p.name())
    .write(data);
```

### Lifecycle Callbacks

Callbacks allow inserting custom rows at specific points during Excel generation.

**Invocation order per sheet:**
```
beforeHeader → header → data rows → afterData
                                      ↓ (last sheet only)
                                    afterAll
```

On sheet rollover: current sheet gets `afterData` → new sheet is created → `beforeHeader` + header.

All callbacks receive a `SheetContext` parameter that provides:
- `getSheet()` — the current `SXSSFSheet`
- `getWorkbook()` — the `SXSSFWorkbook` (useful for creating CellStyles, etc.)
- `getCurrentRow()` — the first available row index for writing
- `getColumnCount()` — the number of configured columns
- `getColumnNames()` — unmodifiable list of column header names
- `columnLetter(int)` — static helper to convert column index to Excel letter (0→"A", 26→"AA")
- `mergeCells(int firstRow, int lastRow, int firstCol, int lastCol)` — merge a rectangular cell region
- `mergeCells(String cellRange)` — merge cells using Excel notation (e.g., `"A1:C3"`)
- `groupRows(int firstRow, int lastRow)` — group rows for collapsible outlining
- `groupRows(int firstRow, int lastRow, boolean collapsed)` — group rows with optional initial collapse

A new `SheetContext` is created for each callback invocation, so the sheet reference is always current (even after rollover).

```java
writer
    .column("Name", p -> p.name())
    .column("Amount", p -> p.amount())
    .beforeHeader(ctx -> {
        // called on every sheet, before the header row
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
                .setCellValue("Generated: 2025-07-19");
        return ctx.getCurrentRow() + 1;  // return the next available row index
    })
    .afterData(ctx -> {
        // called on every sheet after its data rows (e.g. subtotals)
        SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
        row.createCell(0).setCellValue("Subtotal");
        return ctx.getCurrentRow() + 1;
    })
    .afterAll(ctx -> {
        // called once on the last sheet, after afterData (e.g. grand total)
        SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
        row.createCell(0).setCellValue("Grand Total");
        return ctx.getCurrentRow() + 1;
    })
    .write(data);
```

### Cell Merging

Merge cells in lifecycle callbacks using `SheetContext.mergeCells()`:

```java
writer
    .beforeHeader(ctx -> {
        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0)
                .setCellValue("Report Title");
        ctx.mergeCells(0, 0, 0, 2);  // merge first row across 3 columns
        return ctx.getCurrentRow() + 1;
    })
    .column("A", p -> p.a())
    .column("B", p -> p.b())
    .column("C", p -> p.c())
    .write(data);
```

You can also use Excel notation:
```java
ctx.mergeCells("A1:C1");  // same merge using Excel notation
```

### Sheet Auto-Splitting

When the configured maximum rows per sheet is reached, a new sheet is automatically created with `beforeHeader` and header replicated.

```java
// split every 100,000 rows
new ExcelWriter<>(ExcelColor.STEEL_BLUE, 100_000);
```

### Password-Encrypted Export

```java
try (var os = Files.newOutputStream(Path.of("secret.xlsx"))) {
    handler.consumeOutputStreamWithPassword(os, "P@ssw0rd!");
}
```

### Explicit Multi-Sheet Workbook

Use `ExcelWorkbook` to write different data types to separate sheets:

```java
try (ExcelWorkbook workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    workbook.<User>sheet("Users")
        .column("Name", u -> u.getName())
        .column("Status", u -> u.getStatus(), c -> c.dropdown("Active", "Inactive"))
        .column("Age", u -> u.getAge(), c -> c.type(ExcelDataType.INTEGER))
        .rowColor(u -> u.isError() ? ExcelColor.LIGHT_RED : null)
        .write(userStream);

    workbook.<Order>sheet("Orders")
        .column("ID", o -> o.getId())
        .column("Amount", o -> o.getAmount(), c -> c.type(ExcelDataType.DOUBLE))
        .write(orderStream);

    ExcelHandler handler = workbook.finish();
    handler.consumeOutputStream(outputStream);
}
```

Each `ExcelSheetWriter` supports the same features as `ExcelWriter`:
- Column configuration via `Consumer<ColumnConfig>`: `type`, `format`, `alignment`, `verticalAlignment`, `backgroundColor`, `bold`, `fontSize`, `fontName`, `width`, `minWidth`, `maxWidth`, `dropdown`, `cellColor`, `group`, `outline`, `comment`, `border`, `borderTop`, `borderBottom`, `borderLeft`, `borderRight`, `locked`, `hidden`, `rotation`, `fontColor`, `strikethrough`, `underline`, `wrapText`, `indentation`, `validation`
- `beforeHeader()`, `afterData()`, `autoFilter()`, `freezePane()`, `rowColor()`, `constColumn()`, `columnIf()`, `onProgress()`, `protectSheet()`, `conditionalFormatting()`, `chart()` (with full chart options: axis titles, legend position, bar direction, bar grouping, data labels), `printSetup()`, `tabColor()`, `defaultStyle()`, `summary()`

**Sheet auto-rollover** — `ExcelSheetWriter` can also auto-split sheets via `maxRows()`:

```java
workbook.<Order>sheet("Orders")
    .maxRows(500_000)
    .sheetName(index -> "Orders-Page" + (index + 1))  // custom rollover sheet naming
    .column("ID", Order::getId)
    .column("Amount", Order::getAmount, c -> c.type(ExcelDataType.DOUBLE))
    .write(orderStream);
// Creates: "Orders", "Orders-Page2", "Orders-Page3", ... as needed
```

If `sheetName()` is not set, rollover sheets are named `"baseName (2)"`, `"baseName (3)"`, etc.

### Cursor Access

The `Cursor` provides positional information during streaming. Use `ExcelRowFunction` (with cursor) instead of `Function`:

```java
writer
    .column("No.", (row, cursor) -> cursor.getCurrentTotal())  // row number
        .type(ExcelDataType.LONG)
    .column("Name", (row, cursor) -> row.name())
    .write(data);
```

- `cursor.getCurrentTotal()` — total rows written (across all sheets, `long`)
- `cursor.getRowOfSheet()` — current row index within the sheet

### Excel Reading — Advanced Options

**Header row index** (for files with metadata rows above the header):
```java
new ExcelReader<>(User::new, null)
        .headerRowIndex(2)  // use the 3rd row as header (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .build(inputStream);
```

**Specific sheet:**
```java
new ExcelReader<>(User::new, null)
        .sheetIndex(1)  // read the 2nd sheet (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .build(inputStream);
```

**Skip columns:**
```java
reader
    .column((u, cell) -> u.name = cell.asString())
    .skipColumn()       // skip one column
    .skipColumns(2)     // skip two more columns
    .column((u, cell) -> u.age = cell.asInt())
    .build(inputStream);
```

**Stream-based reading:**
```java
try (Stream<ReadResult<User>> stream = rh.readAsStream()) {
    stream.filter(ReadResult::success)
          .map(ReadResult::data)
          .forEach(user -> { /* process */ });
}
```

**Large file support:**
```java
// call once at application startup
ExcelReader.configureLargeFileSupport();
```

**Read progress callback:**
```java
new ExcelReader<>(User::new, null)
        .column((u, cell) -> u.name = cell.asString())
        .onProgress(10_000, (count, cursor) ->
            log.info("Read {} rows", count))
        .build(inputStream)
        .read(consumer);
```

**Bean Validation:**
```java
Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
ExcelReader<User> reader = new ExcelReader<>(User::new, validator);
```

### CellData Conversion Methods

When reading Excel/CSV, `CellData` provides type-safe conversions:

| Method | Return Type |
|--------|-------------|
| `asString()` | `String` |
| `asInt()` | `Integer` |
| `asLong()` | `Long` |
| `asDouble()` | `Double` |
| `asFloat()` | `Float` |
| `asBigDecimal()` | `BigDecimal` |
| `asBoolean()` | `boolean` (`true`/`1`/`y`/`yes`) |
| `asBooleanOrNull()` | `Boolean` |
| `asLocalDate()` | `LocalDate` |
| `asLocalDateTime()` | `LocalDateTime` |
| `asLocalTime()` | `LocalTime` |
| `asEnum(Class<E>)` | `E` (case-insensitive name match) |
| `as(Function<String, R>)` | `R` (custom conversion, null if blank) |
| `isEmpty()` | `boolean` |

**Default value overloads** — return the given default instead of `null` when the cell is blank:

| Method | Return Type |
|--------|-------------|
| `asInt(int defaultValue)` | `int` |
| `asLong(long defaultValue)` | `long` |
| `asDouble(double defaultValue)` | `double` |
| `asString(String defaultValue)` | `String` |
| `as(Function<String, R>, R defaultValue)` | `R` |

**Custom conversion examples:**
```java
// UUID parsing
UUID id = cell.as(UUID::fromString);

// Custom domain object
MyType obj = cell.as(MyType::parse);

// With default value
UUID id = cell.as(UUID::fromString, DEFAULT_UUID);
int qty = cell.asInt(0);       // 0 if blank
String name = cell.asString("N/A");  // "N/A" if blank
```

Custom date formats:
```java
CellData.addDateFormat("dd/MM/yyyy");
CellData.addDateTimeFormat("dd/MM/yyyy HH:mm");
CellData.resetDateFormats();      // restore defaults
CellData.resetDateTimeFormats();
```

Number parsing locale (default: `Locale.KOREA`):
```java
CellData.setDefaultLocale(Locale.US);
```

### ExcelDataType Reference

| Type | Java Type | Default Format |
|------|-----------|---------------|
| `STRING` | String | — |
| `BOOLEAN_TO_YN` | Boolean → "Y"/"N" | — |
| `LONG` | Long | `#,##0` |
| `INTEGER` | Integer | `#,##0` |
| `DOUBLE` | Double | `#,##0.00` |
| `FLOAT` | Float | `#,##0.00` |
| `DOUBLE_PERCENT` | Double | `0.00%` |
| `FLOAT_PERCENT` | Float | `0.00%` |
| `DATETIME` | LocalDateTime | `yyyy-mm-dd hh:mm:ss` |
| `DATE` | LocalDate | `yyyy-mm-dd` |
| `TIME` | LocalTime | `hh:mm:ss` |
| `BIG_DECIMAL_TO_DOUBLE` | BigDecimal | `#,##0.00` |
| `BIG_DECIMAL_TO_LONG` | BigDecimal | `#,##0` |
| `FORMULA` | String (formula) | — |
| `HYPERLINK` | String or `ExcelHyperlink` | — |
| `IMAGE` | `ExcelImage` | — |
| `RICH_TEXT` | `ExcelRichText` | — |

### ExcelDataFormat Presets

Use with `.format(ExcelDataFormat.NUMBER.getFormat())`:

| Preset | Format String |
|--------|---------------|
| `NUMBER` | `#,##0` |
| `NUMBER_1` | `#,##0.0` |
| `NUMBER_2` | `#,##0.00` |
| `NUMBER_4` | `#,##0.0000` |
| `PERCENT` | `0.00%` |
| `DATETIME` | `yyyy-mm-dd hh:mm:ss` |
| `DATE` | `yyyy-mm-dd` |
| `TIME` | `hh:mm:ss` |
| `CURRENCY_KRW` | `#,##0"원"` |
| `CURRENCY_USD` | `"$"#,##0.00` |

### ExcelKitSchema — Unified Read/Write Definitions

Define columns once, use for both reading and writing:

```java
ExcelKitSchema<Book> schema = ExcelKitSchema.<Book>builder()
        .column("Title",  Book::getTitle,  (b, cell) -> b.setTitle(cell.asString()))
        .column("Price",  Book::getPrice,  (b, cell) -> b.setPrice(cell.asInt()),
                c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
        .build();

// writing — type/format config is applied automatically
ExcelHandler handler = schema.excelWriter().write(bookStream);
CsvHandler ch = schema.csvWriter().write(bookStream);

// reading (setter mode) — columns are matched by header name (order-independent)
ExcelReadHandler<Book> rh = schema.excelReader(Book::new, null)
        .build(inputStream);
CsvReadHandler<Book> crh = schema.csvReader(Book::new, null)
        .build(inputStream);

// reading (mapping mode) — for immutable objects
schema.excelReader(row -> new BookRecord(
        row.get("Title").asString(), row.get("Price").asInt()
), null).build(inputStream);
```

The write configurer receives an `ExcelColumnBuilder` — use configuration methods only (`type`, `format`, `alignment`, `verticalAlignment`, `backgroundColor`, `bold`, `fontSize`, `fontName`, `width`, `minWidth`, `maxWidth`, `dropdown`, `cellColor`, `group`, `outline`, `comment`, `border`, `borderTop`, `borderBottom`, `borderLeft`, `borderRight`, `locked`, `hidden`, `rotation`, `fontColor`, `strikethrough`, `underline`, `wrapText`, `indentation`, `validation`). Writer-level features like `defaultStyle`, `summary`, `protectWorkbook`, `headerFontName`, `headerFontSize` are available on `ExcelWriter` and `ExcelWorkbook`.

```java
ExcelKitSchema.<Product>builder()
    .column("Name", Product::getName, (p, cell) -> p.setName(cell.asString()))
    .column("Price", Product::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
            c -> c.type(ExcelDataType.INTEGER)
                  .format(ExcelDataFormat.CURRENCY_KRW.getFormat())
                  .backgroundColor(ExcelColor.LIGHT_YELLOW))
    .column("Discount", Product::getDiscount, (p, cell) -> p.setDiscount(cell.asDouble()),
            c -> c.type(ExcelDataType.DOUBLE_PERCENT))
    .build();
```

### CSV Options

```java
new CsvWriter<Row>()
    .delimiter('\t')                   // tab-separated
    .charset(StandardCharsets.UTF_16)  // custom encoding
    .bom(false)                        // disable UTF-8 BOM
    .column("Name", r -> r.name())
    .write(rows);

new CsvReader<>(Row::new, null)
    .delimiter('\t')
    .charset(StandardCharsets.UTF_16)
    .headerRowIndex(1)
    .column((r, cell) -> r.name = cell.asString())
    .build(inputStream);
```

### Print Setup

Configure print layout for Excel sheets:

```java
new ExcelWriter<Invoice>()
    .printSetup(ps -> ps
        .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
        .paperSize(ExcelPrintSetup.PaperSize.A4)
        .margins(0.5, 0.5, 0.75, 0.75)
        .headerCenter("Invoice Report")
        .footerCenter("Page &P of &N")
        .repeatHeaderRows()
        .fitToPageWidth())
    .column("Invoice #", Invoice::getNumber)
    .write(data);
```

Available options:
- Orientation: `PORTRAIT`, `LANDSCAPE`
- Paper sizes: `LETTER`, `LEGAL`, `A3`, `A4`, `A5`, `B4`, `B5`
- Margins: `margins(left, right, top, bottom)` or individual `leftMargin()`, `rightMargin()`, `topMargin()`, `bottomMargin()` (in inches)
- Headers/Footers: `headerLeft()`, `headerCenter()`, `headerRight()`, `footerLeft()`, `footerCenter()`, `footerRight()`
- Special codes: `&P` (page number), `&N` (total pages), `&D` (date), `&T` (time), `&F` (filename)
- Repeat rows: `repeatHeaderRows()` or `repeatRows(firstRow, lastRow)`
- Fit to page: `fitToPageWidth()` or `fitToPage(width, height)`

## Spring MVC Integration

The `example` module includes `ExcelResponse` and `CsvResponse` helpers that wrap
`ExcelHandler`/`CsvHandler` into a `ResponseEntity<StreamingResponseBody>` with
proper Content-Type, Content-Disposition (including RFC 5987 Korean filename encoding),
and Cache-Control headers.

```java
@GetMapping("/download")
public ResponseEntity<StreamingResponseBody> download() {
    ExcelHandler handler = writer.write(dataStream);
    return ExcelResponse.of(handler, "report");
}

@GetMapping("/download-csv")
public ResponseEntity<StreamingResponseBody> downloadCsv() {
    CsvHandler handler = csvWriter.write(dataStream);
    return CsvResponse.of(handler, "report");
}

// Password-encrypted download
@GetMapping("/download-encrypted")
public ResponseEntity<StreamingResponseBody> downloadEncrypted() {
    ExcelHandler handler = writer.write(dataStream);
    return ExcelResponse.of(handler, "secret", "P@ssw0rd!");
}
```

## Spring WebFlux Integration

Since Apache POI is blocking I/O, wrap the write operation on `boundedElastic`:

```java
@GetMapping("/download")
public Mono<Void> download(ServerHttpResponse response) {
    response.getHeaders().setContentType(
        MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
    response.getHeaders().set(HttpHeaders.CONTENT_DISPOSITION,
        "attachment; filename=\"report.xlsx\"");

    return response.writeWith(DataBufferUtils.readInputStream(
        () -> {
            PipedInputStream pis = new PipedInputStream();
            PipedOutputStream pos = new PipedOutputStream(pis);
            Schedulers.boundedElastic().schedule(() -> {
                try {
                    writer.write(dataStream).consumeOutputStream(pos);
                    pos.close();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            });
            return pis;
        },
        response.bufferFactory(), 8192));
}
```

When using reactive repositories (`Flux<T>`), convert to `Stream` with `flux.toStream()`:

```java
Flux<MyData> flux = repository.findAll();
ExcelHandler handler = writer.write(flux.toStream());
```

`Flux.toStream()` handles backpressure internally, so data is not loaded entirely into memory.

## Exception Handling

| Exception | Description |
|-----------|-------------|
| `ExcelKitException` | Base class for all library exceptions |
| `ExcelWriteException` | Excel write errors (no columns, handler already consumed, etc.) |
| `ExcelReadException` | Excel read/parse errors |
| `CsvWriteException` | CSV write errors |
| `CsvReadException` | CSV read/parse errors |

- Column mapping exceptions are safely logged and fall back to string conversion (Excel writing).
- Calling `consumeOutputStream` on an already-consumed handler throws the corresponding `WriteException`.
- Empty password on encrypted export throws `IllegalArgumentException`.

## Security & Resource Management

### Temporary File Handling

The library creates temporary files during read and write operations (e.g., SAX-based Excel parsing, password encryption, CSV export). All temporary resources are managed with the following guarantees:

- **Secure creation:** On POSIX systems, temp directories are created with `rwx------` (owner-only) permissions. On Windows, ACLs are restricted to the current user.
- **Automatic cleanup:** Temp files and directories are deleted immediately after each operation completes (success or failure).
- **Fallback on failure:** If immediate deletion fails (e.g., Windows file lock), the library logs a warning and registers `deleteOnExit()` so the JVM cleans up on shutdown.
- **UUID-based naming:** Temp files use UUID-based names to prevent path prediction.

### Password Encryption

- Uses Apache POI's **Agile encryption mode** (AES-256, modern Excel standard).
- The `char[]` password overload zeroes the array after use to minimize password exposure in memory.
- Encryption uses a two-stage write (workbook → temp file → encrypted output) to keep memory usage low.

> **Note:** Excel sheet protection (`protectSheet()`) is a UI-level deterrent, not cryptographic security. It prevents casual editing but can be bypassed by tools. Use file-level password encryption (`consumeOutputStreamWithPassword`) for actual data protection.

### CSV Injection Defense

`CsvWriter` automatically defends against CSV formula injection by prefixing dangerous leading characters (`=`, `+`, `-`, `@`, `\t`, `\r`) with a single quote (`'`). This prevents malicious payloads from being executed when the CSV is opened in spreadsheet applications.

The defense can be disabled when writing trusted data where the prefix would corrupt values:

```java
new CsvWriter<Row>()
    .csvInjectionDefense(false)   // disable for trusted data
    .column("Formula", r -> r.formula())
    .write(rows);
```

### Formula Columns

`ExcelDataType.FORMULA` writes cell values as Excel formulas via `setCellFormula()`. This type is intended for developer-controlled formula strings (e.g., `SUM(A2:A100)`). Do **not** pass untrusted user input as formula values — use `ExcelDataType.STRING` for user-supplied data instead.

## Requirements

- **JDK 17+**
- Apache POI 5.x for Excel operations

## Build & Test

```bash
./gradlew build
./gradlew test
```

## CI/CD

| Workflow | Trigger | Description |
|----------|---------|-------------|
| **CI** | Push to `main`, Pull requests | Build + test |
| **Release** | Tag push (`*.*.*`) | Build, test, create GitHub Release with auto-generated notes |
| **Maven Publish** | Tag push (`*.*.*`) | Publish to Maven Central |

## License

MIT License. See [LICENSE](./LICENSE) for details.
