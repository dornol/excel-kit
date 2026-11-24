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

`writeTo(Path)` (including the password overloads) writes through a temporary sibling file and
replaces the target only when the write succeeded, via `AtomicFileWriter`. If the write fails, or the
handler was already consumed, an existing target file is left untouched and no partial file remains.

**Releasing a handler you will not write.** `write(...)` returns an `ExcelHandler` / `CsvHandler`
that owns temp files until `writeTo(...)` consumes it. If the export is cancelled, or you hit an
error after creating the handler, release it with `discard()`:

```java
ExcelHandler handler = writer.write(rows);
if (!stillWanted()) {
    handler.discard();   // idempotent; a no-op after writeTo
    return;
}
handler.writeTo(out);
```

`discard()` is on the shared `FileHandler` interface, so code that holds either handler releases it the same way.
After `discard()` the `writeTo` methods throw.

**Neither handler is `AutoCloseable`, on purpose.** A handler is usually written after the method that made it has
returned: a Spring `StreamingResponseBody` runs once the controller is done. `try (var handler = ...) { return
ResponseEntity.ok(handler::writeTo); }` would release the file when the block ends, and the download would fail. Call
`discard()` yourself where a handler is dropped.

As a safety net a handler that is garbage collected
without being consumed releases its temp files too, but only when the garbage collector runs, so do
not rely on it. `ExcelKitResponse` in the Spring module discards the handler it was given when it
rejects a blank password.

CSV writes report a `CsvWriteException` when a value cannot be encoded in the
selected charset. With formula-injection defense enabled, numeric objects keep
their signs and formula-like text is prefixed with a single quote.
NUL characters and exceptions from value extractors raise a
`CsvWriteException`. Use `.writeErrorPolicy(WriteErrorPolicy.LENIENT)` to write an empty field
(or the column's `nullValue`) when an extractor fails: the same setting, with the same two values and the same default
(`FAIL_FAST`), as on the Excel writers. Null and empty string values both write as empty CSV fields unless a column has a `nullValue`.

## Data Types

Types are inferred from the first non-null value when `.type(...)` is omitted. Use
`.type(...)` when a column needs an explicit or semantic type; explicit types always win.
Columns whose values are all `null` fall back to `STRING`.

Inference recognizes `String`, `Boolean`, the integer and floating-point `Number` types,
`BigDecimal`, `LocalDate`, `LocalDateTime` and `LocalTime`. These date types are also written as
Excel dates: `ZonedDateTime` and `OffsetDateTime` (their own wall-clock time), `java.util.Date`,
`java.sql.Timestamp`, `java.sql.Date` and `java.sql.Time` (JVM default time zone).
`Instant` has no time zone of its own, so it is not guessed: like any other type, including
`YearMonth`, it is written as its `toString()` text. Convert it to a local type in the column
function to get an Excel date, for example `c -> c.getCreatedAt().atZone(zone).toLocalDateTime()`.
An explicit `DATETIME` column also accepts a `LocalDate` and writes it as midnight.

When an inferred column encounters a value of a different type, the writer
preserves that value's type where possible. An inferred integer column can store
later fractional values as numbers.
When a value exceeds Excel's numeric precision or the range or precision of
an explicitly selected integer type, the default `FAIL_FAST` policy rejects it with `ExcelWriteException`.
Use `LENIENT` to write those values as text. Template cell writes also use text
when a value exceeds Excel's numeric precision.

`LocalTime` values are stored as a fraction of a day. Dates outside the workbook's
date system and date-times with sub-millisecond precision are preserved as ISO text
with `LENIENT` and rejected with the default `FAIL_FAST`. Template cell writes preserve such
values as ISO text.
Default date-time and time styles show fractional seconds to milliseconds when
present. An explicit column format controls the cell's display precision.

Built-in formula columns and summary rows ask Excel to recalculate on open.
Their results read as blank until an application calculates and saves them.

XLSX cell text preserves literal escape-like text such as `_x0041_` and control
characters by encoding them as SpreadsheetML text escapes.

You can still set types explicitly, for example `ExcelDataType.INTEGER` or `ExcelDataType.DATE`:

| Type | Java Type | Default Format |
|------|-----------|----------------|
| `STRING` | String or unknown value | — |
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

## Write Error Policy

By default, Excel writing is fail-fast (like CSV): a value extraction failure or a cell
type mismatch aborts the export with `ExcelWriteException`. Use lenient mode when a bad
cell should become a blank cell (extraction failure) or text (type mismatch) instead:

```java
ExcelWriter.<Product>create()
    .writeErrorPolicy(WriteErrorPolicy.LENIENT)
    .column("Price", Product::price, c -> c.type(ExcelDataType.INTEGER))
    .write(products);
```

This policy controls cell extraction and conversion. Configured sheet features
are not silently omitted when backing-sheet resolution fails. Resolution uses
public POI workbook APIs and throws `ExcelWriteException` for removed or unregistered
sheets, preserving underlying causes. Data-bar, icon-set and chart data-label
failures also throw `ExcelWriteException` rather than being logged or ignored.

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

All column configuration methods are documented in [Styling](styling.md). The configurer is a
`Consumer<ColumnConfig<T>>` for every writer, so a setting can be kept in a variable and used on any of them:

```java
Consumer<ColumnConfig<Order>> money = c -> c.type(ExcelDataType.INTEGER).format("#,##0").alignment(HorizontalAlignment.RIGHT);

ExcelWriter.<Order>create().column("Total", Order::total, money);
workbook.<Order>sheet("Orders").column("Total", Order::total, money);
```

`style(CellStyle)` gives a column a cell style built by hand (it must belong to the same workbook); the other
settings of the column still apply.

## Settings shared by `ExcelWriter` and `ExcelSheetWriter`

`ExcelWriter` (one sheet, rolled over when full) and `ExcelSheetWriter` (a sheet of an `ExcelWorkbook`) have the same
columns and sheet settings: they extend `SheetWriter<T, SELF>`, which is public so that a setting is written once. For
example a standard look for the reports of a team:

```java
static <T, W extends SheetWriter<T, W>> W standardLook(W writer) {
    return writer.stripedRows().autoFilter(true).freezePane(0, 1).emptyMessage("No data");
}

standardLook(ExcelWriter.<Order>create()).column("Id", Order::id).write(orders);

try (var workbook = ExcelWorkbook.create()) {
    standardLook(workbook.<Order>sheet("Orders")).column("Id", Order::id).write(orders);
    standardLook(workbook.<Order>sheet("Returns")).column("Id", Order::id).write(returns);
}
```

`SheetWriter` cannot be extended outside the library.

The settings of the workbook (`documentProperty`, `headerColor`, `headerFontName`, `headerFontSize`, `headerStyle`,
`protectWorkbook`, `password`) are written once in the same way: `ExcelWriter` and `ExcelWorkbook` both implement
`WorkbookOptions<SELF>`.

```java
static <W extends WorkbookOptions<W>> W teamLook(W writer) {
    return writer.headerColor(ExcelColor.STEEL_BLUE).headerFontName("Arial").documentProperty("author", "Reporting");
}

teamLook(ExcelWriter.<Order>create()).column("Id", Order::id);
teamLook(ExcelWorkbook.create()).sheet("Orders");
```

What can only be decided when the workbook is created (`rowAccessWindowSize`, `compressTempFiles`, `useSharedStrings`,
`streaming`) is a `WorkbookInitOptions`, the same for `ExcelWriter.create(opts -> ...)`, `ExcelWriter.forMap(opts -> ..., names)` and
`ExcelWorkbook.create(opts -> ...)`.

The columns (`column`, `columnIf`, `constColumn`, `constColumnIf`, `rowNumberColumn`) are one level higher, in
`ColumnsWriter<T, SELF>`, which `SheetWriter` extends. `TemplateListWriter` (a list written into a sheet of a template that
already exists, from `ExcelTemplateWriter.list(...)`) extends it as well, so it has the same columns and takes the same
`Consumer<ColumnConfig<T>>`; it has no sheet settings (frozen panes, filter, rollover...), the template has those already.
It also takes `writeErrorPolicy(...)`.

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

## Closing a row stream

`write(Stream)` reads the stream but does not close it; the code that opened it does. For a JPA or
JDBC stream, wrap the call in try-with-resources so the cursor and connection are released:

```java
try (Stream<Order> orders = orderRepository.streamAll()) {
    return ExcelWriter.<Order>create()
        .column("Id", Order::id)
        .write(orders);
}
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

The `nullValue` is always written, whatever the error policy: on a typed numeric or date column a
text placeholder such as `"-"` is written as text instead of failing the export.

## Auto Width

Column widths are auto-calculated from the first N data rows:

```java
writer.autoWidthSampleRows(200)  // sample the first 200 data rows (default: 100)
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

When the data rolls over to more sheets, each sheet covers its own rows: the first keeps `PriceData`, the
next ones are `PriceData_2`, `PriceData_3`, ... (a workbook cannot hold two names alike).

For manual control, use `SheetContext.namedRange()` in `afterData` callbacks.
## Streaming initialization

SXSSF memory and disk behavior can be selected when the writer is created:

```java
ExcelWriter.<Row>create(options -> options
        .rowAccessWindowSize(500)
        .compressTempFiles(true)
        .useSharedStrings(false));
```

Use `.table("RowsTable")` to create an Excel structured table over the generated header and
data range.
For shared configuration, pass `new StreamingOptions(windowSize, compressTempFiles,
useSharedStrings)` through the initialization options. `TableOptions` controls style, row
stripes, and whether rollover sheets receive uniquely suffixed tables.
Structured tables are also available on `ExcelWorkbook.sheet(...)` and template list writers.
Table names are checked workbook-wide without regard to case, and template tables validate
that the expected header row contains every configured column.

Structured-table column names are retained even when streaming has flushed the
header row out of memory. For template lists using existing headers, table column
names come from the template cells rather than the mapping labels.
