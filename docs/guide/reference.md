# Reference

> [Back to Index](index.md)

## ExcelDataType Reference

| Type | Java Type | Default Format |
|------|-----------|---------------|
| `STRING` | String | — |
| `BOOLEAN_TO_YN` | Boolean -> "Y"/"N" | — |
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

## ExcelDataFormat Presets

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

## ExcelKitSchema — Unified Read/Write

Define columns once for both reading and writing:

```java
ExcelKitSchema<Book> schema = ExcelKitSchema.<Book>builder()
    .column("Title", Book::getTitle, (b, cell) -> b.setTitle(cell.asString()))
    .column("Price", Book::getPrice, (b, cell) -> b.setPrice(cell.asInt()),
            c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
    .build();

// Write
schema.excelWriter().write(bookStream);
schema.csvWriter().write(bookStream);

// Read (setter mode)
schema.excelReader(Book::new, null).build(inputStream);
schema.csvReader(Book::new, null).build(inputStream);

// Read (mapping mode)
schema.excelReader(row -> new BookRecord(
    row.get("Title").asString(), row.get("Price").asInt()
), null).build(inputStream);
```

## ExcelColor Presets

`WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `DARK_GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `LIGHT_ORANGE`, `LIGHT_PURPLE`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`

## Exception Handling

| Exception | Description |
|-----------|-------------|
| `ExcelKitException` | Base class for all library exceptions |
| `ExcelWriteException` | Excel write errors (no columns, handler already consumed, etc.) |
| `ExcelReadException` | Excel read/parse errors |
| `CsvWriteException` | CSV write errors |
| `CsvReadException` | CSV read/parse errors |

- Column mapping exceptions fall back to string conversion (Excel writing).
- Calling `write` on an already-consumed handler throws the corresponding `WriteException`.
- Empty password throws `IllegalArgumentException`.

## Requirements

- **JDK 17+**
- Apache POI 5.x for Excel operations

## Supported Formats

| Format | Read | Write | Notes |
|--------|------|-------|-------|
| `.xlsx` (Excel 2007+) | Yes | Yes | Streaming read (SAX) and write (SXSSF) |
| `.csv` | Yes | Yes | Via OpenCSV, configurable delimiter/charset |
| `.xls` (Excel 97-2003) | No | No | Legacy binary format not supported |
| `.xlsm` (Macro-enabled) | No | No | Macros cannot be generated or preserved |
| `.ods` (OpenDocument) | No | No | Not supported |

## Notes

### Large file configuration is JVM-global

`ExcelReader.configureLargeFileSupport()` adjusts Apache POI's internal limits.
These are JVM-global static settings — call once at application startup.

### `readAsStream()` requires try-with-resources

```java
try (Stream<ReadResult<T>> stream = handler.readAsStream()) {
    stream.filter(ReadResult::success)
          .map(ReadResult::data)
          .forEach(this::process);
}
```
