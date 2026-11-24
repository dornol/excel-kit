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

`ExcelKitSchema` is in `io.github.dornol.excelkit.schema`. Define columns once for both reading and writing:

```java
ExcelKitSchema<Book> schema = ExcelKitSchema.<Book>builder()
    .requiredColumn("Title", List.of("Book Title", "제목"),
            Book::getTitle, (b, cell) -> b.setTitle(cell.asString()))
    .column("Price", Book::getPrice, (b, cell) -> b.setPrice(cell.asInt()),
            c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
    .build();

// Write
schema.excelWriter().write(bookStream);
schema.csvWriter().write(bookStream);

// Read (setter mode)
schema.excelReader(Book::new, null).read(inputStream, result -> { ... });
schema.csvReader(Book::new, null).read(inputStream, result -> { ... });

// Read (mapping mode)
schema.excelReader(row -> new BookRecord(
    row.get("Title").asString(), row.get("Price").asInt()
), null).read(inputStream, result -> { ... });
```

## ExcelColor Presets

`WHITE`, `BLACK`, `LIGHT_GRAY`, `GRAY`, `DARK_GRAY`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `ORANGE`, `LIGHT_RED`, `LIGHT_GREEN`, `LIGHT_BLUE`, `LIGHT_YELLOW`, `LIGHT_ORANGE`, `LIGHT_PURPLE`, `CORAL`, `STEEL_BLUE`, `FOREST_GREEN`, `GOLD`

## Packages and Imports

Every public type and the package it lives in. Each package depends only on the ones to its left: `core` ← `core.io` ←
`core.read`/`core.detect` ← `csv`, `excel.read`, `excel` (with `excel.style` and `excel.feature`) ← `schema`.

| Package | Public types | Use for |
|---------|--------------|---------|
| `io.github.dornol.excelkit.core` | `AtomicFileWriter`, `Cursor`, `ExcelKitException`, `FileHandler`, `ProgressCallback`, `RuntimeDependencyGuard` | Shared contracts: `FileHandler` (what a writer returns), progress callbacks, the base exception |
| `io.github.dornol.excelkit.core.read` | `AbstractReadHandler`, `AbstractReader`, `CancellationToken`, `CellConversionConfig`, `CellData`, `CellError`, `DuplicateHeaderPolicy`, `HeaderPolicy`, `ReadAbortException`, `ReadAbortReason`, `ReadColumn`, `ReadLimitExceededException`, `ReadLimits`, `ReadOptions`, `ReadProgress`, `ReadProgressCallback`, `ReadReport`, `ReadResult`, `ReadSecurityException`, `ReadSecurityPolicy`, `ReadSession`, `ReadStoppedException`, `ReadSummary`, `ReaderOptions`, `RowData`, `RowError`, `RowFunction` | Reading, shared by Excel and CSV: options, limits, security policy, cell/row access, results and errors |
| `io.github.dornol.excelkit.core.io` | `InputStreamSource`, `TempResourceCleaner`, `TempResourceContainer`, `TempResourceCreateException`, `TempResourceCreator`, `TempWorkspace` | Input sources and temp-file management |
| `io.github.dornol.excelkit.core.detect` | `DetectionConfidence`, `TabularDetectionResult`, `TabularFileDetector`, `TabularFileType` | Telling an .xlsx from a CSV by content |
| `io.github.dornol.excelkit.csv` | `CsvAfterDataWriter`, `CsvColumn`, `CsvColumnConfig`, `CsvDialect`, `CsvHandler`, `CsvQuoting`, `CsvReadException`, `CsvReader`, `CsvWriteException`, `CsvWriter` | CSV reading and writing |
| `io.github.dornol.excelkit.excel` | `AfterDataWriter`, `BeforeHeaderWriter`, `ColumnConfig`, `ColumnStyleConfig`, `ColumnsWriter`, `ExcelChartConfig`, `ExcelColumn`, `ExcelConditionalRule`, `ExcelDataType`, `ExcelErrorReport`, `ExcelHandler`, `ExcelKitConfig`, `ExcelPrintSetup`, `ExcelRichText`, `ExcelSheetWriter`, `ExcelSubtotal`, `ExcelSummary`, `ExcelTemplateWriter`, `ExcelValidation`, `ExcelWorkbook`, `ExcelWriteErrorPolicy`, `ExcelWriteException`, `ExcelWriter`, `HeaderStyleConfig`, `RowStyleConfig`, `SheetContext`, `SheetWriter`, `StreamingOptions`, `TableOptions`, `TemplateListWriter`, `WriteRowCallback` | Excel writing: `ExcelWriter`, `ExcelWorkbook`, templates, column and sheet configuration |
| `io.github.dornol.excelkit.excel.read` | `ExcelReadException`, `ExcelReader`, `ExcelSheetInfo`, `ExcelWorkbookInspector`, `ExcelWorkbookReader`, `SheetFailurePolicy`, `SheetReadReport`, `WorkbookReadReport` | Excel reading: `ExcelReader`, multi-sheet `ExcelWorkbookReader`, sheet metadata |
| `io.github.dornol.excelkit.excel.style` | `CellColorFunction`, `ExcelBorderStyle`, `ExcelColor`, `ExcelDataFormat` | Colors, borders and number formats |
| `io.github.dornol.excelkit.excel.feature` | `ExcelCellComment`, `ExcelHyperlink`, `ExcelImage` | Cell content beyond text: comments, hyperlinks, images |
| `io.github.dornol.excelkit.schema` | `ExcelKitSchema` | One column definition for both reading and writing, Excel or CSV |
| `io.github.dornol.excelkit.spring` | `CellErrorResponse`, `DownloadFileType`, `ErrorReportRow`, `ExcelKitErrorResponse`, `ExcelKitMultipartFile`, `ExcelKitResponse`, `ExcelKitTemplateResponse`, `ExcelKitUpload`, `ExcelKitUploadException`, `UploadCollectionLimits`, `UploadError`, `UploadReader`, `UploadResult`, `UploadSummary` | Spring MVC download and upload helpers (`excel-kit-spring` artifact) |

Typical imports:

```java
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
```

## Exception Handling

| Exception | Extends | Description |
|-----------|---------|-------------|
| `ExcelKitException` | `RuntimeException` | Base class for the library's own exceptions |
| `ExcelWriteException` | `ExcelKitException` | Excel write errors (no columns, handler already consumed, etc.) |
| `ExcelReadException` | `ExcelKitException` | Excel read/parse errors |
| `CsvWriteException` | `ExcelKitException` | CSV write errors |
| `CsvReadException` | `ExcelKitException` | CSV read/parse errors |
| `ReadLimitExceededException` | `ExcelKitException` | A `ReadLimits` bound was exceeded; `limit()`, `configured()` and `actual()` say which |
| `ReadSecurityException` | `ExcelKitException` | `ReadSecurityPolicy` rejected the input; `reason()` names the rule |
| `TempResourceCreateException` | `ExcelKitException` | A temporary file or directory could not be created |
| `ReadAbortException` | `RuntimeException` | `readStrict()`/`readAll()` hit a failing row, or `maxErrors()` was exceeded; `reason()` tells which |
| `ReadStoppedException` | `RuntimeException` | Internal signal for `readWhile()`; the reader catches it, you do not see it |

`ReadAbortException` does **not** extend `ExcelKitException`, so `catch (ExcelKitException e)` does not catch a
`readStrict()` or `maxErrors()` abort. Catch it explicitly:

```java
try {
    reader.readStrict(inputStream, result -> { ... });
} catch (ReadAbortException e) {
    // e.reason() is STRICT_FAILURE or MAX_ERRORS_EXCEEDED
} catch (ExcelKitException e) {
    // read errors, limits, security policy
}
```

- Column mapping exceptions fall back to string conversion (Excel writing).
- Calling `write` on an already-consumed handler throws the corresponding `WriteException`.
- Empty password throws `IllegalArgumentException`.

## Requirements

- **JDK 17+**
- Apache POI 5.x for Excel operations

The core artifact keeps format integrations as `compileOnly`. Add
`org.apache.poi:poi-ooxml` for Excel and `com.opencsv:opencsv` for CSV. If one is
missing at runtime, the corresponding reader reports the required artifact and feature
with an actionable error message.

For untrusted input, also set `limits(ReadLimits.UPLOAD)` (readers have no limits by default).
For untrusted Excel input, use `ReadSecurityPolicy.STRICT`. It rejects formulas and
external workbook links and bounds scanned entry size, total scan bytes, compression ratio,
and ZIP entry count. `ReadSecurityException.reason()` identifies the violated rule.

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

`ExcelReader.configureLargeFileSupport()` adjusts Apache POI's internal limits
(`ZipSecureFile.maxFileCount`, `IOUtils.byteArrayMaxOverride`).
These are JVM-global static settings — call once at application startup.

> **Security:** When reading untrusted files, always call this method to enforce
> safe decompression limits. See [Protection — Reading Untrusted Files](protection.md#reading-untrusted-files).

### Reader resource ownership

```java
try (InputStream input = Files.newInputStream(path)) {
    reader.read(input, this::process);
}
```

The caller closes a supplied `InputStream`. For `reader.read(path, ...)` or an
`InputStreamSource`, the reader opens and closes the stream.
