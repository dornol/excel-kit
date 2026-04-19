# excel-kit

Fluent API-based Excel/CSV read-write Java library built on Apache POI (streaming SXSSF).
Not annotation-based — columns are defined programmatically via builder chains or lambda configs.

- Group: `io.github.dornol`
- Artifact: `excel-kit`
- Requires: Java 17+, Apache POI (compileOnly)
- GitHub: https://github.com/dornol/excel-kit
- Javadoc: https://dornol.github.io/excel-kit/

## Quick Reference

| Task | Class | Pattern |
|------|-------|---------|
| Write Excel (typed) | `ExcelWriter<T>` | `.column("Name", T::getName).write(stream).writeTo(path)` |
| Write Excel (map) | `ExcelWriter.forMap(...)` | `ExcelWriter.forMap("Name", "Age").write(stream).writeTo(path)` |
| Write Excel (multi-sheet) | `ExcelWorkbook` | `wb.<T>sheet("Sheet1").column(...).write(stream)` → `wb.finish().writeTo(path)` |
| Write Excel (template) | `ExcelTemplateWriter` | `new ExcelTemplateWriter(template).list(...).write(stream, out)` |
| Read Excel (setter) | `ExcelReader<T>` | `ExcelReader.setter(T::new).column("Name", T::setName).required().build(in).read(r -> ...)` |
| Read Excel (map) | `ExcelReader.forMap()` | `ExcelReader.forMap().build(in).read(r -> r.data().get("Name"))` |
| Read Excel (mapping) | `ExcelReader.mapping()` | `ExcelReader.mapping(row -> new Record(row.get("Name").asString())).build(in).read(r -> ...)` |
| Write CSV | `CsvWriter<T>` | `.column("Name", T::getName).write(stream).writeTo(path)` |
| Write CSV (map) | `CsvWriter.forMap(...)` | `CsvWriter.forMap("Name", "Age").write(stream).writeTo(path)` |
| Read CSV (setter) | `CsvReader<T>` | `CsvReader.setter(T::new).column("Name", T::setName).build(in).read(r -> ...)` |
| Read CSV (map) | `CsvReader.forMap()` | `CsvReader.forMap().build(in).read(r -> r.data().get("Name"))` |

## Detailed Documentation

The following files are located in `META-INF/excel-kit/` within this JAR:

- **[quick-start.md](excel-kit/quick-start.md)** — Basic write/read examples for Excel and CSV
- **[column-config.md](excel-kit/column-config.md)** — Column styling, data types, header font color, default styles
- **[reading.md](excel-kit/reading.md)** — Excel/CSV reading (name-based, index-based, mapping mode, map reader)
- **[advanced.md](excel-kit/advanced.md)** — Multi-sheet, protection, charts, conditional formatting, images, validation
- **[csv.md](excel-kit/csv.md)** — CSV-specific features (dialect, delimiter, BOM, quoting, injection defense)

## Two Writer APIs

`ExcelWriter<T>` (single-type, auto-rollover):
```java
ExcelWriter.<Person>create()
    .column("Name", Person::name)
    .column("Age", Person::age, cfg -> cfg.type(ExcelDataType.INTEGER))
    .write(stream)
    .writeTo(out);
```

`ExcelWorkbook` (multi-sheet, different types per sheet):
```java
try (var wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users").column("Name", User::getName).write(userStream);
    wb.<Order>sheet("Orders").column("ID", Order::getId).write(orderStream);
    wb.finish().writeTo(out);
}
```

All writer APIs (`ExcelWriter`, `ExcelSheetWriter`, `CsvWriter`) use the same `.column("Name", fn, cfg -> cfg.type().bold())` pattern.

## Key API Notes (v0.16.0+)

- `FileHandler.writeTo()` throws unchecked exceptions only — no `throws IOException`
- `nullValue(Object)` on column config — default value for null cells (e.g., `c -> c.nullValue("N/A")`)
- Freeze panes:
  - `freezeRows(int)` — freeze N rows below the header
  - `freezeCols(int)` — freeze N columns from the left
  - `freezePane(int cols, int rows)` — freeze both
- Encrypted output:
  - `.password("pw")` on writer → `handler.writeTo(out)` auto-encrypts
  - `handler.writeTo(out, "pw")` / `handler.writeTo(path, "pw")` — late-binding encryption
  - `char[]` overloads zero the array after use
- Writers use static factory `create()`: `ExcelWriter.<T>create()`, `ExcelWorkbook.create()`,
  `CsvWriter.<T>create()` (no public constructors)
- `required()` on reader columns — blank cells produce validation errors
- `asBigDecimal()` parses directly without Double intermediate — full precision
- `ExcelSheetWriter.write()` is single-call — second call throws `ExcelWriteException`
- `ExcelImage.png()/jpeg()` validates non-null, non-empty byte array at creation time
- CSV injection defense covers leading whitespace + formula characters (e.g., `" =cmd"`)

## Key API Notes (v0.16.9+)

- **Multi-level group headers** — `group(String... levels)` takes N levels, outermost first.
  `.column("Q1", Row::q1, c -> c.group("Financial", "Revenue", "2025"))` produces 3 group rows + 1 header row.
  Source-compatible with single-level `.group("X")`. Reflective readers: field renamed `groupName` → `groupNames`,
  getter `getGroupNames()` replaces `getGroupName()`.
- **Header customization** (writer-level):
  - `.headerRowHeight(float points)` — applies to every header row including group rows; `0` = default
  - `.headerFontName(String)` / `.headerFontSize(int)`
  - `.rowNumberColumn(String name)` — 1-based sequential column; shorthand for
    `column(name, (r, cur) -> cur.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))`, rolls over with `maxRows()`
- **Per-column header background** — `headerBackgroundColor(ExcelColor)` or `(int r, int g, int b)` on column config.
  Overrides workbook-wide `headerColor` for one column only.
- **Group header comments** — `writer.groupComment(String text, String... path)` or
  `groupComment(ExcelCellComment, String... path)`. Path is outermost-first and must match a declared
  `group(...)`; no-op otherwise.

## Key API Notes (v0.16.12+) — Reading

- **Split success/error callbacks** — `read(Consumer<T> onSuccess, Consumer<RowError> onError)` routes
  validated rows vs failed rows. Library buffers nothing — caller decides error memory policy.
- **`RowError`** record — `rowNum` (1-based, header excluded), `type` (`VALIDATION` / `MAPPING`),
  `messages`, nullable `cause`.
- **`ReadResult<T>.cause()`** — nullable throwable from mapping stage. 3-arg constructor retained for
  backward compatibility.

## Key API Notes (v0.16.13+) — Reading

- **Multi-row header** — `ExcelReader.headerRows(int)` combines N header rows into effective column
  names, taking the bottom-most non-blank per column. Use with files written via multi-level `group(...)`:
  ```java
  reader.headerRowIndex(1).headerRows(2).build(in).read(r -> ...);
  ```
  Default `headerRows(1)` preserves existing single-row behavior including empty-string headers.

## Key API Notes (v0.16.14+)

- **Document properties** — `documentProperty(key, value)` on `ExcelWriter` / `ExcelWorkbook`.
  Standard keys (`title`, `subject`, `author`/`creator`, `keywords`, `description`, `category`)
  map to core properties; others become custom properties.
- **Fluent named ranges** — `namedRange(name, columnIndex)` registers a workbook-scoped named
  range covering all data rows in the given column. Replaces manual `afterData` callback usage.
- **Header style config** — `headerStyle(cfg -> cfg.alignment(...).border(...).bold(...).wrapText(...))`
  on `ExcelWriter` / `ExcelWorkbook`. Overrides default header alignment/bold/border/wrap.
- **`password(char[])`** — char-array overload on `ExcelWriter` / `ExcelWorkbook`. Array is copied
  internally and zeroed after encryption.
