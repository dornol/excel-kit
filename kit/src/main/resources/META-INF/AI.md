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
| Write Excel (typed) | `ExcelWriter<T>` | `.column("Name", T::getName).write(stream).toFile(path)` |
| Write Excel (map) | `ExcelWriter.forMap(...)` | `ExcelWriter.forMap("Name", "Age").write(stream).toFile(path)` |
| Write Excel (multi-sheet) | `ExcelWorkbook` | `wb.<T>sheet("Sheet1").column(...).write(stream)` → `wb.finish().toFile(path)` |
| Write Excel (template) | `ExcelTemplateWriter` | `new ExcelTemplateWriter(template).list(...).write(stream, out)` |
| Read Excel (setter) | `ExcelReader<T>` | `ExcelReader.setter(T::new).column("Name", T::setName).required().build(in).read(r -> ...)` |
| Read Excel (map) | `ExcelReader.forMap()` | `ExcelReader.forMap().build(in).read(r -> r.data().get("Name"))` |
| Read Excel (mapping) | `ExcelReader.mapping()` | `ExcelReader.mapping(row -> new Record(row.get("Name").asString())).build(in).read(r -> ...)` |
| Write CSV | `CsvWriter<T>` | `.column("Name", T::getName).write(stream).toFile(path)` |
| Write CSV (map) | `CsvWriter.forMap(...)` | `CsvWriter.forMap("Name", "Age").write(stream).toFile(path)` |
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
ExcelWriter.<Person>builder().build()
    .column("Name", Person::name)
    .column("Age", Person::age, cfg -> cfg.type(ExcelDataType.INTEGER))
    .write(stream)
    .write(out);
```

`ExcelWorkbook` (multi-sheet, different types per sheet):
```java
try (var wb = ExcelWorkbook.builder().color(ExcelColor.STEEL_BLUE).build()) {
    wb.<User>sheet("Users").column("Name", User::getName).write(userStream);
    wb.<Order>sheet("Orders").column("ID", Order::getId).write(orderStream);
    wb.finish().write(out);
}
```

All writer APIs (`ExcelWriter`, `ExcelSheetWriter`, `CsvWriter`) use the same `.column("Name", fn, cfg -> cfg.type().bold())` pattern.

## Key API Notes (v0.16.0+)

- `FileHandler.write()` throws unchecked exceptions only — no `throws IOException`
- `nullValue(Object)` on column config — default value for null cells (e.g., `c -> c.nullValue("N/A")`)
- `freezePane(int cols, int rows)` — freeze both columns and rows
- `required()` on reader columns — blank cells produce validation errors
- `asBigDecimal()` parses directly without Double intermediate — full precision
- `ExcelSheetWriter.write()` is single-call — second call throws `ExcelWriteException`
- `ExcelImage.png()/jpeg()` validates non-null, non-empty byte array at creation time
- CSV injection defense covers leading whitespace + formula characters (e.g., `" =cmd"`)
