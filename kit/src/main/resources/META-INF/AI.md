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
| Write Excel (typed) | `ExcelWriter<T>` | `.column("Name", T::getName).write(stream).consumeOutputStream(out)` |
| Write Excel (map) | `ExcelMapWriter` | `new ExcelMapWriter("Name", "Age").write(stream)` |
| Write Excel (multi-sheet) | `ExcelWorkbook` | `wb.<T>sheet("Sheet1").column(...).write(stream)` |
| Write Excel (template) | `ExcelTemplateWriter` | `new ExcelTemplateWriter(template).list(...).write(stream, out)` |
| Read Excel (typed) | `ExcelReader<T>` | `.column("Name", T::setName).build(in).read(r -> ...)` |
| Read Excel (map) | `ExcelMapReader` | `new ExcelMapReader().build(in).read(r -> r.data().get("Name"))` |
| Read Excel (mapping) | `ExcelReader.mapping()` | `ExcelReader.mapping(row -> new Record(row.get("Name").asString()))` |
| Write CSV | `CsvWriter<T>` | `.column("Name", T::getName).write(stream).consumeOutputStream(out)` |
| Write CSV (map) | `CsvMapWriter` | `new CsvMapWriter("Name", "Age").write(stream)` |
| Read CSV | `CsvReader<T>` | `.column("Name", T::setName).build(in).read(r -> ...)` |
| Read CSV (map) | `CsvMapReader` | `new CsvMapReader().build(in).read(r -> r.data().get("Name"))` |

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
new ExcelWriter<Person>()
    .column("Name", Person::name)
    .column("Age", Person::age, cfg -> cfg.type(ExcelDataType.INTEGER))
    .write(stream)
    .consumeOutputStream(out);
```

`ExcelWorkbook` (multi-sheet, different types per sheet):
```java
try (var wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users").column("Name", User::getName).write(userStream);
    wb.<Order>sheet("Orders").column("ID", Order::getId).write(orderStream);
    wb.finish().consumeOutputStream(out);
}
```

All writer APIs (`ExcelWriter`, `ExcelSheetWriter`, `CsvWriter`) use the same `.column("Name", fn, cfg -> cfg.type().bold())` pattern.
