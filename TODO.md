# excel-kit — Improvement Backlog

Analyzed 2026-04-19. Candidates for future work.

---

## Refactoring

### R1. ExcelColumn constructor — 22 parameters (HIGH)

Every field addition breaks the constructor + `ExcelColumn.of()` + all test call sites.
Extract parameter groups:

```
HeaderConfig    — headerFontColor, headerBackgroundColor, headerComment, commentWidth, commentHeight
WidthConfig     — minWidth, maxWidth, fixedWidth
ValidationConfig — dropdownOptions, validation
```

Reduces constructor from 22 params to ~8-10.

### R2. ExcelWriter / ExcelSheetWriter — 65% method duplication (HIGH)

`freezeRows`, `rowColor`, `onProgress`, `chart`, `summary`, `conditionalFormatting`,
`printSetup`, `tabColor`, `defaultStyle`, `protectSheet`, etc. are near-identical copies.

Option: extract `AbstractSheetWriter<T>` base class.

### R3. ExcelReader / CsvReader — 60% structural duplication (HIGH)

`setter()`, `mapping()`, `forMap()` factories + column registration + validation
logic duplicated across both readers.

Option: extract `AbstractFileReader<T>` base with shared column management.

### R4. writeGroupAndColumnHeaders() — 104 lines (MEDIUM)

Split into: `createHeaderRows()`, `populateGrid()`, `applyHorizontalMerges()`, `applyVerticalMerges()`.

### R5. ColumnStyleConfig — 59 fields (MEDIUM)

Intentionally flat for fluent API, but consider composition (FontConfig, BorderConfig, LayoutConfig)
if field count keeps growing.

### R6. Test coverage gaps (LOW)

No direct unit tests for: AbstractReadHandler, ExcelValidation, CellData,
ColumnConfig, ExcelChartConfig, ExcelPrintSetup, ExcelBorderStyle, ExcelHyperlink, ExcelImage.
Covered indirectly through integration tests but isolated tests would catch edge cases.

---

## Security

### S1. password(String) stores password in immutable heap (HIGH)

`ExcelWriter.password(String)` and `ExcelReader.password(String)` store as `String`.
`writeTo(out, char[])` already zeroes the array — but builder-level API doesn't.

Fix: add `password(char[])` overloads, store internally as `char[]`, zero after use.

### S2. ZIP bomb protection not enabled by default (MEDIUM)

`ExcelKitConfig.configureLargeFileSupport()` must be called manually.
Default POI limits may be too permissive for untrusted input.

Options:
- Document prominently in guide/reference.md
- Auto-configure reasonable defaults on first reader use

### S3. Decrypted temp file on disk in plaintext (MEDIUM)

`ExcelReadHandler.decryptFile()` writes decrypted content to temp file.
Deleted after use, but not securely overwritten.

Options:
- Document that disk encryption (LUKS/BitLocker) is recommended
- Stream decrypted data directly to parser if POI allows

### S4. Windows temp ACL restriction may fail silently (LOW)

`TempResourceCreator.restrictToOwnerOnWindows()` catches IOException and logs warning.
On some filesystems, temp dir may be world-readable.

### S5. XXE — delegated to POI (LOW)

POI 5.5.1 disables external entities by default. Safe as long as POI >= 5.0.
Consider explicit SAX hardening for defense-in-depth.

---

## New API Candidates

### A1. documentProperty(key, value) — Excel metadata (MEDIUM, Simple)

```java
ExcelWorkbook.create()
    .documentProperty("title", "Sales Report Q4")
    .documentProperty("author", "Finance Team")
```

POI supports this natively. No current API.

### A2. Fluent namedRange on writer (MEDIUM, Simple)

```java
writer
    .column("Price", Product::price, c -> c.type(ExcelDataType.DOUBLE))
    .namedRange("PriceData", 0)  // column index 0
    .write(data);
```

Currently requires manual `afterData` callback with raw string reference.

### A3. headerStyle(cfg -> ...) — header-only defaults (LOW, Simple)

```java
writer.headerStyle(h -> h
    .fontSize(14).bold(true)
    .fontColor(ExcelColor.WHITE)
    .backgroundColor(ExcelColor.DARK_BLUE))
```

`defaultStyle()` is data-cell only. No centralized header styling beyond color/font name/size.

### A4. skipRowsWhere(predicate) — read-side row filtering (LOW, Simple)

```java
ExcelReader.setter(User::new)
    .skipRowsWhere(row -> row.get("Name").isEmpty())
    .column("Name", (u, c) -> u.setName(c.asString()))
    .build(in).read(consumer);
```

Currently must filter in read callback manually.
