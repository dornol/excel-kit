# Multi-Sheet Workbook

> [Back to Index](index.md)

## Multi-Sheet Reading

하나의 workbook에서 sheet별 reader/schema를 등록해 서로 다른 DTO를 함께 읽을 수 있습니다.

```java
WorkbookReadReport report = ExcelWorkbookReader.create()
    .sheet("Users", ExcelReader.mapping(row -> new User(row.get("Name").asString())),
        result -> { if (result.success()) importUser(result.data()); })
    .sheet("Orders", orderSchema, Order::new, validator,
        result -> { if (result.success()) importOrder(result.data()); })
    .failurePolicy(SheetFailurePolicy.CONTINUE)
    .read(inputStream);
```

Sheet 이름 또는 0-based index를 사용할 수 있으며, `Path`와 `InputStreamSource`도 지원합니다.
caller가 전달한 `InputStream`은 닫지 않고, source가 연 stream은 reader가 닫습니다.

기본 `CONTINUE` 정책은 구조적 sheet 오류가 발생해도 다음 등록 sheet를 계속 처리합니다.
`SheetFailurePolicy.STOP`을 사용하면 첫 구조적 오류에서 중단합니다. 행 단위 mapping/validation
오류는 `SheetReadReport.errors()`에 bounded list로 집계되며, 각 sheet의 consumer에는 기존
`ReadResult`가 그대로 전달됩니다.

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
