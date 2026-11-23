# Validation & Conditional Formatting

> [Back to Index](index.md)

## Dropdown Validation

```java
.column("Status", p -> p.status(), cfg -> cfg
    .dropdown("Active", "Inactive", "Pending"))
```

Applied to all data rows across all sheets (including rollover).

Excel limits the comma-joined option text to 255 characters; `dropdown(...)` throws
`IllegalArgumentException` beyond that. Use `listFromRange(range)` for longer lists.

## Dependent Dropdown

A cell whose options depend on another column of the same row: the parent column holds `"A"`, the cell offers
the options of `"A"`; the parent holds `"B"`, it offers those of `"B"`.

```java
.column("Category", p -> p.category(), cfg -> cfg.dropdown("A", "B"))
.column("Item", p -> p.item(), cfg -> cfg.dependentDropdown("Category", Map.of(
        "A", List.of("1", "2", "3"),
        "B", List.of("4", "5", "6"))))   // use a LinkedHashMap to keep the parent order
```

- `parentColumn` is the name (header) of another column of the same sheet; an unknown name fails when the sheet is written.
- The options go to a hidden sheet (`_excelkit_dep_1`, ...). Equal option sets share one sheet, also across rollover sheets.
- The parent value is matched as text, case-insensitively. A row whose parent value is not a key (or is empty) offers no options.
- Excel checks a value when it is typed, so a child left over after the parent changes stays: validate the combination when reading the upload.
- A column has one data validation: combining it with `dropdown(...)` or `validation(...)` throws `IllegalArgumentException`.

## Advanced Data Validation

```java
.column("Age", p -> p.age(), cfg -> cfg
    .validation(ExcelValidation.integerBetween(0, 150)))
.column("GPA", p -> p.gpa(), cfg -> cfg
    .validation(ExcelValidation.decimalBetween(0.0, 4.0)))
.column("Name", p -> p.name(), cfg -> cfg
    .validation(ExcelValidation.textLength(1, 100)))
.column("Date", p -> p.date(), cfg -> cfg
    .validation(ExcelValidation.dateRange(LocalDate.of(2024, 1, 1), LocalDate.of(2024, 12, 31))))
.column("Custom", p -> p.value(), cfg -> cfg
    .validation(ExcelValidation.formula("AND(A2>0,A2<100)")))
```

**Factory methods:**
- `integerBetween(min, max)` / `integerGreaterThan(min)` / `integerLessThan(max)`
- `decimalBetween(min, max)`
- `textLength(min, max)`
- `dateRange(start, end)`
- `formula(formula)` — custom Excel formula
- `listFromRange(range)` — dropdown from cell range (e.g., `"Sheet2!$A$1:$A$10"`)

**Error messages:**
```java
ExcelValidation.integerBetween(1, 100)
    .errorTitle("Invalid Value")
    .errorMessage("Please enter a number between 1 and 100")
```

## List Validation from Cell Range

```java
.column("Status", p -> p.status(), cfg -> cfg
    .validation(ExcelValidation.listFromRange("Options!$A$1:$A$5")))
```

Useful for large option lists on a separate sheet.

## Conditional Formatting

```java
writer
    .column("Name", Product::name)
    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.INTEGER))
    .conditionalFormatting(cf -> cf
        .columns(1)                                     // apply to column 1 only
        .greaterThan("10000", ExcelColor.LIGHT_RED)
        .lessThan("1000", ExcelColor.LIGHT_GREEN)
        .between("5000", "10000", ExcelColor.LIGHT_YELLOW))
    .write(data);
```

**Operators:** `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equalTo`, `notEqualTo`, `between`, `notBetween`

If `columns()` is not set, rules apply to all columns.

Also available: `.dataBar()` for gradient bars, `.iconSet()` for arrows/traffic lights.

## Required Columns (Reading)

See [Reading — Required Columns](reading.md#required-columns).
