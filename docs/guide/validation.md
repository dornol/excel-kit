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

### Input message and required cells

```java
.column("Age", p -> p.age(), cfg -> cfg.validation(ExcelValidation.integerBetween(0, 150)
    .inputTitle("Age")                       // tooltip shown when the cell is selected (title <= 32 chars)
    .inputMessage("Enter a whole number 0-150") // (text <= 255 chars)
    .allowBlank(false)))                      // an empty cell no longer passes (default: allowed)
```

Excel rejects *typing* an empty value into an `allowBlank(false)` cell but does not flag a cell that was never filled,
so check required values when reading the upload.

A plain `dropdown(...)` has neither option; use `ExcelValidation.list(...)` for the same fixed list with them, and
`ExcelValidation.dependentList(parent, options)` for a dependent one:

```java
.column("Status", p -> p.status(), cfg -> cfg.validation(
    ExcelValidation.list("Active", "Inactive").allowBlank(false).inputMessage("Pick one")))
.column("Item", p -> p.item(), cfg -> cfg.validation(
    ExcelValidation.dependentList("Category", options).allowBlank(false)))
```

`list(...)` follows the rules of `dropdown(...)` (no comma in an option, 255 characters in all).

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
- An option that is a plain number (`"1"`, `"-2"`, `"3.5"`; no leading zero, no trailing zero after the point, at most 15 digits) is written to that
  sheet as a **number**, the others as text. Excel compares a list entry with the cell by type: a list of the text `"1"` refuses the
  number 1 that a person types or picks into a General cell, a list of the number 1 accepts it (checked on Excel). A fixed list
  (`dropdown(...)`, `ExcelValidation.list(...)`) accepts both. A cell your own writer filled with the text `"1"` is not
  checked until it is edited, and typing the number then matches.
- The parent value is matched as text, case-insensitively; a numeric parent cell (a code column) matches the key written as text (`Map.of("1", ...)`). A row whose parent value is not a key (or is empty) offers no options.
- Excel checks a value when it is typed, so a child left over after the parent changes stays: validate the combination when reading the upload.
- A column has one data validation: combining it with `dropdown(...)` throws `IllegalArgumentException`, and a `validation(...)` set after it replaces it (it is itself a validation: `ExcelValidation.dependentList(...)` takes the options of one, see above).

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

With `groupSubtotal(...)` the rules skip the subtotal rows (a subtotal is a sum, not a data row); call `.includeSubtotalRows()` on the rule to cover them too.

## Required Columns (Reading)

See [Reading — Required Columns](reading.md#required-columns).
