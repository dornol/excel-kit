# CSV

> [Back to Index](index.md)

## CSV Writing

```java
CsvHandler ch = CsvWriter.<Row>create()
    .column("ID", r -> r.id())
    .column("Name", r -> r.name())
    .write(rows);

ch.writeTo(Path.of("rows.csv"));
```

### Columns

The column methods are those of `ExcelWriter`, with the settings a text file can have: `column`, `columnIf`, `constColumn`,
`constColumnIf` and `rowNumberColumn`. A column takes a configurer, `CsvColumnConfig`, with `nullValue(...)`: what to write
where the value is `null` (an empty field without it).

```java
CsvWriter.<Item>create()
    .rowNumberColumn("No.")                                       // 1, 2, 3, ...
    .column("Name", Item::name)
    .column("Status", Item::status, c -> c.nullValue("N/A"))      // null -> N/A
    .column("Qty", Item::qty, c -> c.nullValue(0))
    .write(items);
```

The placeholder goes through the same rules as any value: it is quoted when it has to be, and with the injection
defense on (below) a text that starts with `=`, `+`, `-` or `@` gets a leading `'`, so use `"N/A"` rather than `"-"`.
A CSV has no styles or types: format a date or a number in the function.

### Options

```java
CsvWriter.<Row>create()
    .delimiter('\t')                    // tab-separated (default: ',')
    .charset(StandardCharsets.UTF_16)   // custom encoding (default: UTF-8)
    .bom(false)                         // disable UTF-8 BOM (default: true)
    .column("Name", r -> r.name())
    .write(rows);
```

### Dialect Presets

```java
CsvWriter.create().dialect(CsvDialect.TSV);     // tab-separated
CsvWriter.create().dialect(CsvDialect.PIPE);    // pipe-separated
CsvWriter.create().dialect(CsvDialect.RFC4180); // strict RFC 4180
CsvWriter.create().dialect(CsvDialect.EXCEL);   // Excel-compatible
```

### Quoting Strategies

```java
CsvWriter.create().quoting(CsvQuoting.ALL);          // quote every field
CsvWriter.create().quoting(CsvQuoting.NON_NUMERIC);  // quote non-numeric fields
CsvWriter.create().quoting(CsvQuoting.MINIMAL);      // quote only when needed (default)
```

### Map-Based Writing

```java
CsvWriter.forMap("Name", "Age")
    .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
    .writeTo(outputStream);
```

### CSV Injection Defense

Enabled by default. Prefixes dangerous characters with `'`:

```java
CsvWriter.create().csvInjectionDefense(false);  // disable for trusted data
```

## Line Separator

Rows end with CRLF as RFC 4180 requires, on every operating system. `CsvDialect.TSV` and `PIPE` use LF.
Use `lineSeparator` to choose explicitly (it also applies to lines an `afterData` callback prints with
`println`, and wins over a dialect applied before it):

```java
CsvWriter.create().lineSeparator("\n");  // "\r\n" (default), "\n" or "\r"
```

## CSV Reading

```java
CsvReader.setter(Product::new)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .read(inputStream, result -> { ... });
```

CSV content can be detected from a sample before parsing. Delimiter detection
ignores characters inside quoted fields and uses the logical record selected by
`.headerRowIndex(...)`. The sample size defaults to 8192 bytes. Explicit charset
and delimiter settings, including those set through `.dialect(...)`, take
precedence over detection:

```java
CsvReader.<Row>mapping(this::map)
    .detectionSampleBytes(16 * 1024)
    .readDetectedWithSummary(source, result -> { ... });
```

Detected settings apply only to that read operation; the reusable reader keeps its configured
charset and delimiter for subsequent reads.

If the sample ends before the selected header, increase the sample size or
specify the delimiter explicitly.

`readDetected(InputStreamSource, ...)` and `readDetectedWithSummary(...)` close streams opened by
the source. Caller-owned `InputStream` instances remain open.

### What the parser does with sloppy quoting

Files written by `CsvWriter` always read back as written. Files written by hand are read as far as the parser
(opencsv) follows them. A space (or tab) before the opening quote of a quoted value is skipped by default
(`ignoreLeadingWhiteSpace`), so `1, "x,y"` reads `x,y`; a doubled quote, a line break and the delimiter inside the value
all work, with any delimiter, quote or escape character. These shapes are **not** read as one might expect, with either
setting:

| Input | Reads as | Why |
|---|---|---|
| `"x" ,2` (a space after the closing quote) | `x" ` | the closing quote followed by a space is taken as a literal |
| `  "x",1` (spaces before the quote of the **first** value of a row) | `  x` | the spaces stay; so a header such as `"a"` with spaces before it reads as ` a` |
| `1, ""` (an empty quoted value after a space) | ` "` | the quote is not recognised as a pair |
| `1, ab "cd"` (quotes inside unquoted text) | ` ab "cd` | the closing quote is dropped |

Write such files without the stray spaces, or read them with `.strictQuotes(true)` when every value is quoted.

### Mapping Mode

```java
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(),
        row.get("Age").asInt()))
    .read(inputStream, result -> { ... });
```

### Map Mode

```java
CsvReader.forMap()
    .read(inputStream, result -> {
        Map<String, String> row = result.data();
    });
```

### Reading Options

```java
CsvReader.setter(Row::new)
    .delimiter('\t')
    .charset(StandardCharsets.UTF_16)
    .quoteChar('\'')
    .escapeChar('\\')
    .strictQuotes(false)             // true drops text outside quotes: quote the header row too
    .ignoreLeadingWhiteSpace(true)   // default: a space before a quote is skipped, so  a, "b"  reads b
    .headerRowIndex(1)            // skip first row
    .dialect(CsvDialect.TSV)
    .column((r, cell) -> r.name = cell.asString())
    .read(inputStream, result -> { ... });
```

A quote inside a value must be escaped as `""` within a quoted field (`"12"" pipe"`). A bare quote in
the middle of an unquoted value (`12" pipe`) opens a quoted field, so the read fails with a
`CsvReadException` whose cause names the line where the unterminated record starts. The reader does
not guess, so rows are never merged silently; fix the source file or quote the value.

Index-based column mapping is also supported:
```java
CsvReader.setter(User::new)
    .columnAt(0, (u, cell) -> u.name = cell.asString())
    .columnAt(2, (u, cell) -> u.city = cell.asString())
    .read(inputStream, result -> { ... });
```

Progress callback and Bean Validation work the same as Excel reading.
