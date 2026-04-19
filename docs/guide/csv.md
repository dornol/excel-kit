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

## CSV Reading

```java
CsvReader.setter(Product::new)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .build(inputStream)
    .read(result -> { ... });
```

### Mapping Mode

```java
CsvReader.<Person>mapping(row -> new Person(
        row.get("Name").asString(),
        row.get("Age").asInt()))
    .build(inputStream)
    .read(result -> { ... });
```

### Map Mode

```java
CsvReader.forMap()
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
    });
```

### Reading Options

```java
CsvReader.setter(Row::new)
    .delimiter('\t')
    .charset(StandardCharsets.UTF_16)
    .headerRowIndex(1)            // skip first row
    .dialect(CsvDialect.TSV)
    .column((r, cell) -> r.name = cell.asString())
    .build(inputStream);
```

Index-based column mapping is also supported:
```java
CsvReader.setter(User::new)
    .columnAt(0, (u, cell) -> u.name = cell.asString())
    .columnAt(2, (u, cell) -> u.city = cell.asString())
    .build(inputStream);
```

Progress callback and Bean Validation work the same as Excel reading.
