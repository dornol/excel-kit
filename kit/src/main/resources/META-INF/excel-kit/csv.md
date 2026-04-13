# excel-kit — CSV Features

> Other topics: [Index](../AI.md) | [Quick Start](quick-start.md) | [Column Config](column-config.md) | [Reading](reading.md) | [Advanced](advanced.md)

## CSV Writing

```java
CsvWriter<Product> csv = CsvWriter.create();
CsvHandler handler = csv
    .column("Name", Product::name)
    .column("Price", Product::price)
    .write(productStream);

handler.writeTo(outputStream);
```

## CSV Writing Configuration

```java
CsvWriter.<Product>create()
    .delimiter('\t')                              // tab-separated (default: comma)
    .charset(StandardCharsets.UTF_8)              // default: UTF-8
    .bom(true)                                    // UTF-8 BOM for Excel (default: true)
    .quoting(CsvWriter.QuotingStrategy.ALL)       // MINIMAL (default), ALL, NON_NUMERIC
    .csvInjectionDefense(false)                   // disable formula prefixing (default: true)
    .column("Name", Product::name)
    .write(stream);
```

## CSV Dialects

```java
CsvWriter.<Product>create()
    .dialect(CsvDialect.RFC4180)    // standard CSV
    .column(...).write(stream);

CsvWriter.<Product>create()
    .dialect(CsvDialect.TSV)        // tab-separated
    .column(...).write(stream);
```

| Dialect | Delimiter | Quoting | BOM |
|---------|-----------|---------|-----|
| `RFC4180` | `,` | MINIMAL | false |
| `EXCEL` | `,` | MINIMAL | true |
| `TSV` | `\t` | MINIMAL | false |
| `PIPE` | `\|` | MINIMAL | false |

## CSV Reading

### Setter Mode
```java
new CsvReader<>(Product::new, null)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .build(inputStream)
    .read(result -> {
        if (result.success()) { Product p = result.data(); }
    });
```

### Mapping Mode (Records)
```java
CsvReader.<Product>mapping(row -> new Product(
    row.get("Name").asString(),
    row.get("Price").asInt()
)).build(inputStream).read(result -> { ... });
```

### Map Mode
```java
CsvReader.forMap()
    .delimiter(',')
    .charset(StandardCharsets.UTF_8)
    .headerRowIndex(0)
    .build(inputStream)
    .read(result -> {
        Map<String, String> row = result.data();
    });
```

## CSV Reading Configuration

```java
new CsvReader<>(Product::new, null)
    .delimiter(',')                    // default: ','
    .charset(StandardCharsets.UTF_8)   // default: UTF-8
    .headerRowIndex(0)                 // default: 0
    .dialect(CsvDialect.RFC4180)       // preset config
    .onProgress(10_000, (count, cursor) -> log.info("Read {}", count))
    .column(...)
    .build(inputStream);
```

## CSV Map Writer

```java
CsvWriter<Map<String, Object>> writer = CsvWriter.forMap("Name", "Age", "City");
writer.delimiter('\t')
      .charset(StandardCharsets.UTF_8)
      .bom(true)
      .write(Stream.of(
          Map.of("Name", "Alice", "Age", "30", "City", "Seoul")
      )).writeTo(out);
```

## CSV Injection Defense

By default, cells starting with `=`, `+`, `-`, `@`, `\t`, `\r` are prefixed with a single quote to prevent formula injection. Disable with:

```java
CsvWriter.create().csvInjectionDefense(false)
```

## Progress Callback

```java
csv.column("Name", Product::name)
   .onProgress(50_000, (count, cursor) -> log.info("Wrote {} rows", count))
   .write(stream);
```
