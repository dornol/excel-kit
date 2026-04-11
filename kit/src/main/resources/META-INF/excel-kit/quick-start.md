# excel-kit — Quick Start

> Other topics: [Index](../AI.md) | [Column Config](column-config.md) | [Reading](reading.md) | [Advanced](advanced.md) | [CSV](csv.md)

## Excel Writing

```java
record Person(long id, String name, int age) {}

var data = Stream.of(new Person(1, "Alice", 30), new Person(2, "Bob", 28));

ExcelHandler handler = new ExcelWriter<Person>()
        .column("ID", p -> p.id())
            .type(ExcelDataType.LONG)
            .alignment(HorizontalAlignment.RIGHT)
        .column("Name", p -> p.name())
        .column("Age", p -> p.age())
            .type(ExcelDataType.INTEGER)
        .write(data);

try (var os = Files.newOutputStream(Path.of("people.xlsx"))) {
    handler.write(os);
}
```

## Excel Writing (Multi-Sheet)

```java
try (var wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
    wb.<User>sheet("Users")
        .column("Name", User::getName)
        .column("Age", User::getAge)
        .write(userStream);

    wb.<Order>sheet("Orders")
        .column("ID", Order::getId)
        .column("Total", Order::getTotal, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0"))
        .write(orderStream);

    wb.finish().write(outputStream);
}
```

## Excel Reading (Setter Mode)

```java
class User {
    String name;
    Integer age;
}

new ExcelReader<>(User::new, null)
    .column("Name", (u, cell) -> u.name = cell.asString())
    .column("Age", (u, cell) -> u.age = cell.asInt())
    .build(Files.newInputStream(Path.of("users.xlsx")))
    .read(result -> {
        if (result.success()) {
            User u = result.data();
        }
    });
```

## Excel Reading (Mapping Mode — Records/Immutable)

```java
record PersonRecord(String name, Integer age) {}

ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
    row.get("Name").asString(),
    row.get("Age").asInt()
)).build(inputStream).read(result -> {
    if (result.success()) {
        PersonRecord p = result.data();
    }
});
```

## CSV Writing

```java
CsvWriter<Row> csv = new CsvWriter<>();
csv.column("ID", r -> r.id())
   .column("Name", r -> r.name())
   .write(rows)
   .write(outputStream);
```

## CSV Reading

```java
new CsvReader<>(Product::new, null)
    .column("Name", (p, cell) -> p.name = cell.asString())
    .column("Price", (p, cell) -> p.price = cell.asInt())
    .build(inputStream)
    .read(result -> { ... });
```

## Output Consumption

- `write(out)` — stream directly to OutputStream
- `consumeFile(path)` — write to file
- `password("pw")` on writer — automatic encryption
- `consumeOutputStreamWithPassword(out, "pw")` — late-binding encryption

Output is consume-once via `ExcelHandler` / `CsvHandler`.

## Map-Based (No POJO)

```java
// Write
new ExcelMapWriter("Name", "Age").write(Stream.of(
    Map.of("Name", "Alice", "Age", 30)
)).write(out);

// Read
new ExcelMapReader().build(inputStream).read(result -> {
    Map<String, String> row = result.data();
    String name = row.get("Name");
});
```
