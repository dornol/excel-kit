# Protection & Encryption

> [Back to Index](index.md)

## Sheet Protection

Protect sheets with a password and selectively unlock columns:

```java
writer
    .column("Name", Product::name, c -> c.locked(false))   // editable
    .column("Price", p -> p.price(), c -> c.locked(true))   // read-only
    .protectSheet("password123")
    .write(data);
```

When sheet protection is enabled, all cells are locked by default. Use `.locked(false)` to allow editing.

> Sheet protection is a UI-level deterrent, not cryptographic security. Use password encryption for actual data protection.

## Workbook Protection

Prevent adding, deleting, renaming, or reordering sheets:

```java
writer.protectWorkbook("password123");
```

Can be combined with `protectSheet()` — workbook protection prevents structural changes, sheet protection prevents cell editing.

## Password Encryption

Uses Apache POI's Agile encryption mode (AES-256).

### Eager — Set on Writer

```java
ExcelHandler handler = ExcelWriter.<Product>create()
    .password("P@ssw0rd!")
    .column("Name", Product::name)
    .write(data);

handler.writeTo(outputStream);
```

Works the same with `ExcelWorkbook`:
```java
try (var wb = ExcelWorkbook.create()) {
    wb.password("P@ssw0rd!");
    wb.<Product>sheet("Products").column("Name", Product::name).write(data);
    wb.finish().writeTo(outputStream);
}
```

### Late-Binding — Password at Output Time

When the service layer builds the handler but the password is only known at the presentation layer:

```java
// Service
ExcelHandler handler = writer.column(...).write(dataStream);

// Presentation
handler.writeTo(outputStream, "P@ssw0rd!");   // OutputStream
handler.writeTo(path, "P@ssw0rd!");           // File path
handler.writeTo(outputStream, pwChars);       // char[] (zeroed after use)
handler.writeTo(path, pwChars);               // Path + char[]
```

> Using both `.password()` and `writeTo(out, password)` on the same handler throws `IllegalStateException`.

## Security Notes

### Temporary File Handling

- Temp directories: POSIX `rwx------`, Windows ACL restricted to current user
- Automatic cleanup after each operation (success or failure)
- Fallback: `deleteOnExit()` if immediate deletion fails
- UUID-based naming to prevent path prediction

### CSV Injection Defense

`CsvWriter` prefixes dangerous leading characters (`=`, `+`, `-`, `@`, `\t`, `\r`) with `'`.

Disable for trusted data:
```java
CsvWriter.create().csvInjectionDefense(false);
```

### Formula Safety

`ExcelDataType.FORMULA` is for developer-controlled formula strings. Do not pass untrusted user input — use `STRING` type instead.
