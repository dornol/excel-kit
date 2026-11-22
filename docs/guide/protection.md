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
```

> Using both `.password()` and `writeTo(out, password)` on the same handler throws `IllegalStateException`.
> A null or blank password throws `IllegalArgumentException` and consumes the handler: its workbook is
> closed and it cannot be written afterwards. The `IllegalStateException` above leaves the handler usable.

### Resources and a JDK 24+ Warning

Encrypting does not need the file's size in heap: the encrypted container is built in a temporary file, so a
workbook of hundreds of megabytes encrypts within a small heap. The disk is used instead: for a moment the plain and
the encrypted copy exist side by side (about twice the size of the file), and both are deleted when the write ends.

On **JDK 24 and later** the first encryption in a JVM may print this to stderr:

```
WARNING: sun.misc.Unsafe::invokeCleaner has been called by org.apache.poi.poifs.nio.CleanerUtil
WARNING: sun.misc.Unsafe::invokeCleaner will be removed in a future release
```

It comes from Apache POI, not from excel-kit. POI memory-maps the temporary file and releases the mapping through
`sun.misc.Unsafe`, which [JEP 498](https://openjdk.org/jeps/498) warns about on first use. The output is not affected,
and reading encrypted files does not trigger it. To silence it, start the JVM with
`--sun-misc-unsafe-memory-access=allow`. Fixing it is up to POI; it should disappear with a POI release that no
longer uses `Unsafe` for this.

## Security Notes

### Reading Untrusted Files

When reading Excel files from untrusted sources (user uploads, external systems),
call `configureLargeFileSupport()` at application startup to protect against
ZIP bomb and decompression attacks:

```java
// Call once at startup — sets JVM-global POI limits
ExcelReader.configureLargeFileSupport();
```

Without this, the default POI limits may allow malicious `.xlsx` files to
consume excessive memory or disk during decompression.

### Encrypted File Decryption

When reading password-encrypted files, the library decrypts to a temporary file
on disk before parsing. The plaintext temp file is deleted immediately after use
but is not cryptographically overwritten.

For sensitive data, deploy on systems with full-disk encryption (LUKS, BitLocker,
FileVault) so that temp files are encrypted at rest.

### Temporary File Handling

- Temp directories: POSIX `rwx------`, Windows ACL restricted to current user
- Automatic cleanup after each operation (success or failure)
- Fallback: `deleteOnExit()` if immediate deletion fails
- UUID-based naming to prevent path prediction
- On Windows, ACL restriction may silently fall back to a warning log if the
  filesystem does not support ACLs. Use NTFS for reliable access control.

### XXE Protection

XML parsing is delegated to Apache POI, which disables external entity resolution
by default since version 5.0. Ensure your POI dependency is 5.0 or later.

### CSV Injection Defense

`CsvWriter` prefixes dangerous leading characters (`=`, `+`, `-`, `@`, `\t`, `\r`) with `'`.

Disable for trusted data:
```java
CsvWriter.create().csvInjectionDefense(false);
```

### Formula Safety

`ExcelDataType.FORMULA` is for developer-controlled formula strings. Do not pass untrusted user input — use `STRING` type instead.
