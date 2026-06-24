# Changelog

All notable changes to this project will be documented in this file.

## [0.18.0] - 2026-06-24

### Added

- Header aliases for Excel and CSV readers via `column(List<String>, ...)`.
- Strict reader header validation via `strictHeaders()` / `requireHeaders()`.
- Duplicate header handling via `DuplicateHeaderPolicy` (`FIRST`, `LAST`, `FAIL`).
- Physical source row tracking via `ReadResult.fileRowNum()` and `RowError.fileRowNum()`.
- Structured cell-level read errors via `CellError` and `RowError.cellErrors()`.
- Schema read aliases and required columns via `ExcelKitSchema.Builder.column(name, aliases, ...)`
  and `requiredColumn(...)`.
- Strict selected-column validation for `ExcelReader.forMap(...)` and `CsvReader.forMap(...)`.
- Example app MockMvc upload smoke tests for schema-based Excel and CSV reads.

### Changed

- Map-mode readers now apply the configured duplicate header policy consistently
  when building output maps.

## [0.17.1] - 2026-06-24

### Added

- Example app smoke test coverage now starts the Spring context without Docker
  and verifies a real MyBatis mapper query against an H2 database.

### Changed

- CSV and Excel read handlers are now explicitly one-shot. Calling `read()`,
  `readStrict()`, or `readAsStream()` consumes the handler; a second read attempt
  now fails with a clear `ExcelKitException` instead of a lower-level temp-file
  error.

### Fixed

- CSV `readAsStream()` cleanup coverage now verifies temporary resources are
  removed after failed row mapping results are consumed.

## [0.17.0] - 2026-06-22

### Added

- `ExcelWriter.create(opts -> opts.rowAccessWindowSize(...))` and
  `ExcelWorkbook.create(opts -> opts.rowAccessWindowSize(...))` for configuring
  SXSSF's row access window at workbook construction time.

### Fixed

- `ExcelReadHandler.readAsStream()` now releases temporary reader resources when
  the stream is fully consumed, even if callers do not explicitly close the stream.
- `ExcelReader.getSheetHeaders(...)` now throws `ExcelReadException` when the
  requested sheet index does not exist, matching the main read path.
- `ExcelReader.sheetIndex(...)` now accepts any non-negative index and lets the
  workbook contents determine whether that sheet exists.
- `password(char[])` now rejects blank passwords, matching `password(String)`.

## [0.16.15] - 2026-04-19

### Added

- **`countRows()`** on `ExcelReader` — opt-in pre-scan that counts data rows before
  parsing, making percentage-based progress reporting possible (e.g., via SSE).
  The total is available via `Cursor.getTotalRows()` in the `ProgressCallback`.
  ```java
  ExcelReader.setter(MyDto::new)
      .countRows()
      .onProgress(500, (processed, cursor) -> {
          int percent = (int) (processed * 100 / cursor.getTotalRows());
      })
      .build(inputStream).read(result -> { ... });
  ```
- **`rowStyle(predicate, config)`** — conditional row-level styling on `ExcelWriter`,
  `ExcelSheetWriter`, and `TemplateListWriter`. Unlike `rowColor()` (background only),
  supports bold, italic, strikethrough, font size, font color, and background color.
  Multiple rules can be registered; the first matching predicate wins.
  ```java
  writer.rowStyle(p -> p.price() > 10000,
      style -> style.bold(true).backgroundColor(ExcelColor.LIGHT_YELLOW));
  ```
- **`ExcelImage.size(width, height)`** — control image cell span (columns × rows).
  Default remains 1×1. Previously images were always anchored to a single cell.
- **`ExcelImage.fromUrl(url)`** — download an image from a URL and embed it directly.
  Auto-detects PNG/JPEG from URL extension or magic bytes. 10-second timeout.
  ```java
  .column("Photo", user -> ExcelImage.fromUrl(user.getPhotoUrl()),
      c -> c.type(ExcelDataType.IMAGE))
  ```
- **`Cursor.getTotalRows()`** — returns the total row count when known (via
  `countRows()`), or `-1` if unknown.

### Changed

- **`ExcelImage`** refactored from `record` to `class` to support `size()` chaining
  and `fromUrl()` factory. The constructor signature and `data()`/`imageType()` accessors
  are unchanged — existing code compiles without modification.

## [0.16.14] - 2026-04-19

### Added

- **`documentProperty(key, value)`** — set Excel document metadata (title, author,
  subject, keywords, description, category) on `ExcelWriter` and `ExcelWorkbook`.
  Standard keys map to core properties; other keys become custom properties.
- **`namedRange(name, columnIndex)`** — fluent named range registration on any writer.
  Automatically creates a workbook-scoped named range covering all data rows in the
  specified column, replacing manual `afterData` callback usage.
- **`headerStyle(cfg -> ...)`** — configure header cell alignment, bold, border style,
  and wrap text via `HeaderStyleConfig`. Available on `ExcelWriter` and `ExcelWorkbook`.
- **`password(char[])`** — char-array overload for `ExcelWriter.password()` and
  `ExcelWorkbook.password()`. The array is copied internally and zeroed after encryption,
  preventing password from lingering in heap as an immutable String.

### Changed

- **Internal: `ExcelColumn` constructor** simplified from 22 positional parameters to
  `(name, function, style, setter, ColumnStyleConfig)`. Adding new column fields now
  requires changes in only 2 places instead of 5+.
- **Internal: `AbstractSheetWriter` base class** extracted from `ExcelWriter` and
  `ExcelSheetWriter`, eliminating 22 duplicated sheet-config methods (~320 lines removed).
- **Internal: `AbstractReader` base class** extracted from `ExcelReader` and `CsvReader`,
  eliminating 10 duplicated column-registration methods (~177 lines removed).
- **Internal: `writeGroupAndColumnHeaders`** split into 5 focused helpers
  (`createHeaderRows`, `buildGroupGrid`, `populateHeaderCells`,
  `applyHorizontalMerges`, `applyVerticalMerges`).
- **Internal: `ExcelHandler`** stores password as `char[]` internally (zeroed after use).

### Documentation

- **Guide restructured** — `docs/guide.md` (2219 lines) split into `docs/guide/` directory
  with 13 topic files. `index.md` contains inline code examples covering 80% of use cases;
  topic files provide full details.
- **AI/LLM integration section** added to README with classpath vs HTTP agent guidance.
- **Security notes** expanded in `docs/guide/protection.md`: ZIP bomb protection,
  encrypted temp file handling, Windows ACL fallback, XXE delegation to POI.

## [0.16.13] - 2026-04-15

### Added

- **Multi-row header support** — `ExcelReader.headerRows(int)` lets the reader
  combine N header rows into the effective column names, using the bottom-most
  non-blank value per column. This round-trips files written with multi-level
  `group(...)` on the writer side (where column header cells are often blank
  because they're part of a vertical merge with a group label above).
  ```java
  reader
      .headerRowIndex(1)     // last header row (0-based)
      .headerRows(2)         // 2 header rows total (group row + column header row)
      .build(in)
      .read(record -> ...);
  ```
  Default remains `headerRows(1)`; existing single-row behavior unchanged
  (including preserving empty-string headers).

## [0.16.12] - 2026-04-15

### Added

- **Split success/error read callbacks** — `read(Consumer<T> onSuccess,
  Consumer<RowError> onError)` delivers successfully parsed rows to one
  callback and failed rows (validation or mapping) to another. The library
  buffers nothing — the caller decides how to manage error memory (log,
  keep top N, stream to elsewhere, or abort by throwing).
  ```java
  reader.read(
      record -> process(record),
      err -> {
          if (err.type() == RowError.Type.VALIDATION) { ... }
          else { log.warn("row {} mapping failed", err.rowNum(), err.cause()); }
      });
  ```
- **`RowError` record** — carries row ordinal (1-based, excluding header),
  category (`VALIDATION` / `MAPPING`), human-readable messages, and the
  original cause when available.

### Changed

- `ReadResult<T>` record gains a nullable `cause` field so mapping-stage
  exceptions are preserved alongside the usual messages. The 3-arg
  constructor is retained for backward compatibility.

## [0.16.11] - 2026-04-15

### Added

- **Group header comments** — `writer.groupComment(String text, String... path)`
  and `groupComment(ExcelCellComment, String... path)` attach a note to a
  merged group header cell, identified by its path (outermost first). No-op
  if no column declares that path.
  ```java
  .column("Q1", Row::q1, c -> c.group("Financial", "Revenue"))
  .groupComment("Quarterly revenue", "Financial", "Revenue")
  ```
- **Per-column header background color** — `headerBackgroundColor(ExcelColor)` /
  `headerBackgroundColor(int r, int g, int b)`. Overrides the workbook-wide
  `headerColor` for a single column only. Useful for highlighting alert or
  required columns.
- **`headerRowHeight(float points)`** — writer-level setter that applies the
  given height to every header row (including group header rows in multi-level
  setups). Pass `0` for default.
- **`rowNumberColumn(String name)`** — convenience for a 1-based sequential
  row-number column. Equivalent to a manual `column(name, (r, cursor) ->
  cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG))` call. Works
  across auto-rollover sheets.

## [0.16.10] - 2026-04-15

### Fixed

- Vertically merged group header cells rendered with bottom-aligned text
  because the column name value existed in both the top-left and the
  original column header row cell of the merge region. Non-top cells in
  a vertical merge are now explicitly blanked so Excel honors the
  header style's `VerticalAlignment.CENTER`.

## [0.16.9] - 2026-04-15

### Added

- **Multi-level group headers** — `group(String... levels)` accepts N levels,
  ordered outermost (top) to innermost (just above the column header). Previously
  only one level was supported.

  ```java
  .column("Q1", Row::q1, c -> c.group("Financial", "Revenue", "2025"))
  ```
  Produces 3 group rows + 1 column header row. Adjacent columns with equal values
  on the same row are horizontally merged; columns with fewer levels are vertically
  merged with the column header cell. Header comments attach to the top-left cell
  of the merged region.

### Changed (Breaking)

- `ColumnStyleConfig.group(String)` signature changed to `group(String... levels)`.
  Source-compatible — existing `.group("X")` calls still work unchanged. The
  internal field `groupName` (String) became `groupNames` (String[]); if you
  reflected on it or read `ExcelColumn.getGroupName()`, use `getGroupNames()`
  instead.

## [0.16.8] - 2026-04-14

### Added

- `ColumnStyleConfig.headerComment(ExcelCellComment)` — overload for full
  customization of header comment (author, width, height) via the
  `ExcelCellComment` record.
- `ColumnStyleConfig.commentSize(int width, int height)` — column-level
  comment box size applied to both header and data-cell comments.
- `ExcelCellComment` fluent builders: `ExcelCellComment.of(text)`,
  `.author(String)`, `.size(int, int)`.
- `Comment.setAuthor(...)` is now actually propagated to POI (previously
  stored in the record but never written out).

  Precedence: `ExcelCellComment.size()` > `commentSize()` > POI default (2×3).

  ```java
  // Simple
  .headerComment("Enter in YYYY-MM-DD format")

  // Column-level size for all comments in the column
  .column(..., c -> c.comment(p -> "note").commentSize(4, 5))

  // Full customization
  .headerComment(ExcelCellComment.of("Reviewed").author("System").size(4, 3))
  ```

### Changed (Breaking)

- **`ExcelCellComment` record canonical constructor signature changed** from
  `(String text, String author)` to `(String text, String author, int width, int height)`.
  The old convenience constructor `new ExcelCellComment(text)` / `(text, author)`
  were removed. Use `ExcelCellComment.of(text)` going forward.

  This record was previously dead code (never instantiated outside its own
  file), so real-world impact is nil.

## [0.16.7] - 2026-04-14

### Added

- `ColumnStyleConfig.headerComment(String)` — attach a static comment (note) to
  a column's header cell. Useful for documenting expected input format or units
  (e.g., `"YYYY-MM-DD format"`, `"In KRW"`) as an Excel cell comment visible on
  hover. Complements the existing per-row `comment(Function)` targeting data
  cells.
  ```java
  .column("Birth Date", User::birth, c -> c
      .type(ExcelDataType.DATE)
      .headerComment("Enter in YYYY-MM-DD format"))
  ```
  With group headers, the comment attaches to the column header row (not the
  group row).

## [0.16.6] - 2026-04-13

### Added

- `ExcelHandler.writeTo(Path, String)` and `ExcelHandler.writeTo(Path, char[])` —
  convenience overloads that encrypt and write directly to a file path, mirroring
  the existing `writeTo(OutputStream, String/char[])` overloads. Useful for CLI/batch
  tools that don't need a Spring `StreamingResponseBody`. The `char[]` variant zeroes
  the password array after use.
- `ExcelWriter.freezeCols(int)` / `ExcelSheetWriter.freezeCols(int)` — freeze N
  columns from the left. Previously required `freezePane(N, 0)`.

### Changed (Breaking)

- **`ExcelWriter.freezePane(int rows)` renamed to `freezeRows(int)`.** Single-arg
  `freezePane(N)` was ambiguous — a caller reading `.freezePane(3)` couldn't tell
  whether it froze rows or columns without checking the Javadoc. The new names
  make intent explicit:
  - `freezeRows(int)` — freeze N rows below the header (replaces `freezePane(int)`)
  - `freezeCols(int)` — freeze N columns from the left (new)
  - `freezePane(int cols, int rows)` — freeze both axes (unchanged)
  Same rename applies to `ExcelSheetWriter`.

  Migration: `.freezePane(1)` → `.freezeRows(1)`.

- **`new CsvWriter<T>()` replaced by `CsvWriter.<T>create()`.** The public no-arg
  constructor is now package-private, matching the `ExcelWriter.create()` /
  `ExcelWorkbook.create()` pattern. Users switching between Excel and CSV writers
  now share one mental model.

  Migration: `new CsvWriter<Row>()` → `CsvWriter.<Row>create()`.

### Changed

- Strengthened Javadoc on `ExcelReadHandler` and `CsvReadHandler` with an explicit
  "Resource management" section, clarifying when temp files are released and why
  `readAsStream()` must be used with try-with-resources.

## [0.16.5] - 2026-04-13

### Changed (Breaking)

- **Unified `FileHandler` output API under `writeTo(...)`.** Every way a handler
  emits its payload is now a `writeTo` overload:
  - `FileHandler.write(OutputStream)` → `writeTo(OutputStream)`
  - `FileHandler.toFile(Path)` → `writeTo(Path)`
  - `ExcelHandler.consumeOutputStreamWithPassword(OutputStream, String)` → `writeTo(OutputStream, String)`
  - `ExcelHandler.consumeOutputStreamWithPassword(OutputStream, char[])` → `writeTo(OutputStream, char[])`
  - `CsvHandler.write(OutputStream)` → `writeTo(OutputStream)`
  Motivation: the old names collided conceptually with `ExcelWriter.write(Stream<T>)`
  (which produces a handler) — reading `.write(stream).write(out)` looked like two
  invocations of the same operation. They are not: the first pulls data *in*, the
  second pushes bytes *out*. Renaming the handler side to `writeTo` makes the
  direction explicit and frees the `write(Stream)` verb on the writer.
  `consumeOutputStreamWithPassword` (30 chars, different verb family) was also
  replaced by a plain `writeTo` overload so every output-stage call reads
  `.writeTo(destination, ...)`.

Migration:

```java
// Before
writer.column(...).write(stream).write(out);
writer.column(...).write(stream).toFile(path);
writer.column(...).write(stream).consumeOutputStreamWithPassword(out, "pw");
writer.column(...).write(stream).consumeOutputStreamWithPassword(out, pwdChars);

// After
writer.column(...).write(stream).writeTo(out);
writer.column(...).write(stream).writeTo(path);
writer.column(...).write(stream).writeTo(out, "pw");
writer.column(...).write(stream).writeTo(out, pwdChars);

// Spring StreamingResponseBody
.body(handler::write)   →   .body(handler::writeTo)
```

`ExcelWriter.password(String)` / `ExcelWorkbook.password(String)` is unchanged —
pre-set-then-`writeTo(out)` remains the preferred path. Passing the password to
`writeTo(out, pw)` still throws `IllegalStateException` if `password()` was
already set at the writer level.

## [0.16.4] - 2026-04-13

### Changed (Breaking)

- **`ExcelWriter.builder()` / `ExcelWorkbook.builder()` removed** — use
  `ExcelWriter.create()` / `ExcelWriter.create(opts -> ...)` and
  `ExcelWorkbook.create()` / `ExcelWorkbook.create(opts -> ...)` instead.
  The two-step `builder().xxx.build()` pattern was awkward because most
  configuration (columns, filters, callbacks, etc.) was already a fluent chain
  on the writer itself — `build()` served no real boundary.
- **`InitOptions` is now scoped to what POI forces at construction time.** Only
  `rowAccessWindowSize` remains in `InitOptions` (SXSSFWorkbook's constructor
  argument — cannot be changed after the workbook is created). The former
  `color` and `maxRows` builder methods are now fluent setters on the
  writer/workbook itself.
- **`color()` renamed to `headerColor()`** on both `ExcelWriter` and
  `ExcelWorkbook`, for consistency with the existing `tabColor()` and
  `rowColor()` methods and to make the target of the color unambiguous.
- **`ExcelWriter.forMap(Builder, String...)` → `forMap(Consumer<InitOptions>, String...)`.**
  Header color / max rows for a map writer are now set via fluent
  `.headerColor()` / `.maxRows()` on the returned writer.

Migration:

```java
// Before
ExcelWriter.<User>builder()
    .color(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .rowAccessWindowSize(500)
    .build()
    .column("Name", User::name)
    .write(stream);

// After
ExcelWriter.<User>create(opts -> opts.rowAccessWindowSize(500))
    .headerColor(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .column("Name", User::name)
    .write(stream);

// Or, when the default row window (1000) is fine — the common case:
ExcelWriter.<User>create()
    .headerColor(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .column("Name", User::name)
    .write(stream);
```

### Added

- **`ExcelWriter.headerColor(ExcelColor)` / `ExcelWorkbook.headerColor(ExcelColor)`**
  fluent setters (rebuild header style, preserving already-configured font
  settings).
- **`ExcelWriter.maxRows(int)`** fluent setter.

### Tests

- Strengthened header-color tests to verify exact RGB values instead of merely
  asserting "some fill color is present" (which passed even for the default
  white header).
- Added regression tests for the new `headerColor()` setter composition with
  `headerFontName/Size` (ensures font settings survive the internal
  `headerStyle` rebuild).

## [0.16.3] - 2026-04-12

### Fixed

- **readAsStream() producer thread deadlock** — when consumer closed the stream early
  (e.g., `limit()`, `takeWhile()`), the producer thread could block forever on
  `queue.put()`. Replaced with `offer(timeout)` + interrupt check loop so the producer
  exits cleanly on consumer close.
- **Temp file leak on read initialization failure** — if `Files.copy()` failed during
  `initTempFile()`, already-created temp directory and file were not cleaned up. Now
  calls `close()` before rethrowing the exception.

## [0.16.2] - 2026-04-12

### Fixed

- **Decrypted temp file not cleaned up** — when reading password-protected Excel files,
  the decrypted temp file was not deleted after reading, leaving sensitive data on disk.
  Now explicitly deleted in a finally block.
- **DataBar gradient maxColor not applied** — `dataBar(minColor, maxColor)` stored the
  max color but never wrote it to the Excel XML. 2-color gradient data bars now correctly
  render both colors.
- **ExcelDataType.FORMULA DDE security warning** — added javadoc warning that DDE formulas
  (e.g., `cmd|'/c calc'`) can execute commands. No runtime guard is applied because DDE
  pipe syntax is used by legitimate integrations (Bloomberg, Reuters). Input validation
  is the caller's responsibility.

## [0.16.1] - 2026-04-12

### Fixed

- **CSV injection whitespace bypass** — values like `" =cmd|'/c calc"` (leading spaces
  before formula characters) were not caught by the injection defense. Now detects formula
  characters after leading spaces.
- **Required column bypass in sparse Excel rows** — when trailing cells were empty (SAX
  omits them), required column validation was skipped. Now correctly reports missing
  required columns in sparse rows.
- **ExcelSheetWriter rollover header color loss** — custom `headerFontColor()` was not
  preserved on auto-rollover sheets. Fixed by passing workbook and style cache to the
  rollover header writer.
- **beforeHeaderWriter position ignored on rollover** — if `beforeHeaderWriter` wrote rows,
  rollover sheets placed headers at row 0 instead of after the preamble. Now captures and
  applies the preamble row offset.
- **Summary row overlap with afterData** — when both `afterData()` and `summary()` were
  used, the summary writer's return value was discarded, causing rows to overlap. Now
  correctly chains the row position.
- **ExcelImage null/empty data** — `ExcelImage.png(null)` and `ExcelImage.jpeg(new byte[0])`
  now throw `IllegalArgumentException` at creation time instead of failing later during write.

### Changed

- **ExcelSheetWriter.write() single-call guard** — calling `write()` twice on the same
  sheet now throws `ExcelWriteException` instead of silently corrupting data.

### Improved

- **docs/guide.md** — all `new ExcelReader<>(T::new, null)` / `new CsvReader<>(T::new, null)`
  patterns updated to modern `ExcelReader.setter(T::new)` / `CsvReader.setter(T::new)` API.
- **Progress callback warning** — documentation now warns that callbacks run on the
  read/write thread and must be fast and non-blocking.
- **Removed `docs/llms.txt`** — redundant with `META-INF/AI.md` in the JAR. Removed from
  project, CI workflow, and release checklist.

## [0.16.0] - 2026-04-12

### Added

- **`nullValue(Object)`** on `ColumnStyleConfig` — sets a default value to write when a
  column function returns null (e.g., `c -> c.nullValue("N/A")`). Works with `defaultStyle()`
  for writer-level fallback, overridable per column.
- **`freezePane(int cols, int rows)`** overload on `ExcelWriter` and `ExcelSheetWriter` —
  freezes both columns and rows simultaneously. The existing `freezePane(int rows)` is unchanged.
- **`required()`** on `ExcelReader` and `CsvReader` — marks the last registered column as
  required. Blank/empty cells produce a validation error in `ReadResult`.
- **`ReadColumn.required()`** — returns a new `ReadColumn` with `isRequired=true`.

### Changed (Breaking)

- **`FileHandler.write()` no longer throws `IOException`** — I/O errors are wrapped as
  unchecked exceptions (`ExcelWriteException`, `CsvWriteException`). `toFile()` wraps as
  `ExcelKitException`. Callers no longer need try-catch for checked exceptions.
- **`ExcelHandler.write()` and `consumeOutputStreamWithPassword()` no longer throw
  `IOException`** — same unchecked wrapping as above.
- **`AbstractReadHandler.readAsStream()` is now abstract** — the default implementation
  (which loaded all results into memory) is removed. Both `ExcelReadHandler` and
  `CsvReadHandler` already had proper lazy overrides.

### Fixed

- **ExcelWriter workbook leak on stream exception** — if the data stream threw during
  `write()`, the `SXSSFWorkbook` was never closed, leaking temp files. Now wrapped in
  try-catch with `closeWorkbookQuietly()`.
- **CsvReadHandler.readAsStream() CSVReader leak** — if initialization failed after creating
  the CSVReader but before returning the stream, the reader was not closed. Fixed with
  null-safe cleanup in all error paths.
- **CellData.asBigDecimal() precision loss** — previously routed through `asNumber()` →
  `Double` → `BigDecimal(toString())`, which lost precision for very large numbers. Now
  parses the cleaned string directly as `BigDecimal`.
- **CsvHandler.write() javadoc** — documented that IOException is wrapped as
  `CsvWriteException` (unchecked), clarifying the `FileHandler` contract.

### Improved

- **CellData regex pre-compiled** — currency symbol pattern (`[$,₩€%원]`) compiled once
  as `static final Pattern` instead of per-call `replaceAll()`.
- **Exception catch simplified** — `catch (A) throw; catch (B) throw;` patterns replaced
  with `catch (A | B) throw;` in both read handlers.
- **AtomicLong for lambda counters** — `long[]` array wrapper replaced with `AtomicLong`
  in `readStrict()` and `CsvReadHandler.readAsStream()`.
- **Duplicate header detection** — `ExcelReadHandler` now logs a warning when duplicate
  header names are found, noting that only the first occurrence is used in mapping mode.
- **README restructured** — reduced from 2043 to 231 lines. Detailed guide moved to
  `docs/guide.md`.
- **Release checklist updated** — reflects new README structure and `docs/guide.md`.

## [0.15.0] - 2026-04-12

### Added

- **`FileHandler.toFile(Path)`** — convenience default method that writes directly to a
  file path without manually opening an OutputStream. Works for both Excel and CSV:
  `writer.write(stream).toFile(Path.of("output.xlsx"))`.
- **`ExcelReader.password(String)`** — enables reading password-encrypted Excel files.
  Symmetric with `ExcelWriter.password()` for writing. Uses POI's Decryptor API.
- **`CsvWriter.constColumnIf(name, condition, value)`** — conditional constant column,
  symmetric with `ExcelWriter.constColumnIf()`.

### Fixed

- **BookJpaRepository JPQL path** — example app's JPQL query referenced a wrong package
  (`book.domain.BookDto` instead of `book.application.port.out.BookDto`). Pre-existing bug
  unrelated to library code.

### Changed

- **Release checklist strengthened** — CLAUDE.md now requires `./gradlew :example:bootRun`
  to verify Spring context initialization (catches JPQL/HQL errors that `compileJava` misses),
  plus `./gradlew :kit:javadoc` for zero-warning verification.

## [0.14.0] - 2026-04-12

v0.14.0 is a comprehensive API cleanup and internal refactoring release.

### Changed (Breaking)

- **Package renamed: `shared` → `core`** — `io.github.dornol.excelkit.shared` is now
  `io.github.dornol.excelkit.core`. All imports must be updated. The package holds the
  library's central abstractions (FileHandler, CellData, ReadColumn, RowData, etc.) —
  "core" describes their role more accurately.
- **`ExcelRowFunction` and `CsvRowFunction` deleted** — both were empty marker interfaces
  extending `RowFunction` with zero additional methods. Use `RowFunction` (in `core/`)
  directly.
- **`ExcelConsumer` renamed to `WriteRowCallback`** — better signals that it's a per-row
  callback during writes, not a generic consumer.
- **`ExcelReadColumn` and `CsvReadColumn` merged into `ReadColumn`** — identical records
  consolidated into `core/ReadColumn<T>`.
- **`ExcelWorkbook` constructors removed** — use `ExcelWorkbook.builder()` (same pattern
  as `ExcelWriter.builder()`).
- **`ExcelReader.configureLargeFileSupport()` moved to `ExcelKitConfig`** — JVM-global
  POI configuration belongs at application bootstrap level, not on a reader class.
- **`ExcelWorkbook.finish()` now throws on duplicate calls** — second call to `finish()`
  throws `ExcelWriteException` instead of silently creating a new handler.

### Added

- **`ExcelReader.setter()` / `CsvReader.setter()` static factories** — all three read
  modes now have symmetric entry points: `setter()`, `mapping()`, `forMap()`.
- **`ExcelReader(Supplier)` / `CsvReader(Supplier)` no-validator constructors** —
  eliminates the `null` parameter: `new ExcelReader<>(User::new)` instead of
  `new ExcelReader<>(User::new, null)`.
- **`ExcelWriter.forMap(Builder, String...)` overload** — allows setting color, maxRows,
  rowAccessWindowSize on map-mode writers.
- **`ExcelReader.forMap(String...)` / `CsvReader.forMap(String...)` column selection** —
  include only specified headers in the output map.
- **`ExcelKitConfig` utility class** — centralized JVM-global POI configuration.
- **`ExcelColumn.of(name, fn, style, setter)` factory** — 4-param shortcut for the
  17-param constructor (mainly useful in tests).
- **`CellStyleParams.of()` convenience factories** — eliminates 11-null padding in
  common call patterns.

### Internal improvements

- **Javadoc warnings: 100 → 0** — all public APIs now have complete javadoc.
- **`ExcelWriteSupport` shared methods** — `writeAfterDataAndSummary()` and
  `applyPostProcessing()` extracted, eliminating duplication between ExcelWriter and
  ExcelSheetWriter. `writeRowCells()` now accepts `SheetConfig<T>` directly (9→7 params).
- **`TemplateListWriter` → `SheetConfig` delegation** — 8 duplicate fields replaced with
  a shared `SheetConfig<T>` instance.
- **`AbstractReadHandler` validation helpers** — `validateHeaderRowIndex()` and
  `validateColumns()` eliminate repeated if/throw blocks across 4 handler constructors.
- **Mutable array hack removed** — `ExcelSheetWriter.write()` uses an Iterator loop
  instead of `SXSSFSheet[]` array workaround.
- **Magic numbers extracted** — `WIDTH_PER_CHAR`, `WIDTH_BASE_PADDING`,
  `DEFAULT_FONT_SIZE`, `FONT_HEIGHT_MULTIPLIER`, `LUMINANCE_DARK_THRESHOLD`,
  `DEFAULT_ROW_HEIGHT_POINTS` are now named constants.
- **Defensive array copies** — `ExcelColumn.getDropdownOptions()` and
  `getHeaderFontColor()` return cloned arrays.
- **`ColumnStyleConfig` fields reorganized** — grouped by category (Layout, Font, Color,
  Borders, Validation, Protection, Grouping) with section comments.
- **Package-info expanded** — all three package descriptions upgraded from one-liners to
  meaningful summaries listing key classes and capabilities.
- **Error message consistency** — duplicate-name errors now use consistent `'name'`
  quoting format.
- **ExcelSheetWriter javadoc** — documents the `ColumnConfig` vs `ExcelColumnBuilder`
  design difference.

### Migration Guide

```java
// ─── Package rename ───
// Before
import io.github.dornol.excelkit.shared.FileHandler;
import io.github.dornol.excelkit.shared.CellData;
// After
import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.core.CellData;

// ─── RowFunction ───
// Before
ExcelRowFunction<User, Object> fn = (user, cursor) -> user.getName();
CsvRowFunction<User, Object> fn = (user, cursor) -> user.getName();
// After
RowFunction<User, Object> fn = (user, cursor) -> user.getName();

// ─── ExcelConsumer → WriteRowCallback ───
// Before
writer.write(stream, (ExcelConsumer<User>) (row, cursor) -> log(row));
// After
writer.write(stream, (WriteRowCallback<User>) (row, cursor) -> log(row));

// ─── ExcelWorkbook ───
// Before
new ExcelWorkbook(ExcelColor.STEEL_BLUE)
// After
ExcelWorkbook.create(opts -> opts.color(ExcelColor.STEEL_BLUE))

// ─── Reader setter mode ───
// Before
new ExcelReader<>(User::new, null)
// After (pick one)
new ExcelReader<>(User::new)          // no-validator constructor
ExcelReader.setter(User::new)         // static factory (symmetric with mapping/forMap)

// ─── Large file support ───
// Before
ExcelReader.configureLargeFileSupport();
// After
ExcelKitConfig.configureLargeFileSupport();
```

## [0.13.0] - 2026-04-12

### Changed (Breaking)

- **`ExcelSheetWriter.ColumnConfig` and `TemplateListWriter.ColumnConfig` inner classes deleted** —
  replaced by a single top-level `ColumnConfig<T>` class that extends `ColumnStyleConfig<T, ColumnConfig<T>>`.
  Both inner classes were empty wrappers (zero methods) inheriting all 47 styling methods from
  `ColumnStyleConfig`. The new class is identical in behavior; only the qualified name changes.
  `ExcelColumn.ExcelColumnBuilder` is unaffected (it has its own unique `style()` / `build()` methods).

### Migration Guide

```java
// Before (v0.12.0) — qualified references
ExcelSheetWriter.ColumnConfig<MyRow> config = new ExcelSheetWriter.ColumnConfig<>();
TemplateListWriter.ColumnConfig<MyRow> config2 = new TemplateListWriter.ColumnConfig<>();

// After (v0.13.0) — top-level class
ColumnConfig<MyRow> config = new ColumnConfig<>();

// Lambda-style callers are unchanged (type inferred):
sheetWriter.column("Name", row -> row.getName(), cfg -> cfg.bold(true).type(ExcelDataType.STRING));
```

## [0.12.0] - 2026-04-12

v0.12.0 completes the Map I/O symmetry that v0.11.0 deferred: the Map Reader
classes are absorbed into `ExcelReader.forMap()` / `CsvReader.forMap()` static
factories, matching the Writer side.

### Changed (Breaking)

- **`ExcelMapReader` class removed** — use `ExcelReader.forMap()`. The returned
  reader is an `ExcelReader<Map<String, String>>` with the full fluent API
  (`sheetIndex`, `headerRowIndex`, `onProgress`, `readAsStream`). Gains
  `onProgress` support that `ExcelMapReader` never had.
- **`CsvMapReader` class removed** — use `CsvReader.forMap()`. Same benefits:
  full `CsvReader` API including `dialect`, `delimiter`, `charset`,
  `headerRowIndex`, `onProgress`, `readAsStream`.
- **Mixed-mode runtime guard** — calling `column(setter)`, `column(name, setter)`,
  `columnAt(idx, setter)`, `skipColumn()`, or `skipColumns(int)` on a reader
  obtained from `forMap()` now throws `IllegalStateException`. Map mode
  auto-discovers columns; the setter API doesn't apply.

### Behavioral notes

- **`readAsStream` on a non-existent sheet** now throws `ExcelReadException`
  via `ExcelReadHandler`'s missing-sheet check. The deleted `ExcelMapReader`
  silently returned an empty stream, which hid caller bugs.
- Map building still uses positional pairing truncated at
  `min(headerCount, cellCount)` — matches the deleted Map Readers exactly.

### Migration Guide

```java
// ─── Excel map reading ───
// Before (v0.11.0)
new ExcelMapReader()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> process(r.data()));

// After (v0.12.0)
ExcelReader.forMap()
    .sheetIndex(0)
    .headerRowIndex(0)
    .build(inputStream)
    .read(r -> process(r.data()));

// ─── CSV map reading ───
// Before
new CsvMapReader()
    .dialect(CsvDialect.EXCEL)
    .onProgress(1000, (count, total) -> System.out.println(count))
    .build(inputStream)
    .read(r -> process(r.data()));

// After
CsvReader.forMap()
    .dialect(CsvDialect.EXCEL)
    .onProgress(1000, (count, total) -> System.out.println(count))
    .build(inputStream)
    .read(r -> process(r.data()));

// ─── Excel map reading now supports onProgress (new) ───
ExcelReader.forMap()
    .onProgress(1000, (count, total) -> System.out.println("read " + count))
    .build(inputStream)
    .read(r -> process(r.data()));
```

### Internal note

Absorption reuses the existing mapping-mode infrastructure
(`Function<RowData, T>`) via a synthetic `Function<RowData, Map<String, String>>`.
No SAX handler or `ExcelReadHandler` / `CsvReadHandler` changes were needed —
`RowData` already exposes `headerNames()` and `get(name)`, so the entire
"map reader" can be expressed as a 5-line row mapper. This eliminated the
"SAX callback state-machine rewrite" risk called out in the Plan.

## [0.11.0] - 2026-04-12

v0.11.0 is an **API cleanup release**. It removes a handful of parallel or
stale entry points that accumulated through v0.9.x–v0.10.0 and lands the
Reader-side half of the "unified column API" work that started in v0.10.0.
No new features.

### Changed (Breaking)

- **`ExcelWriter` constructors removed** — all 5 public constructors
  (`ExcelWriter()`, `ExcelWriter(maxRows)`, `ExcelWriter(color)`,
  `ExcelWriter(color, maxRows)`, `ExcelWriter(color, maxRows, windowSize)`)
  are deleted. Use `ExcelWriter.<T>builder()` instead.
- **`FileHandler` interface + `write()` rename** — `ExcelHandler` and
  `CsvHandler` now implement `shared.FileHandler` and expose
  `write(OutputStream)` instead of `consumeOutputStream(OutputStream)`.
  Both handler classes are now `final`. `ExcelHandler`'s
  `consumeOutputStreamWithPassword` Excel-only overloads are unchanged.
  `FileHandler` is a plain `interface` rather than `sealed` because
  excel-kit ships as an automatic module (no `module-info.java`); the
  `final` implementations preserve the closed-hierarchy intent and
  third-party implementations remain unsupported.
- **Reader column API unified** — `ExcelReader` / `CsvReader` no longer have
  `addColumn`, `columnAtBuilder`, `ExcelReadColumnBuilder`, or
  `CsvReadColumnBuilder`. `column(setter)` and `column(name, setter)` now
  return the reader itself (previously returned a chain-continuation
  builder). `columnAt(int, setter)` is unchanged.
- **`ExcelMapWriter` and `CsvMapWriter` deleted** — replaced by
  `ExcelWriter.forMap(...)` and `CsvWriter.forMap(...)` static factories
  that return the underlying writer. Use the writer's full fluent API
  directly instead of reaching through `.writer()` or a limited set of
  shortcut methods.

### Migration Guide

```java
// ─── ExcelWriter construction ───
// Before
new ExcelWriter<User>()
new ExcelWriter<User>(ExcelColor.STEEL_BLUE)
new ExcelWriter<User>(ExcelColor.STEEL_BLUE, 500_000, 500)
// After
ExcelWriter.<User>create()
ExcelWriter.<User>create(opts -> opts.color(ExcelColor.STEEL_BLUE))
ExcelWriter.<User>builder()
    .color(ExcelColor.STEEL_BLUE).maxRows(500_000).rowAccessWindowSize(500).build()

// ─── Writing the output ───
// Before
handler.consumeOutputStream(out)
// After
handler.write(out)

// ─── Reader column binding ───
// Before
reader.addColumn(User::setName)
reader.addColumn("Name", User::setName)
reader.column(User::setName)          // (returned a builder)
reader.columnAtBuilder(2, User::setAge)
// After
reader.column(User::setName)          // now returns Reader<T>
reader.column("Name", User::setName)
reader.columnAt(2, User::setAge)

// ─── Map writers ───
// Before
new ExcelMapWriter("Name", "Age").write(stream).consumeOutputStream(out)
new CsvMapWriter("Name", "Age").dialect(CsvDialect.EXCEL).write(stream).consumeOutputStream(out)
// After
ExcelWriter.forMap("Name", "Age").write(stream).write(out)
CsvWriter.forMap("Name", "Age").dialect(CsvDialect.EXCEL).write(stream).write(out)
```

### Deferred to v0.12.0

- Map Reader absorption — `ExcelMapReader` and `CsvMapReader` remain as
  standalone classes. Their header-auto-detect logic is woven into the
  SAX-style row callbacks, so folding them into `ExcelReader.forMap()` /
  `CsvReader.forMap()` needs a separate refactor and is scheduled for
  v0.12.0.

## [0.10.0] - 2026-04-09

### Changed (Breaking)
- **Unified column API** — `ExcelWriter` now uses the same `.column()` / `.columnIf()` / `.constColumn()`
  pattern as `ExcelSheetWriter` and `CsvWriter`. All column methods return `ExcelWriter<T>` for chaining.
  Column configuration is done via lambda configurer: `.column("Name", fn, cfg -> cfg.type(...).bold(true))`.
- **Removed builder-chaining style** — The old `ExcelColumnBuilder` navigation methods (`column()`,
  `columnIf()`, `constColumn()`, `write()`, `beforeHeader()`, `afterData()`, `onProgress()`) are removed.
  `ExcelColumnBuilder` is now only used internally for column configuration.
- **Renamed `addColumn` → `column`** on `ExcelWriter` for consistency across all writer APIs.

### Added
- **`columnIf` on `ExcelWriter`** — conditional column with all 4 overloads
  (Function, Function+Consumer, ExcelRowFunction, ExcelRowFunction+Consumer).
- **`constColumn` with configurer** — `.constColumn("name", value, cfg -> cfg.type(...))`.
- **`constColumnIf`** — conditional constant column.

### Migration Guide
```java
// Before (0.9.x)
writer.column("Price", Product::price).type(ExcelDataType.INTEGER).format("#,##0")
writer.addColumn("Price", Product::price, cfg -> cfg.type(ExcelDataType.INTEGER))

// After (0.10.0)
writer.column("Price", Product::price, cfg -> cfg.type(ExcelDataType.INTEGER).format("#,##0"))
```

## [0.9.6] - 2026-04-08

### Added
- **AI context documentation** in JAR (`META-INF/AI.md` + `META-INF/excel-kit/*.md`) —
  structured documentation for AI agents to discover library usage when exploring dependencies.
- **`llms.txt`** published to GitHub Pages for web-accessible AI context.
- **`CLAUDE.md`** with release checklist and project conventions.

## [0.9.5] - 2026-04-08

### Added
- **Per-column header font color** via `headerFontColor(ExcelColor)` / `headerFontColor(int, int, int)` —
  override the header font color for individual columns. Useful for conditionally highlighting
  specific column headers (e.g., error indicators). Available on both `ExcelWriter` (builder chaining)
  and `ExcelSheetWriter` (lambda config). Passes `null` to use the default header style.

## [0.9.4] - 2026-04-01

### Added
- **`ExcelWriter.password(String)`** / **`ExcelWorkbook.password(String)`** — set encryption password
  at the writer level. `write()` automatically encrypts without needing
  `consumeOutputStreamWithPassword()`. Consistent with `protectSheet()` / `protectWorkbook()` API pattern.
- **CsvMapReader** — read CSV files into `Map<String, String>` without typed POJOs,
  matching the `ExcelMapReader` API. Supports `dialect()`, `delimiter()`, `charset()`,
  `headerRowIndex()`, `onProgress()`, and `readAsStream()`.
- **CsvMapWriter** — `dialect()`, `delimiter()`, `charset()`, `bom()` configuration
  methods for symmetry with `CsvMapReader`.
- README: Quick Reference table for all read/write classes.
- Benchmark CI workflow — runs performance tests on every main-branch push.
- Read benchmarks — Excel MapReader/typed 100K rows, CSV MapReader/typed 1M rows.
- Multi-JDK CI — test matrix across JDK 17, 21, and 25.
- Javadoc GitHub Pages workflow — auto-deploy API docs on release.
- Example app: CSV Map Reader upload endpoint.

### Improved
- `ExcelHandler` internals refactored: `markConsumed()` and `writePlain()` extracted
  to eliminate duplication between encrypted and plain-text write paths.
- `char[]` password validation now rejects blank (whitespace-only) arrays,
  consistent with `String` password validation.
- Calling `consumeOutputStreamWithPassword()` when `password()` is already set throws
  `IllegalStateException` with a descriptive message instead of silently using
  the wrong password.
- Test coverage boost: 38 new targeted tests for CsvMapReader, CsvWriter quoting,
  AbstractReadHandler validation, and TempResourceContainer edge cases.

### Fixed
- `ExcelHandler.consumeOutputStreamWithPassword(char[])`: char array is now zeroed
  even when `IllegalStateException` is thrown due to password conflict.
  Previously the array was left intact if the exception occurred before the
  `try/finally` block.

### Dependencies
- poi-ooxml 5.4.1 → 5.5.1
- junit-bom 5.10.0 → 6.0.3
- actions/checkout v4 → v6
- actions/upload-artifact v4 → v7
- actions/setup-java v4 → v5
- Dependabot configured for Gradle and GitHub Actions.

## [0.9.2] - 2026-03-30

### Added
- **Data bar conditional formatting** via `dataBar(ExcelColor)` — gradient bars
  proportional to cell values. Supports single-color and 2-color gradient
  (`dataBar(minColor, maxColor)`).
- **Icon set conditional formatting** via `iconSet(IconSetType)` — 10 icon set
  types including arrows, traffic lights, flags, signs, symbols, ratings, quarters.
- **Timezone-aware date parsing** via `CellData.asZonedDateTime(ZoneId)` and
  `CellData.asZonedDateTime(String format, ZoneId)`.
- **CSV dialect presets** via `CsvDialect` enum — RFC4180, EXCEL, TSV, PIPE.
  Apply with `CsvWriter.dialect()` and `CsvReader.dialect()`.
- **CSV quoting strategies** via `CsvQuoting` enum — MINIMAL (default), ALL
  (quote everything), NON_NUMERIC (quote strings, leave numbers unquoted).
  Configure with `CsvWriter.quoting()`.
- README: Supported Formats table, Notes section (JVM-global config warning,
  readAsStream try-with-resources requirement).

### Changed
- **ExcelMapReader.readAsStream()**: Converted from List-collect approach to true
  streaming via BlockingQueue + producer thread (same pattern as ExcelReadHandler).
  Now memory-efficient for large datasets.

### Improved
- Branch test coverage: 84% → 89% (+82 new tests, +46 branches covered).
- Test assertion quality: replaced `assertTrue(out.size() > 0)` patterns with
  actual POI API content verification (validation rules, chart types, cell values,
  font styles, formula content).

## [0.9.0] - 2026-03-19

### Added
- **ExcelTemplateWriter** — fill data into existing .xlsx templates while preserving
  formatting, images, charts, and merged regions. Supports cell-level writes,
  list streaming via `TemplateListWriter`, and mixed mode (cell + list + afterData).
- **TemplateListWriter** — stream tabular data into a template sheet with column
  definition, afterData callbacks, summary rows, and all existing styling options.
- Write benchmarks (`WriteBenchmarkTest`) — 7 scenarios measuring Excel/CSV throughput.
- Migration guide in README for 0.8.1 → 0.8.2 breaking changes.
- Performance section in README with throughput table.

## [0.8.2] - 2026-03-18

### Changed
- **CellData default locale**: Changed from `Locale.KOREA` to `Locale.getDefault()`.
  Korean environments are unaffected (JVM default is already `Locale.KOREA`).
  Users who relied on the hard-coded Korean locale should call
  `CellData.setDefaultLocale(Locale.KOREA)` explicitly.
- **ExcelDataType.IMAGE**: Now throws `ExcelWriteException` when a non-`ExcelImage` value
  is passed to an IMAGE column. Previously it silently fell back to `String.valueOf()`.
- **RowData.get(int)**: Now throws `IllegalArgumentException` on negative index.
  Previously it silently mapped negative indices to 0.

### Added
- `SheetConfig<T>` — internal shared configuration class for `ExcelWriter` and
  `ExcelSheetWriter`, eliminating 17 duplicated field declarations.
- JaCoCo test coverage verification with 70% minimum threshold.
- CI workflow uploads JaCoCo report as artifact for PR visibility.
- `readAsStream()` Javadoc warns about try-with-resources requirement.
- `SXSSFSheetHelper` logs warning on reflection failure instead of silent null.
- `ExcelChartConfig.categoryColumn()` / `valueColumn()` reject negative indices.
- `CellData` date format add/reset methods are now synchronized (`FORMAT_LOCK`).
- `ExcelReadHandler` producer thread named `"excel-kit-reader"` for debuggability.
- `CsvReadHandler` has its own dedicated logger.

### Fixed
- Misplaced Javadoc blocks across 6 source files (caused by linter reordering).
- Silent exception swallowing in `ExcelWorkbook.close()` and `CsvReadHandler.closeQuietly()`.
- `assert` statements in production code replaced with `IllegalStateException`
  (`CsvReadHandler`, `ExcelReadHandler`, `AbstractReadHandler`).
- `ExcelHandler.consumeOutputStreamWithPassword(char[])` error message unified to "blank".
- `CsvReadHandler` BOM error message clarified.

## [0.8.1] - 2025-07-29

### Added
- `CellData.as(Function)` — custom type conversion (e.g., `UUID::fromString`).
- `CellData.as(Function, defaultValue)` — custom conversion with default.
- Default value overloads: `asInt(int)`, `asLong(long)`, `asDouble(double)`, `asString(String)`.
- `CsvWriter.csvInjectionDefense(boolean)` — toggle CSV injection defense.
- Round-trip integration tests for CellData conversions.

## [0.8.0] - 2025-07-28

### Added
- **Mapping mode** for immutable object / Java record reading via `ExcelReader.mapping()` and `CsvReader.mapping()`.
- `RowData` class for positional and named cell access.
- Map-based reading via `ExcelMapReader` and `CsvMapReader`.

## [0.7.2] - 2025-07-27

### Added
- Workbook protection via `protectWorkbook()`.
- Header font customization via `headerFontName()` and `headerFontSize()`.
- Default column style via `defaultStyle()`.
- Summary/footer rows via `summary()` — SUM, AVERAGE, COUNT, MIN, MAX.
- Named ranges via `SheetContext.namedRange()`.
- List validation from cell range via `ExcelValidation.listFromRange()`.
