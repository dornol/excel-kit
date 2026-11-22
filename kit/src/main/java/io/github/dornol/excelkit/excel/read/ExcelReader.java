package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.AbstractReader;
import io.github.dornol.excelkit.core.read.CellData;
import io.github.dornol.excelkit.core.read.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.ArrayList;
import java.util.function.Function;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.function.Supplier;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.RowError;
import io.github.dornol.excelkit.core.read.ReadSummary;
import io.github.dornol.excelkit.core.read.ReadReport;
import io.github.dornol.excelkit.core.read.ReadSession;
import io.github.dornol.excelkit.core.RuntimeDependencyGuard;
import org.apache.poi.openxml4j.opc.OPCPackage;

/**
 * Builder-style class for configuring Excel row readers.
 * <p>
 * {@code ExcelReader} allows you to define how each Excel cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, call {@code read} with an input source and row consumer.
 *
 * @param <T> The type of the object that represents one Excel row
 *
 * @author dhkim
 * @since 0.1.2
 */
public class ExcelReader<T> extends AbstractReader<T, ExcelReader<T>> {
    private int sheetIndex = 0;
    private int headerRows = 1;
    private boolean countRows = false;
    private @Nullable String password;

    /**
     * Constructs an ExcelReader in setter mode with instance supplier and optional validator.
     */
    public ExcelReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        super(instanceSupplier, validator);
        RuntimeDependencyGuard.require("org.apache.poi.xssf.eventusermodel.XSSFReader",
                "org.apache.poi:poi-ooxml", "ExcelReader");
    }

    /**
     * Constructs an ExcelReader in setter mode without Bean Validation.
     */
    public ExcelReader(Supplier<T> instanceSupplier) {
        this(instanceSupplier, null);
    }

    private ExcelReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        super(rowMapper, validator);
        RuntimeDependencyGuard.require("org.apache.poi.xssf.eventusermodel.XSSFReader",
                "org.apache.poi:poi-ooxml", "ExcelReader");
    }

    /**
     * Creates an ExcelReader in setter mode. Symmetric with {@link #mapping(Function)}
     * and {@link #forMap()} — all three read modes start with a static factory.
     *
     * <pre>{@code
     * ExcelReader.setter(User::new)
     *     .column("Name", (u, cell) -> u.name = cell.asString())
     *     .read(inputStream, result -> { ... });
     * }</pre>
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param <T>              The row data type
     * @return A new ExcelReader configured in setter mode
     * @since 0.14.0
     */
    public static <T> ExcelReader<T> setter(Supplier<T> instanceSupplier) {
        return new ExcelReader<>(instanceSupplier, null);
    }

    /**
     * Creates an ExcelReader in setter mode with Bean Validation.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Bean Validation validator
     * @param <T>              The row data type
     * @return A new ExcelReader configured in setter mode
     * @since 0.14.0
     */
    public static <T> ExcelReader<T> setter(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        return new ExcelReader<>(instanceSupplier, validator);
    }

    /**
     * Creates an ExcelReader in mapping mode for immutable object construction.
     * <p>
     * In this mode, each row is passed as a {@link RowData} to the mapping function,
     * which creates the target object in a single step. Column definitions are not needed —
     * access columns by header name or index within the mapping function.
     *
     * <pre>{@code
     * ExcelReader.mapping(row -> new PersonRecord(
     *         row.get("Name").asString(),
     *         row.get("Age").asInt(),
     *         row.get("Email").asString()
     * )).read(inputStream, result -> { ... });
     * }</pre>
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param <T>       The type of the object that represents one Excel row
     * @return A new ExcelReader configured in mapping mode
     */
    public static <T> ExcelReader<T> mapping(Function<RowData, T> rowMapper) {
        return new ExcelReader<>(rowMapper, null);
    }

    /**
     * Creates an ExcelReader in mapping mode with Bean Validation support.
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param validator Optional Bean Validation validator (nullable)
     * @param <T>       The type of the object that represents one Excel row
     * @return A new ExcelReader configured in mapping mode
     * @see #mapping(Function)
     */
    public static <T> ExcelReader<T> mapping(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        return new ExcelReader<>(rowMapper, validator);
    }

    /**
     * Creates a reader that parses Excel files into {@code Map<String, String>} rows by
     * auto-discovering columns from the header row.
     * <p>
     * The returned reader exposes the standard fluent API ({@link #sheetIndex(int)},
     * {@link #headerRowIndex(int)}, {@link #onProgress(int, ProgressCallback)}) but rejects
     * {@link #column(BiConsumer)}, {@link #column(String, BiConsumer)},
     * {@link #columnAt(int, BiConsumer)}, {@link #skipColumn()}, and {@link #skipColumns(int)}
     * at runtime — map mode infers columns automatically from the header row and does not
     * use the setter API.
     *
     * <pre>{@code
     * ExcelReader.forMap()
     *     .sheetIndex(0)
     *     .headerRowIndex(0)
     *     .read(inputStream, result -> {
     *         Map<String, String> row = result.data();
     *         String name = row.get("Name");
     *     });
     * }</pre>
     *
     * @return a new ExcelReader in map mode
     * @since 0.12.0
     */
    public static ExcelReader<Map<String, String>> forMap() {
        return forMap((Set<String>) null);
    }

    /**
     * Creates a reader that parses Excel files into {@code Map<String, String>} rows,
     * including only the specified columns. Columns not listed are ignored.
     *
     * <pre>{@code
     * ExcelReader.forMap("Name", "Age")
     *     .read(inputStream, result -> {
     *         // result.data() contains only "Name" and "Age" keys
     *     });
     * }</pre>
     *
     * @param columnNames the header names to include (others are filtered out)
     * @return a new ExcelReader in map mode with column filtering
     * @since 0.14.0
     */
    public static ExcelReader<Map<String, String>> forMap(String... columnNames) {
        return forMap(new LinkedHashSet<>(List.of(columnNames)));
    }

    /**
     * Like {@link #forMap()}, but numeric cells are returned as their stored number in plain
     * notation without trailing zeros (for example {@code 2.5} in a {@code 0}-formatted cell, {@code 0.25} in a
     * {@code 0%} cell, or {@code 123456789012345} in a General cell) instead of the displayed text.
     * Date cells become Excel serial numbers. Other cells return their text.
     *
     * @return a new ExcelReader in map mode that returns stored numbers
     */
    public static ExcelReader<Map<String, String>> forMapRaw() {
        return forMap(null, true);
    }

    /**
     * Like {@link #forMap(String...)}, but numeric cells are returned as stored numbers.
     *
     * @param columnNames the header names to include (others are filtered out)
     * @return a new ExcelReader in map mode that returns stored numbers
     * @see #forMapRaw()
     */
    public static ExcelReader<Map<String, String>> forMapRaw(String... columnNames) {
        return forMap(new LinkedHashSet<>(List.of(columnNames)), true);
    }

    private static ExcelReader<Map<String, String>> forMap(@Nullable Set<String> selectedColumns) {
        return forMap(selectedColumns, false);
    }

    private static ExcelReader<Map<String, String>> forMap(@Nullable Set<String> selectedColumns, boolean raw) {
        Function<RowData, Map<String, String>> mapMapper = row -> {
            Map<String, String> map = new LinkedHashMap<>();
            List<String> headers = row.headerNames();
            int bound = headers.size();
            for (int i = 0; i < bound; i++) {
                String header = headers.get(i);
                if (header == null) continue;
                if (selectedColumns != null && !selectedColumns.contains(header)) continue;
                if (row.headerIndexOf(header) != i) continue;
                CellData cell = row.get(i);
                var stored = raw ? cell.rawNumber() : null;
                map.put(header, stored != null ? stored.stripTrailingZeros().toPlainString() : cell.formattedValue());
            }
            return map;
        };
        ExcelReader<Map<String, String>> reader = ExcelReader.mapping(mapMapper);
        reader.enableMapMode(selectedColumns);
        return reader;
    }

    /**
     * Sets the zero-based sheet index to read from. Defaults to 0.
     */
    public ExcelReader<T> sheetIndex(int sheetIndex) {
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("sheetIndex must be non-negative");
        }
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * Sets how many rows make up the header block, anchored at {@link #headerRowIndex(int)}
     * as the last (column header) row. Previous rows are treated as group header rows.
     * <p>
     * For each column, the effective header name is the bottom-most non-blank value within the
     * header block. This accommodates Excel files whose column header cells are part of a
     * vertical merge with a group label above (common when the file was produced by
     * multi-level {@code group(...)} on the writer side).
     * <p>
     * Default: {@code 1} (single header row — existing behavior).
     *
     * @param headerRows total header row count (must be &gt;= 1)
     * @return this reader for chaining
     * @since 0.16.13
     */
    public ExcelReader<T> headerRows(int headerRows) {
        if (headerRows < 1) {
            throw new IllegalArgumentException("headerRows must be >= 1");
        }
        this.headerRows = headerRows;
        return this;
    }

    /**
     * Enables a pre-scan pass to count the total number of data rows before parsing.
     * <p>
     * When enabled, the reader performs a lightweight SAX scan of the sheet XML to count
     * {@code <row>} elements before the actual read. The total is then available via
     * {@link io.github.dornol.excelkit.core.Cursor#getTotalRows()} in the
     * {@link io.github.dornol.excelkit.core.ProgressCallback}.
     * <p>
     * This is useful for reporting percentage-based progress (e.g., via SSE):
     * <pre>{@code
     * ExcelReader.setter(MyDto::new)
     *     .column((dto, cell) -> dto.setName(cell.getString()))
     *     .countRows()
     *     .onProgress(500, (processed, cursor) -> {
     *         long total = cursor.getTotalRows();
     *         int percent = (int) (processed * 100 / total);
     *     })
     *     .read(inputStream, result -> { ... });
     * }</pre>
     * <p>
     * Note: the pre-scan adds a small overhead (typically 20–30% of total read time)
     * since it streams through the sheet XML without parsing cell values.
     *
     * @return this reader for chaining
     */
    public ExcelReader<T> countRows() {
        this.countRows = true;
        return this;
    }

    /**
     * Sets the password for reading encrypted Excel files.
     * <p>
     * If the file is encrypted with the "agile" encryption mode (as produced by
     * {@link io.github.dornol.excelkit.excel.ExcelWriter#password(String)}), this password will be used to decrypt
     * it before parsing.
     *
     * @param password the file password
     * @return this reader for chaining
     * @since 0.14.0
     */
    public ExcelReader<T> password(@Nullable String password) {
        this.password = password;
        return this;
    }

    /**
     * Finalizes the configuration and builds an {@link ExcelReadHandler} for parsing the given Excel stream.
     *
     * @param inputStream The input stream of the Excel file
     * @return A handler to execute Excel parsing
     */
    private ExcelReadHandler<T> createHandler(InputStream inputStream) {
        requireInputStream(inputStream);
        inputStream = limitInput(inputStream);
        return new ExcelReadHandler<>(inputStream, sessionConfig());
    }

    private ExcelReadHandler<T> createHandler(Path path) {
        requirePath(path);
        return new ExcelReadHandler<>(path, sessionConfig());
    }

    /** Opens a reusable execution session backed by a caller-owned stream. */
    public ReadSession<T> open(InputStream inputStream) {
        return createHandler(inputStream);
    }

    /** Opens a session whose path resource is owned by the session. */
    public ReadSession<T> open(Path path) {
        return createHandler(path);
    }

    /** Opens a session from a source; the source stream is closed after session initialization. */
    public ReadSession<T> open(InputStreamSource source) {
        requireInputSource(source);
        try (InputStream input = source.openStream()) {
            return createHandler(input);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    private ExcelReadSessionConfig<T> sessionConfig() {
        return new ExcelReadSessionConfig<>(rowMapper == null ? configuredColumns() : null,
                instanceSupplier, rowMapper, configuredValidator(), sheetIndex, configuredHeaderRowIndex(), headerRows,
                password, countRows,
                configuredSelectedMapColumns(), snapshotReadOptions());
    }

    /** Reads an input stream without closing it. */
    public void read(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executeInputStream(inputStream, this::createHandler, handler -> handler.read(consumer));
    }

    /** Reads all successful rows into a list and fails on the first mapping or validation error. */
    public List<T> readAll(InputStream inputStream) {
        List<T> rows = new ArrayList<>();
        readStrict(inputStream, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads all successful rows from a path into an immutable list, failing on the first mapping or
     * validation error. The file is not modified or deleted.
     */
    public List<T> readAll(Path path) {
        List<T> rows = new ArrayList<>();
        readStrict(path, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads all successful rows from an {@link InputStreamSource} into an immutable list, failing
     * on the first mapping or validation error. The source's stream is opened and closed here.
     */
    public List<T> readAll(InputStreamSource source) {
        List<T> rows = new ArrayList<>();
        readStrict(source, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads an input stream, passing every row's {@link ReadResult} to the consumer, and returns a
     * {@link ReadSummary} with the row counts. The stream is not closed.
     */
    public ReadSummary readWithSummary(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> summarize(handler, consumer));
    }

    /**
     * Reads a path, passing every row's {@link ReadResult} to the consumer, and returns a {@link
     * ReadSummary} with the row counts. The file is not modified or deleted.
     */
    public ReadSummary readWithSummary(Path path, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeWithPath(path, this::createHandler,
                handler -> summarize(handler, consumer));
    }

    /**
     * Reads an {@link InputStreamSource}, passing every row's {@link ReadResult} to the consumer,
     * and returns a {@link ReadSummary} with the row counts. The source's stream is opened and
     * closed here.
     */
    public ReadSummary readWithSummary(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeFromSource(source, input -> readWithSummary(input, consumer),
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    private ReadSummary summarize(ExcelReadHandler<T> handler, Consumer<ReadResult<T>> consumer) {
        return summarizeRead(
                handler::read, handler::wasStoppedEarly, consumer);
    }

    /** The input size this reader accepts, or -1 when it has no limit. */
    long maxInputBytes() {
        return snapshotReadOptions().limits().maxInputBytes();
    }

    boolean canSharePackage() {
        return password == null;
    }

    /** The password this reader opens an encrypted workbook with, or {@code null}. */
    @Nullable String password() {
        return password;
    }

    ReadSummary readSharedPackage(Path path, OPCPackage pkg, Consumer<ReadResult<T>> consumer) {
        ExcelReadHandler<T> handler = createHandler(path);
        return summarizeRead(callback -> handler.readSharedPackage(pkg, callback), handler::wasStoppedEarly, consumer);
    }

    /**
     * Reads an input stream without delivering rows, and returns a {@link ReadReport} with the row
     * counts and up to {@code maxCollectedErrors} failed rows. The stream is not closed.
     */
    public ReadReport readReport(InputStream inputStream, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> report(handler, maxCollectedErrors));
    }

    /**
     * Reads a path without delivering rows, and returns a {@link ReadReport} with the row counts
     * and up to {@code maxCollectedErrors} failed rows. The file is not modified or deleted.
     */
    public ReadReport readReport(Path path, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeWithPath(path, this::createHandler,
                handler -> report(handler, maxCollectedErrors));
    }

    /**
     * Reads an {@link InputStreamSource} without delivering rows, and returns a {@link ReadReport}
     * with the row counts and up to {@code maxCollectedErrors} failed rows. The source's stream is
     * opened and closed here.
     */
    public ReadReport readReport(InputStreamSource source, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeFromSource(source, input -> readReport(input, maxCollectedErrors),
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    private ReadReport report(ExcelReadHandler<T> handler, int maxCollectedErrors) {
        return collectReadReport(
                consumer -> handler.read(consumer), handler::wasStoppedEarly, maxCollectedErrors);
    }

    /**
     * Reads an input stream, passing each valid row to {@code onSuccess} and each failed row to
     * {@code onError}. The stream is not closed.
     */
    public void read(InputStream inputStream, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executeInputStream(inputStream, this::createHandler,
                handler -> handler.read(onSuccess, onError));
    }

    /**
     * Reads an input stream, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The stream is not closed.
     */
    public void readStrict(InputStream inputStream, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executeInputStream(inputStream, this::createHandler, handler -> handler.readStrict(consumer));
    }

    /**
     * Reads an input stream while the predicate returns true. Stopping early is normal, not an
     * error; the returned {@link ReadSummary} says whether it happened. The stream is not closed.
     */
    public ReadSummary readWhile(InputStream inputStream, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> readWhile(handler, predicate));
    }

    /** Reads directly from a caller-owned path without modifying or deleting it. */
    public void read(Path path, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executePath(path, this::createHandler, handler -> handler.read(consumer));
    }

    /**
     * Reads a path, passing each valid row to {@code onSuccess} and each failed row to {@code
     * onError}. The file is not modified or deleted.
     */
    public void read(Path path, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executePath(path, this::createHandler, handler -> handler.read(onSuccess, onError));
    }

    /**
     * Reads a path, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The file is not modified or deleted.
     */
    public void readStrict(Path path, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executePath(path, this::createHandler, handler -> handler.readStrict(consumer));
    }

    /**
     * Reads a path while the predicate returns true. Stopping early is normal, not an error; the
     * returned {@link ReadSummary} says whether it happened. The file is not modified or deleted.
     */
    public ReadSummary readWhile(Path path, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeWithPath(path, this::createHandler,
                handler -> readWhile(handler, predicate));
    }

    /** Opens and closes the source-owned stream. */
    public void read(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executeFromSource(source, input -> { read(input, consumer); return null; },
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    /**
     * Reads an {@link InputStreamSource}, passing each valid row to {@code onSuccess} and each
     * failed row to {@code onError}. The source's stream is opened and closed here.
     */
    public void read(InputStreamSource source, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executeFromSource(source, input -> { read(input, onSuccess, onError); return null; },
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    /**
     * Reads an {@link InputStreamSource}, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The source's stream is opened and closed here.
     */
    public void readStrict(InputStreamSource source, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executeFromSource(source, input -> { readStrict(input, consumer); return null; },
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    /**
     * Reads an {@link InputStreamSource} while the predicate returns true. Stopping early is
     * normal, not an error; the returned {@link ReadSummary} says whether it happened. The source's
     * stream is opened and closed here.
     */
    public ReadSummary readWhile(InputStreamSource source, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeFromSource(source, input -> readWhile(input, predicate),
                (message, error) -> new ExcelReadException("Failed to open Excel input", error));
    }

    private ReadSummary readWhile(ExcelReadHandler<T> handler, Predicate<ReadResult<T>> predicate) {
        return summarizeReadWhile(
                handler::readWhile, handler::wasStoppedEarly, predicate);
    }

    /**
     * Returns the list of sheet names and indices from an Excel file.
     * <p>
     * The whole stream is copied to a temporary file without any limit; for a file from an untrusted source use
     * {@link #getSheetNames(InputStream, ReadLimits)}.
     *
     * @param inputStream The input stream of the Excel file (will be consumed)
     * @return A list of {@link ExcelSheetInfo} records containing sheet names and indices
     */
    public static List<ExcelSheetInfo> getSheetNames(InputStream inputStream) {
        return ExcelMetadataReader.sheetNames(inputStream);
    }

    /** Like {@link #getSheetNames(InputStream)}, for a file. */
    public static List<ExcelSheetInfo> getSheetNames(Path path) {
        Objects.requireNonNull(path, "path cannot be null");
        return ExcelMetadataReader.sheetNames((InputStreamSource) () -> Files.newInputStream(path));
    }

    /** Like {@link #getSheetNames(InputStream)}, for a source that is opened (and closed) here. */
    public static List<ExcelSheetInfo> getSheetNames(InputStreamSource source) {
        Objects.requireNonNull(source, "source cannot be null");
        return ExcelMetadataReader.sheetNames(source);
    }

    /**
     * Returns the sheet names and indices of a password-protected Excel file.
     *
     * @param inputStream the input stream of the file (will be consumed)
     * @param password    the file's password
     * @return the sheets
     * @throws ExcelReadException if the password is wrong, or the file is not an encrypted workbook
     * @since 0.27.0
     */
    public static List<ExcelSheetInfo> getSheetNames(InputStream inputStream, String password) {
        return ExcelMetadataReader.sheetNames(inputStream, Objects.requireNonNull(password, "password cannot be null"));
    }

    /** Like {@link #getSheetNames(InputStream, String)}, for a file. @since 0.27.0 */
    public static List<ExcelSheetInfo> getSheetNames(Path path, String password) {
        Objects.requireNonNull(path, "path cannot be null");
        Objects.requireNonNull(password, "password cannot be null");
        return ExcelMetadataReader.sheetNames((InputStreamSource) () -> Files.newInputStream(path), password);
    }

    /** Like {@link #getSheetNames(InputStream, String)}, for a source that is opened (and closed) here. @since 0.27.0 */
    public static List<ExcelSheetInfo> getSheetNames(InputStreamSource source, String password) {
        Objects.requireNonNull(source, "source cannot be null");
        return ExcelMetadataReader.sheetNames(source, Objects.requireNonNull(password, "password cannot be null"));
    }

    /**
     * Returns the header names from a specific sheet.
     * <p>
     * The whole stream is copied to a temporary file without any limit; for a file from an untrusted source use
     * {@link #getSheetHeaders(InputStream, int, int, ReadLimits)}.
     *
     * @param inputStream    The input stream of the Excel file (will be consumed)
     * @param sheetIndex     The 0-based sheet index
     * @param headerRowIndex The 0-based header row index
     * @return A list of header names
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex) {
        return ExcelMetadataReader.headers(inputStream, sheetIndex, headerRowIndex);
    }

    /** Like {@link #getSheetHeaders(InputStream, int, int)}, for a file. */
    public static List<String> getSheetHeaders(Path path, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(path, "path cannot be null");
        return ExcelMetadataReader.headers((InputStreamSource) () -> Files.newInputStream(path), sheetIndex, headerRowIndex);
    }

    /**
     * Like {@link #getSheetHeaders(InputStream, int, int)}, for a source that is opened (and closed)
     * here.
     */
    public static List<String> getSheetHeaders(InputStreamSource source, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(source, "source cannot be null");
        return ExcelMetadataReader.headers(source, sheetIndex, headerRowIndex);
    }

    /**
     * Returns the header names from a specific sheet of a password-protected Excel file.
     *
     * @param inputStream    the input stream of the file (will be consumed)
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @param password       the file's password
     * @return the header names
     * @throws ExcelReadException if the password is wrong, or the file is not an encrypted workbook
     * @since 0.27.0
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex,
                                               String password) {
        return ExcelMetadataReader.headers(inputStream, sheetIndex, headerRowIndex,
                Objects.requireNonNull(password, "password cannot be null"));
    }

    /** Like {@link #getSheetHeaders(InputStream, int, int, String)}, for a file. @since 0.27.0 */
    public static List<String> getSheetHeaders(Path path, int sheetIndex, int headerRowIndex, String password) {
        Objects.requireNonNull(path, "path cannot be null");
        Objects.requireNonNull(password, "password cannot be null");
        return ExcelMetadataReader.headers((InputStreamSource) () -> Files.newInputStream(path),
                sheetIndex, headerRowIndex, password);
    }

    /** Like {@link #getSheetHeaders(InputStream, int, int, String)}, for a source opened (and closed) here. @since 0.27.0 */
    public static List<String> getSheetHeaders(InputStreamSource source, int sheetIndex, int headerRowIndex,
                                               String password) {
        Objects.requireNonNull(source, "source cannot be null");
        return ExcelMetadataReader.headers(source, sheetIndex, headerRowIndex,
                Objects.requireNonNull(password, "password cannot be null"));
    }

    /**
     * Returns the sheet names and indices of an Excel file that comes from an untrusted source, such as an
     * upload.
     * <p>
     * The stream is copied to a temporary file first. The copy stops once it passes
     * {@link ReadLimits#maxInputBytes()}, and a workbook with more sheets than {@link ReadLimits#maxSheets()}
     * is refused. The other limits do not apply to this call.
     *
     * @param inputStream the input stream of the file (will be consumed)
     * @param limits      the limits to enforce, for example {@link ReadLimits#UPLOAD}
     * @return the sheets
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the number of sheets passes a limit
     * @since 0.28.0
     */
    public static List<ExcelSheetInfo> getSheetNames(InputStream inputStream, ReadLimits limits) {
        return ExcelMetadataReader.sheetNames(inputStream, null,
                Objects.requireNonNull(limits, "limits cannot be null"));
    }

    /**
     * Like {@link #getSheetNames(InputStream, ReadLimits)}, for a password-protected file.
     *
     * @param inputStream the input stream of the file (will be consumed)
     * @param password    the file's password
     * @param limits      the limits to enforce, for example {@link ReadLimits#UPLOAD}
     * @return the sheets
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the number of sheets passes a limit
     * @throws ExcelReadException         if the password is wrong, or the file is not an encrypted workbook
     * @since 0.28.0
     */
    public static List<ExcelSheetInfo> getSheetNames(InputStream inputStream, String password, ReadLimits limits) {
        return ExcelMetadataReader.sheetNames(inputStream,
                Objects.requireNonNull(password, "password cannot be null"),
                Objects.requireNonNull(limits, "limits cannot be null"));
    }

    /**
     * Returns the header names of a sheet of an Excel file that comes from an untrusted source, such as an
     * upload.
     * <p>
     * The stream is copied to a temporary file first. The copy stops once it passes
     * {@link ReadLimits#maxInputBytes()}, and a header row with more cells than
     * {@link ReadLimits#maxColumns()} is refused. The other limits do not apply to this call.
     *
     * @param inputStream    the input stream of the file (will be consumed)
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @param limits         the limits to enforce, for example {@link ReadLimits#UPLOAD}
     * @return the header names
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the header row passes a limit
     * @since 0.28.0
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex,
                                               ReadLimits limits) {
        return ExcelMetadataReader.headers(inputStream, sheetIndex, headerRowIndex, null,
                Objects.requireNonNull(limits, "limits cannot be null"));
    }

    /**
     * Like {@link #getSheetHeaders(InputStream, int, int, ReadLimits)}, for a password-protected file.
     *
     * @param inputStream    the input stream of the file (will be consumed)
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @param password       the file's password
     * @param limits         the limits to enforce, for example {@link ReadLimits#UPLOAD}
     * @return the header names
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the header row passes a limit
     * @throws ExcelReadException         if the password is wrong, or the file is not an encrypted workbook
     * @since 0.28.0
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex,
                                               String password, ReadLimits limits) {
        return ExcelMetadataReader.headers(inputStream, sheetIndex, headerRowIndex,
                Objects.requireNonNull(password, "password cannot be null"),
                Objects.requireNonNull(limits, "limits cannot be null"));
    }

}
