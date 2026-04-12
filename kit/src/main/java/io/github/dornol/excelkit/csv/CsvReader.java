package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.core.ReadColumn;
import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * Builder-style class for configuring CSV row readers.
 * <p>
 * {@code CsvReader} allows you to define how each CSV cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, use {@link #build(InputStream)} to create a {@link CsvReadHandler}.
 *
 * @param <T> The type of the object that represents one CSV row
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReader<T> {
    private final List<ReadColumn<T>> columns = new ArrayList<>();
    private final @Nullable Supplier<T> instanceSupplier;
    private final @Nullable Function<RowData, T> rowMapper;
    private final @Nullable Validator validator;
    private int headerRowIndex = 0;
    private char delimiter = ',';
    private Charset charset = StandardCharsets.UTF_8;
    private @Nullable ProgressCallback progressCallback;
    private int progressInterval;
    private boolean mapMode = false;

    /**
     * Constructs a CsvReader in setter mode with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public CsvReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.rowMapper = null;
        this.validator = validator;
    }

    /**
     * Constructs a CsvReader in setter mode without Bean Validation.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     */
    public CsvReader(Supplier<T> instanceSupplier) {
        this(instanceSupplier, null);
    }

    /**
     * Creates a CsvReader in setter mode. Symmetric with {@link #mapping(Function)} and {@link #forMap()}.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param <T>              The row data type
     * @return A new CsvReader configured in setter mode
     * @since 0.14.0
     */
    public static <T> CsvReader<T> setter(Supplier<T> instanceSupplier) {
        return new CsvReader<>(instanceSupplier, null);
    }

    /**
     * Creates a CsvReader in setter mode with Bean Validation.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Bean Validation validator
     * @param <T>              The row data type
     * @return A new CsvReader configured in setter mode
     * @since 0.14.0
     */
    public static <T> CsvReader<T> setter(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        return new CsvReader<>(instanceSupplier, validator);
    }

    private CsvReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        this.instanceSupplier = null;
        this.rowMapper = Objects.requireNonNull(rowMapper, "rowMapper cannot be null");
        this.validator = validator;
    }

    /**
     * Creates a CsvReader in mapping mode for immutable object construction.
     * <p>
     * In this mode, each row is passed as a {@link RowData} to the mapping function,
     * which creates the target object in a single step.
     *
     * <pre>{@code
     * CsvReader.mapping(row -> new PersonRecord(
     *         row.get("Name").asString(),
     *         row.get("Age").asInt()
     * )).build(inputStream).read(result -> { ... });
     * }</pre>
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param <T>       The type of the object that represents one CSV row
     * @return A new CsvReader configured in mapping mode
     */
    public static <T> CsvReader<T> mapping(Function<RowData, T> rowMapper) {
        return new CsvReader<>(rowMapper, null);
    }

    /**
     * Creates a CsvReader in mapping mode with Bean Validation support.
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param validator Optional Bean Validation validator (nullable)
     * @param <T>       The type of the object that represents one CSV row
     * @return A new CsvReader configured in mapping mode
     * @see #mapping(Function)
     */
    public static <T> CsvReader<T> mapping(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        return new CsvReader<>(rowMapper, validator);
    }

    /**
     * Creates a reader that parses CSV files into {@code Map<String, String>} rows by
     * auto-discovering columns from the header row.
     * <p>
     * The returned reader exposes the standard fluent API ({@link #dialect(CsvDialect)},
     * {@link #delimiter(char)}, {@link #charset(Charset)}, {@link #headerRowIndex(int)},
     * {@link #onProgress(int, ProgressCallback)}) but rejects
     * {@link #column(BiConsumer)}, {@link #column(String, BiConsumer)},
     * {@link #columnAt(int, BiConsumer)}, {@link #skipColumn()}, and {@link #skipColumns(int)}
     * at runtime — map mode infers columns automatically from the header row and does not
     * use the setter API.
     *
     * <pre>{@code
     * CsvReader.forMap()
     *     .dialect(CsvDialect.EXCEL)
     *     .build(inputStream)
     *     .read(result -> {
     *         Map<String, String> row = result.data();
     *         String name = row.get("Name");
     *     });
     * }</pre>
     *
     * @return a new CsvReader in map mode
     * @since 0.12.0
     */
    public static CsvReader<Map<String, String>> forMap() {
        return forMap((Set<String>) null);
    }

    /**
     * Creates a reader that parses CSV files into {@code Map<String, String>} rows,
     * including only the specified columns. Columns not listed are ignored.
     *
     * <pre>{@code
     * CsvReader.forMap("Name", "Age")
     *     .build(inputStream)
     *     .read(result -> {
     *         // result.data() contains only "Name" and "Age" keys
     *     });
     * }</pre>
     *
     * @param columnNames the header names to include (others are filtered out)
     * @return a new CsvReader in map mode with column filtering
     * @since 0.14.0
     */
    public static CsvReader<Map<String, String>> forMap(String... columnNames) {
        return forMap(new LinkedHashSet<>(List.of(columnNames)));
    }

    private static CsvReader<Map<String, String>> forMap(@Nullable Set<String> selectedColumns) {
        Function<RowData, Map<String, String>> mapMapper = row -> {
            Map<String, String> map = new LinkedHashMap<>();
            List<String> headers = row.headerNames();
            int bound = Math.min(headers.size(), row.size());
            for (int i = 0; i < bound; i++) {
                String header = headers.get(i);
                if (header == null) continue;
                if (selectedColumns != null && !selectedColumns.contains(header)) continue;
                map.put(header, row.get(i).formattedValue());
            }
            return map;
        };
        CsvReader<Map<String, String>> reader = CsvReader.mapping(mapMapper);
        reader.mapMode = true;
        return reader;
    }

    private void requireNotMapMode(String method) {
        if (mapMode) {
            throw new IllegalStateException(
                    method + " cannot be called on a forMap() reader; "
                            + "map mode auto-discovers columns from the header row");
        }
    }

    /**
     * Applies a predefined CSV dialect configuration.
     * <p>
     * Sets the delimiter and charset in one call.
     * Individual settings can be overridden after calling this method.
     *
     * @param dialect the dialect to apply
     * @return This CsvReader instance for chaining
     * @since 0.9.2
     */
    public CsvReader<T> dialect(CsvDialect dialect) {
        this.delimiter = dialect.getDelimiter();
        this.charset = dialect.getCharset();
        return this;
    }

    /**
     * Sets the zero-based row index of the header row.
     * Rows before this index will be skipped during reading.
     * Defaults to 0 (the first row).
     *
     * @param headerRowIndex The zero-based index of the header row
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    /**
     * Sets the delimiter character used to separate fields.
     * Defaults to comma ({@code ','}).
     *
     * @param delimiter The delimiter character
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> delimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    /**
     * Sets the character encoding for reading the CSV file.
     * Defaults to {@link StandardCharsets#UTF_8}.
     *
     * @param charset The charset to use
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> charset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Adds a column mapping to the internal list.
     *
     * @param column A CSV column with setter logic
     */
    void addColumn(ReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Registers a positional column mapping. Columns are matched to the CSV in the order
     * they are registered (after {@link #headerRowIndex(int)} is accounted for).
     *
     * @param setter a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public CsvReader<T> column(BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(BiConsumer)");
        columns.add(new ReadColumn<>(setter));
        return this;
    }

    /**
     * Registers a name-based column mapping. The column is matched to the CSV column
     * whose header equals {@code headerName}.
     *
     * @param headerName the header name to match
     * @param setter     a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public CsvReader<T> column(String headerName, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(String, BiConsumer)");
        columns.add(new ReadColumn<>(headerName, setter));
        return this;
    }

    /**
     * Registers an index-based column mapping. The column is matched to the CSV column
     * at the given 0-based index.
     *
     * @param columnIndex 0-based column index
     * @param setter      a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public CsvReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        requireNotMapMode("columnAt(int, BiConsumer)");
        columns.add(new ReadColumn<>(null, columnIndex, setter));
        return this;
    }

    /**
     * Skips one column during reading by registering a no-op mapping.
     *
     * @return this reader for chaining
     */
    public CsvReader<T> skipColumn() {
        requireNotMapMode("skipColumn()");
        columns.add(new ReadColumn<>((instance, cellData) -> {}));
        return this;
    }

    /**
     * Skips the specified number of positional columns.
     *
     * @param count the number of columns to skip (must be non-negative)
     * @return this reader for chaining
     * @throws IllegalArgumentException if {@code count} is negative
     */
    public CsvReader<T> skipColumns(int count) {
        requireNotMapMode("skipColumns(int)");
        if (count < 0) {
            throw new IllegalArgumentException("skipColumns count must be non-negative");
        }
        for (int i = 0; i < count; i++) {
            columns.add(new ReadColumn<>((instance, cellData) -> {}));
        }
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows during reading.
     *
     * @param interval the number of rows between each callback invocation (must be positive)
     * @param callback the callback to invoke
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Finalizes the configuration and builds a {@link CsvReadHandler} for parsing the given CSV stream.
     *
     * @param inputStream The input stream of the CSV file
     * @return A handler to execute CSV parsing
     */
    public CsvReadHandler<T> build(InputStream inputStream) {
        if (rowMapper != null) {
            return new CsvReadHandler<>(inputStream, rowMapper, validator,
                    headerRowIndex, delimiter, charset, progressInterval, progressCallback);
        }
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator,
                headerRowIndex, delimiter, charset, progressInterval, progressCallback);
    }
}
