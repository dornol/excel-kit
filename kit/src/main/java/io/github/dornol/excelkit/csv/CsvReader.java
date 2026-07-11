package io.github.dornol.excelkit.csv;


import com.opencsv.ICSVParser;
import io.github.dornol.excelkit.core.AbstractReader;
import io.github.dornol.excelkit.core.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import io.github.dornol.excelkit.core.InputStreamSource;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowError;
import io.github.dornol.excelkit.core.ReadSummary;
import io.github.dornol.excelkit.core.ReadReport;

/**
 * Builder-style class for configuring CSV row readers.
 * <p>
 * {@code CsvReader} allows you to define how each CSV cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, call {@code read} with an input source and row consumer.
 *
 * @param <T> The type of the object that represents one CSV row
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReader<T> extends AbstractReader<T, CsvReader<T>> {
    private char delimiter = ',';
    private Charset charset = StandardCharsets.UTF_8;
    private char quoteChar = ICSVParser.DEFAULT_QUOTE_CHARACTER;
    private char escapeChar = ICSVParser.DEFAULT_ESCAPE_CHARACTER;
    private boolean strictQuotes = ICSVParser.DEFAULT_STRICT_QUOTES;
    private boolean ignoreLeadingWhiteSpace = ICSVParser.DEFAULT_IGNORE_LEADING_WHITESPACE;

    /**
     * Constructs a CsvReader in setter mode with instance supplier and optional validator.
     */
    public CsvReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        super(instanceSupplier, validator);
    }

    /**
     * Constructs a CsvReader in setter mode without Bean Validation.
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
        super(rowMapper, validator);
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
     * )).read(inputStream, result -> { ... });
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
     *     .read(inputStream, result -> {
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
     *     .read(inputStream, result -> {
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
                if (row.headerIndexOf(header) != i) continue;
                if (selectedColumns != null && !selectedColumns.contains(header)) continue;
                map.put(header, row.get(i).formattedValue());
            }
            return map;
        };
        CsvReader<Map<String, String>> reader = CsvReader.mapping(mapMapper);
        reader.mapMode = true;
        reader.selectedMapColumns(selectedColumns);
        return reader;
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
     * Sets the quote character used by the CSV parser.
     *
     * @since 0.19.0
     */
    public CsvReader<T> quoteChar(char quoteChar) {
        this.quoteChar = quoteChar;
        return this;
    }

    /**
     * Sets the escape character used by the CSV parser.
     *
     * @since 0.19.0
     */
    public CsvReader<T> escapeChar(char escapeChar) {
        this.escapeChar = escapeChar;
        return this;
    }

    /**
     * Enables or disables OpenCSV strict quote parsing.
     *
     * @since 0.19.0
     */
    public CsvReader<T> strictQuotes(boolean strictQuotes) {
        this.strictQuotes = strictQuotes;
        return this;
    }

    /**
     * Controls whether leading whitespace before quoted values is ignored.
     *
     * @since 0.19.0
     */
    public CsvReader<T> ignoreLeadingWhiteSpace(boolean ignoreLeadingWhiteSpace) {
        this.ignoreLeadingWhiteSpace = ignoreLeadingWhiteSpace;
        return this;
    }

    /**
     * Finalizes the configuration and builds a {@link CsvReadHandler} for parsing the given CSV stream.
     *
     * @param inputStream The input stream of the CSV file
     * @return A handler to execute CSV parsing
     */
    private CsvReadHandler<T> createHandler(InputStream inputStream) {
        CsvReadHandler<T> handler;
        if (rowMapper != null) {
            handler = new CsvReadHandler<>(inputStream, rowMapper, validator,
                    headerRowIndex, delimiter, charset, progressInterval, progressCallback,
                    strictHeaders, duplicateHeaderPolicy,
                    selectedMapColumns == null ? null : Set.copyOf(selectedMapColumns), cellConversionConfig,
                    quoteChar, escapeChar, strictQuotes, ignoreLeadingWhiteSpace,
                    maxRows, skipBlankRows, stopAtBlankRows);
        } else {
            handler = new CsvReadHandler<>(inputStream, List.copyOf(columns), instanceSupplier, validator,
                    headerRowIndex, delimiter, charset, progressInterval, progressCallback,
                    strictHeaders, duplicateHeaderPolicy, cellConversionConfig,
                    quoteChar, escapeChar, strictQuotes, ignoreLeadingWhiteSpace,
                    maxRows, skipBlankRows, stopAtBlankRows);
        }
        handler.options(snapshotReadOptions());
        return handler;
    }

    private CsvReadHandler<T> createHandler(Path path) {
        CsvReadHandler<T> handler = rowMapper != null
                ? CsvReadHandler.forPath(path, rowMapper, validator, headerRowIndex, delimiter, charset,
                    progressInterval, progressCallback, strictHeaders, duplicateHeaderPolicy,
                    selectedMapColumns == null ? null : Set.copyOf(selectedMapColumns), cellConversionConfig,
                    quoteChar, escapeChar, strictQuotes, ignoreLeadingWhiteSpace,
                    maxRows, skipBlankRows, stopAtBlankRows)
                : CsvReadHandler.forPath(path, List.copyOf(columns), instanceSupplier, validator,
                    headerRowIndex, delimiter, charset, progressInterval, progressCallback, strictHeaders,
                    duplicateHeaderPolicy, cellConversionConfig, quoteChar, escapeChar, strictQuotes,
                    ignoreLeadingWhiteSpace, maxRows, skipBlankRows, stopAtBlankRows);
        handler.options(snapshotReadOptions());
        return handler;
    }

    /** Reads an input stream without closing it. */
    public void read(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        createHandler(inputStream).read(consumer);
    }

    public ReadSummary readWithSummary(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        CsvReadHandler<T> handler = createHandler(inputStream);
        handler.read(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            consumer.accept(result);
        });
        return new ReadSummary(counts[0], counts[1], counts[2], handler.wasStoppedEarly(),
                java.time.Duration.ofNanos(System.nanoTime() - started));
    }

    public ReadReport readReport(InputStream inputStream, int maxCollectedErrors) {
        if (maxCollectedErrors < 0) throw new IllegalArgumentException("maxCollectedErrors must be non-negative");
        List<RowError> errors = new java.util.ArrayList<>();
        long[] row = {0};
        ReadSummary summary = readWithSummary(inputStream, result -> {
            row[0]++;
            if (!result.success() && errors.size() < maxCollectedErrors) {
                errors.add(new RowError(row[0], result.fileRowNum(),
                        result.cause() == null ? RowError.Type.VALIDATION : RowError.Type.MAPPING,
                        result.messages() == null ? List.of() : result.messages(), result.cause(),
                        result.cellErrors(), result.rawValues()));
            }
        });
        return new ReadReport(summary, errors, summary.errorRows() > errors.size());
    }

    public void read(InputStream inputStream, Consumer<T> onSuccess, Consumer<RowError> onError) {
        createHandler(inputStream).read(onSuccess, onError);
    }

    public void readStrict(InputStream inputStream, Consumer<T> consumer) {
        createHandler(inputStream).readStrict(consumer);
    }

    public void readWhile(InputStream inputStream, Predicate<ReadResult<T>> predicate) {
        createHandler(inputStream).readWhile(predicate);
    }

    /** Reads directly from a caller-owned path without modifying or deleting it. */
    public void read(Path path, Consumer<ReadResult<T>> consumer) {
        createHandler(path).read(consumer);
    }

    public void read(Path path, Consumer<T> onSuccess, Consumer<RowError> onError) {
        createHandler(path).read(onSuccess, onError);
    }

    public void readStrict(Path path, Consumer<T> consumer) {
        createHandler(path).readStrict(consumer);
    }

    public void readWhile(Path path, Predicate<ReadResult<T>> predicate) {
        createHandler(path).readWhile(predicate);
    }

    public void read(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        withSource(source, input -> read(input, consumer));
    }

    public void read(InputStreamSource source, Consumer<T> onSuccess, Consumer<RowError> onError) {
        withSource(source, input -> read(input, onSuccess, onError));
    }

    public void readStrict(InputStreamSource source, Consumer<T> consumer) {
        withSource(source, input -> readStrict(input, consumer));
    }

    public void readWhile(InputStreamSource source, Predicate<ReadResult<T>> predicate) {
        withSource(source, input -> readWhile(input, predicate));
    }

    private void withSource(InputStreamSource source, Consumer<InputStream> operation) {
        java.util.Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            operation.accept(input);
        } catch (IOException e) {
            throw new CsvReadException("Failed to open CSV input", e);
        }
    }
}
