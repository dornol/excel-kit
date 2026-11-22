package io.github.dornol.excelkit.csv;


import com.opencsv.ICSVParser;
import io.github.dornol.excelkit.core.read.AbstractReader;
import io.github.dornol.excelkit.core.read.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.io.BufferedInputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.ReadSummary;
import io.github.dornol.excelkit.core.read.ReadSession;
import io.github.dornol.excelkit.core.detect.TabularFileDetector;
import io.github.dornol.excelkit.core.detect.TabularFileType;
import io.github.dornol.excelkit.core.RuntimeDependencyGuard;

/**
 * Builder-style class for configuring CSV row readers.
 * <p>
 * {@code CsvReader} allows you to define how each CSV cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, call {@code read} with an input source and row consumer.
 *
 * @param <T> The type of the object that represents one CSV row
 * @author dhkim
 * @since 0.1.3
 */
public class CsvReader<T> extends AbstractReader<T, CsvReader<T>> {
    private @Nullable Character configuredDelimiter;
    private @Nullable Charset configuredCharset;
    private char quoteChar = '"';
    private char escapeChar = ICSVParser.NULL_CHARACTER;
    private boolean strictQuotes = false;
    private boolean ignoreLeadingWhiteSpace = true;
    private int detectionSampleBytes = TabularFileDetector.DEFAULT_SAMPLE_BYTES;

    /**
     * Constructs a CsvReader in setter mode with instance supplier and optional validator.
     */
    public CsvReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        super(instanceSupplier, validator);
        RuntimeDependencyGuard.require("com.opencsv.CSVReader", "com.opencsv:opencsv", "CsvReader");
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
        RuntimeDependencyGuard.require("com.opencsv.CSVReader", "com.opencsv:opencsv", "CsvReader");
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
            int bound = headers.size();
            for (int i = 0; i < bound; i++) {
                String header = headers.get(i);
                if (header == null) continue;
                if (selectedColumns != null && !selectedColumns.contains(header)) continue;
                if (row.headerIndexOf(header) != i) continue;
                map.put(header, row.get(i).formattedValue());
            }
            return map;
        };
        CsvReader<Map<String, String>> reader = CsvReader.mapping(mapMapper);
        reader.enableMapMode(selectedColumns);
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
        this.configuredDelimiter = dialect.getDelimiter();
        this.configuredCharset = dialect.getCharset();
        return this;
    }

    /**
     * Sets the delimiter character used to separate fields.
     * Defaults to comma ({@code ','}).
     * An explicitly supplied delimiter takes precedence in detected reads.
     *
     * @param delimiter The delimiter character
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> delimiter(char delimiter) {
        this.configuredDelimiter = delimiter;
        return this;
    }

    /**
     * Sets the character encoding for reading the CSV file.
     * Defaults to {@link StandardCharsets#UTF_8}.
     * An explicitly supplied charset also applies to detected reads.
     *
     * @param charset The charset to use
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> charset(Charset charset) {
        this.configuredCharset = Objects.requireNonNull(charset, "charset cannot be null");
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
     * By default quotes are escaped by doubling them, and backslashes are literal.
     *
     * @since 0.19.0
     */
    public CsvReader<T> escapeChar(char escapeChar) {
        this.escapeChar = escapeChar;
        return this;
    }

    /**
     * Enables or disables OpenCSV strict quote parsing (default: disabled).
     * <p>
     * When enabled, text outside quotes is dropped ({@code "A"junk} reads as {@code A}). That
     * includes the header row, so every header cell must be quoted too; an unquoted header reads as
     * blank and the read fails with a {@link CsvReadException} that says so.
     *
     * @since 0.19.0
     */
    public CsvReader<T> strictQuotes(boolean strictQuotes) {
        this.strictQuotes = strictQuotes;
        return this;
    }

    /**
     * Controls whether leading whitespace before quoted values is ignored (default: ignored).
     * <p>
     * When enabled, {@code a, "b"} reads the second value as {@code b}. Whitespace around values
     * that are not quoted is always kept.
     * <p>
     * Turn it off only for files in which a space before a quote is meant literally. The parser then does not
     * treat that quote as the start of a quoted value but still lets it hide a following delimiter, and it drops
     * the closing quote: {@code 1, "x,y"} reads as {@code  "x,y}, which is neither what the file says nor the quoted
     * value.
     *
     * @since 0.19.0
     */
    public CsvReader<T> ignoreLeadingWhiteSpace(boolean ignoreLeadingWhiteSpace) {
        this.ignoreLeadingWhiteSpace = ignoreLeadingWhiteSpace;
        return this;
    }

    /** Sets the number of bytes sampled by {@link #readDetected} (default 8192). */
    public CsvReader<T> detectionSampleBytes(int sampleBytes) {
        if (sampleBytes < 64) throw new IllegalArgumentException("sampleBytes must be at least 64");
        this.detectionSampleBytes = sampleBytes;
        return this;
    }

    /**
     * Finalizes the configuration and builds a {@link CsvReadHandler} for parsing the given CSV stream.
     *
     * @param inputStream The input stream of the CSV file
     * @return A handler to execute CSV parsing
     */
    @Override
    protected CsvReadHandler<T> createHandler(InputStream inputStream) {
        requireInputStream(inputStream);
        inputStream = limitInput(inputStream);
        return new CsvReadHandler<>(inputStream, sessionConfig());
    }

    @Override
    protected CsvReadHandler<T> createHandler(Path path) {
        requirePath(path);
        return new CsvReadHandler<>(path, sessionConfig());
    }

    private CsvReadSessionConfig<T> sessionConfig() {
        char delimiter = configuredDelimiter == null ? ',' : configuredDelimiter;
        return sessionConfig(delimiter, configuredCharset == null ? StandardCharsets.UTF_8 : configuredCharset);
    }

    private CsvReadSessionConfig<T> sessionConfig(char sessionDelimiter, Charset sessionCharset) {
        requireDistinctCharacters(sessionDelimiter, quoteChar, escapeChar);
        return new CsvReadSessionConfig<>(rowMapper == null ? configuredColumns() : null,
                instanceSupplier, rowMapper, configuredValidator(), configuredHeaderRowIndex(),
                sessionDelimiter, sessionCharset,
                configuredSelectedMapColumns(), quoteChar,
                escapeChar, strictQuotes, ignoreLeadingWhiteSpace, snapshotReadOptions());
    }

    /**
     * The delimiter, the quote and the escape character have to be told apart, and none of them may be a line
     * break: otherwise the parser fails on its own terms, or reads something other than the file says. A quote
     * or escape character of NUL means "none".
     */
    private static void requireDistinctCharacters(char delimiter, char quote, char escape) {
        if (delimiter == '\0') throw new IllegalArgumentException("The delimiter cannot be the NUL character");
        for (var entry : new char[][]{{delimiter, 'd'}, {quote, 'q'}, {escape, 'e'}}) {
            if (entry[0] == '\n' || entry[0] == '\r') {
                String what = entry[1] == 'd' ? "delimiter" : entry[1] == 'q' ? "quote character" : "escape character";
                throw new IllegalArgumentException("The " + what + " cannot be a line break");
            }
        }
        if (quote != '\0' && delimiter == quote) {
            throw new IllegalArgumentException("The delimiter and the quote character must differ, both are '" + delimiter + "'");
        }
        if (escape != '\0' && delimiter == escape) {
            throw new IllegalArgumentException("The delimiter and the escape character must differ, both are '" + delimiter + "'");
        }
        if (quote != '\0' && escape != '\0' && quote == escape) {
            throw new IllegalArgumentException("The quote and the escape character must differ, both are '" + quote + "'");
        }
    }

    private CsvReadHandler<T> createDetectedHandler(InputStream input) {
        BufferedInputStream buffered = input instanceof BufferedInputStream existing
                ? existing : new BufferedInputStream(input);
        int detectionHeader = configuredDelimiter == null ? configuredHeaderRowIndex() : 0;
        var detection = TabularFileDetector.detectDetailed(buffered, detectionSampleBytes,
                configuredCharset, quoteChar, escapeChar, detectionHeader);
        if (detection.type() != TabularFileType.CSV || detection.charset() == null) {
            throw new CsvReadException("Expected CSV content but detected " + detection.type());
        }
        Character delimiter = configuredDelimiter == null ? detection.delimiter() : configuredDelimiter;
        return new CsvReadHandler<>(limitInput(buffered), sessionConfig(
                delimiter == null ? ',' : delimiter, detection.charset()));
    }

    /** Detects text charset and delimiter from a sample before reading without closing the caller stream. */
    public void readDetected(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executeInputStream(inputStream, this::createDetectedHandler, handler -> handler.read(consumer));
    }

    /**
     * Detects the charset and delimiter from a sample of a path, then reads it, passing every row's
     * {@link ReadResult} to the consumer. The file is not modified or deleted.
     */
    public void readDetected(Path path, Consumer<ReadResult<T>> consumer) {
        requirePath(path);
        requireCallback(consumer, "consumer");
        try (InputStream input = Files.newInputStream(path)) { readDetected(input, consumer); }
        catch (IOException e) { throw new CsvReadException("Failed to open CSV input", e); }
    }

    /** Reads detected CSV content from a source whose stream is owned by the reader. */
    public void readDetected(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireInputSource(source);
        try (InputStream input = source.openStream()) {
            readDetected(input, consumer);
        } catch (IOException e) {
            throw new CsvReadException("Failed to open CSV input", e);
        }
    }

    /**
     * Like {@code readDetected} for an input stream, and returns a {@link ReadSummary} with the row
     * counts. The stream is not closed.
     */
    public ReadSummary readDetectedWithSummary(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeWithInputStream(inputStream, this::createDetectedHandler,
                handler -> summarize(handler, consumer));
    }

    /**
     * Like {@code readDetected} for an {@link InputStreamSource}, and returns a {@link ReadSummary}
     * with the row counts. The source's stream is opened and closed here.
     */
    public ReadSummary readDetectedWithSummary(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireInputSource(source);
        try (InputStream input = source.openStream()) {
            return readDetectedWithSummary(input, consumer);
        } catch (IOException e) {
            throw new CsvReadException("Failed to open CSV input", e);
        }
    }

    /**
     * Like {@code readDetected} for a path, and returns a {@link ReadSummary} with the row counts.
     * The file is not modified or deleted.
     */
    public ReadSummary readDetectedWithSummary(Path path, Consumer<ReadResult<T>> consumer) {
        requirePath(path);
        requireCallback(consumer, "consumer");
        try (InputStream input = Files.newInputStream(path)) {
            return readDetectedWithSummary(input, consumer);
        } catch (IOException e) {
            throw new CsvReadException("Failed to open CSV input", e);
        }
    }

    @Override
    protected RuntimeException openFailure(IOException cause) {
        return new CsvReadException("Failed to open CSV input", cause);
    }
}
