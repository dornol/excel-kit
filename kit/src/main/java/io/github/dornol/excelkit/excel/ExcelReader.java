package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.AbstractReader;
import io.github.dornol.excelkit.core.RowData;
import io.github.dornol.excelkit.core.TempResourceCreator;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.jspecify.annotations.Nullable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.function.Function;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.function.Supplier;
import io.github.dornol.excelkit.core.InputStreamSource;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowError;
import io.github.dornol.excelkit.core.ReadSummary;
import io.github.dornol.excelkit.core.ReadReport;

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
 * @since 2025-07-19
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
    }

    /**
     * Constructs an ExcelReader in setter mode without Bean Validation.
     */
    public ExcelReader(Supplier<T> instanceSupplier) {
        this(instanceSupplier, null);
    }

    private ExcelReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        super(rowMapper, validator);
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

    private static ExcelReader<Map<String, String>> forMap(@Nullable Set<String> selectedColumns) {
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
        ExcelReader<Map<String, String>> reader = ExcelReader.mapping(mapMapper);
        reader.mapMode = true;
        reader.selectedMapColumns(selectedColumns);
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
     * {@link ExcelWriter#password(String)}), this password will be used to decrypt
     * it before parsing.
     *
     * @param password the file password
     * @return this reader for chaining
     * @since 0.14.0
     */
    public ExcelReader<T> password(String password) {
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
        inputStream = limitInput(inputStream);
        ExcelReadHandler<T> handler;
        if (rowMapper != null) {
            handler = new ExcelReadHandler<>(inputStream, rowMapper, validator,
                    sheetIndex, headerRowIndex, headerRows, progressInterval, progressCallback, password, countRows,
                    strictHeaders, duplicateHeaderPolicy,
                    selectedMapColumns == null ? null : Set.copyOf(selectedMapColumns), cellConversionConfig,
                    maxRows, skipBlankRows, stopAtBlankRows);
        } else {
            handler = new ExcelReadHandler<>(inputStream, List.copyOf(columns), instanceSupplier, validator,
                    sheetIndex, headerRowIndex, headerRows, progressInterval, progressCallback, password, countRows,
                    strictHeaders, duplicateHeaderPolicy, cellConversionConfig,
                    maxRows, skipBlankRows, stopAtBlankRows);
        }
        handler.options(snapshotReadOptions());
        return handler;
    }

    private ExcelReadHandler<T> createHandler(Path path) {
        ExcelReadHandler<T> handler = rowMapper != null
                ? ExcelReadHandler.forPath(path, rowMapper, validator, sheetIndex, headerRowIndex, headerRows,
                    progressInterval, progressCallback, password, countRows, strictHeaders,
                    duplicateHeaderPolicy, selectedMapColumns == null ? null : Set.copyOf(selectedMapColumns),
                    cellConversionConfig, maxRows, skipBlankRows, stopAtBlankRows)
                : ExcelReadHandler.forPath(path, List.copyOf(columns), instanceSupplier, validator, sheetIndex,
                    headerRowIndex, headerRows, progressInterval, progressCallback, password, countRows,
                    strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows, skipBlankRows,
                    stopAtBlankRows);
        handler.options(snapshotReadOptions());
        return handler;
    }

    /** Reads an input stream without closing it. */
    public void read(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        createHandler(inputStream).read(consumer);
    }

    public ReadSummary readWithSummary(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        return summarize(createHandler(inputStream), consumer);
    }

    public ReadSummary readWithSummary(Path path, Consumer<ReadResult<T>> consumer) {
        return summarize(createHandler(path), consumer);
    }

    public ReadSummary readWithSummary(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        final ReadSummary[] summary = new ReadSummary[1];
        withSource(source, input -> summary[0] = readWithSummary(input, consumer));
        return summary[0];
    }

    private ReadSummary summarize(ExcelReadHandler<T> handler, Consumer<ReadResult<T>> consumer) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        handler.read(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            consumer.accept(result);
        });
        return new ReadSummary(counts[0], counts[1], counts[2], handler.wasStoppedEarly(),
                java.time.Duration.ofNanos(System.nanoTime() - started));
    }

    public ReadReport readReport(InputStream inputStream, int maxCollectedErrors) {
        return report(createHandler(inputStream), maxCollectedErrors);
    }

    public ReadReport readReport(Path path, int maxCollectedErrors) {
        return report(createHandler(path), maxCollectedErrors);
    }

    public ReadReport readReport(InputStreamSource source, int maxCollectedErrors) {
        final ReadReport[] report = new ReadReport[1];
        withSource(source, input -> report[0] = readReport(input, maxCollectedErrors));
        return report[0];
    }

    private ReadReport report(ExcelReadHandler<T> handler, int maxCollectedErrors) {
        if (maxCollectedErrors < 0) throw new IllegalArgumentException("maxCollectedErrors must be non-negative");
        List<RowError> errors = new ArrayList<>();
        long[] row = {0};
        ReadSummary summary = summarize(handler, result -> {
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

    public ReadSummary readWhile(InputStream inputStream, Predicate<ReadResult<T>> predicate) {
        return readWhile(createHandler(inputStream), predicate);
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

    public ReadSummary readWhile(Path path, Predicate<ReadResult<T>> predicate) {
        return readWhile(createHandler(path), predicate);
    }

    /** Opens and closes the source-owned stream. */
    public void read(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        withSource(source, input -> read(input, consumer));
    }

    public void read(InputStreamSource source, Consumer<T> onSuccess, Consumer<RowError> onError) {
        withSource(source, input -> read(input, onSuccess, onError));
    }

    public void readStrict(InputStreamSource source, Consumer<T> consumer) {
        withSource(source, input -> readStrict(input, consumer));
    }

    public ReadSummary readWhile(InputStreamSource source, Predicate<ReadResult<T>> predicate) {
        final ReadSummary[] summary = new ReadSummary[1];
        withSource(source, input -> summary[0] = readWhile(input, predicate));
        return summary[0];
    }

    private ReadSummary readWhile(ExcelReadHandler<T> handler, Predicate<ReadResult<T>> predicate) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        handler.readWhile(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            return predicate.test(result);
        });
        return new ReadSummary(counts[0], counts[1], counts[2], handler.wasStoppedEarly(),
                java.time.Duration.ofNanos(System.nanoTime() - started));
    }

    private void withSource(InputStreamSource source, Consumer<InputStream> operation) {
        java.util.Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            operation.accept(input);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    /**
     * Returns the list of sheet names and indices from an Excel file.
     *
     * @param inputStream The input stream of the Excel file (will be consumed)
     * @return A list of {@link ExcelSheetInfo} records containing sheet names and indices
     */
    public static List<ExcelSheetInfo> getSheetNames(InputStream inputStream) {
        Path tempDir = null;
        Path tempFile = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            Files.copy(java.util.Objects.requireNonNull(inputStream, "inputStream cannot be null"),
                    tempFile, StandardCopyOption.REPLACE_EXISTING);

            try (OPCPackage pkg = OPCPackage.open(tempFile.toFile())) {
                XSSFReader reader = new XSSFReader(pkg);
                XSSFReader.SheetIterator sheetsData = (XSSFReader.SheetIterator) reader.getSheetsData();
                List<ExcelSheetInfo> result = new ArrayList<>();
                int index = 0;
                while (sheetsData.hasNext()) {
                    try (InputStream ignored = sheetsData.next()) {
                        result.add(new ExcelSheetInfo(index, sheetsData.getSheetName()));
                    }
                    index++;
                }
                return result;
            }
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read sheet names", e);
        } finally {
            cleanupTemp(tempDir, tempFile);
        }
    }

    public static List<ExcelSheetInfo> getSheetNames(Path path) {
        java.util.Objects.requireNonNull(path, "path cannot be null");
        return getSheetNames((InputStreamSource) () -> Files.newInputStream(path));
    }

    public static List<ExcelSheetInfo> getSheetNames(InputStreamSource source) {
        java.util.Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return getSheetNames(input);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    /**
     * Returns the header names from a specific sheet.
     *
     * @param inputStream    The input stream of the Excel file (will be consumed)
     * @param sheetIndex     The 0-based sheet index
     * @param headerRowIndex The 0-based header row index
     * @return A list of header names
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex) {
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("sheetIndex must be non-negative");
        }
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
        Path tempDir = null;
        Path tempFile = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            Files.copy(java.util.Objects.requireNonNull(inputStream, "inputStream cannot be null"),
                    tempFile, StandardCopyOption.REPLACE_EXISTING);

            try (OPCPackage pkg = OPCPackage.open(tempFile.toFile())) {
                XSSFReader reader = new XSSFReader(pkg);
                SharedStrings ss = reader.getSharedStringsTable();
                StylesTable styles = reader.getStylesTable();

                XMLReader parser = XMLHelper.newXMLReader();
                HeaderExtractor extractor = new HeaderExtractor(headerRowIndex);
                XSSFSheetXMLHandler sheetParser = new XSSFSheetXMLHandler(styles, ss, extractor, false);
                parser.setContentHandler(sheetParser);

                Iterator<InputStream> sheetsData = reader.getSheetsData();
                int currentIndex = 0;
                boolean found = false;
                while (sheetsData.hasNext()) {
                    try (InputStream sheet = sheetsData.next()) {
                        if (currentIndex == sheetIndex) {
                            parser.parse(new InputSource(sheet));
                            found = true;
                            break;
                        }
                    }
                    currentIndex++;
                }
                if (!found) {
                    throw new ExcelReadException("Sheet index " + sheetIndex + " not found. File has "
                            + currentIndex + " sheet(s).");
                }
                return extractor.getHeaders();
            }
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read sheet headers", e);
        } finally {
            cleanupTemp(tempDir, tempFile);
        }
    }

    public static List<String> getSheetHeaders(Path path, int sheetIndex, int headerRowIndex) {
        java.util.Objects.requireNonNull(path, "path cannot be null");
        return getSheetHeaders((InputStreamSource) () -> Files.newInputStream(path), sheetIndex, headerRowIndex);
    }

    public static List<String> getSheetHeaders(InputStreamSource source, int sheetIndex, int headerRowIndex) {
        java.util.Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return getSheetHeaders(input, sheetIndex, headerRowIndex);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(ExcelReader.class);

    private static void cleanupTemp(Path tempDir, Path tempFile) {
        if (tempFile != null) {
            try {
                Files.deleteIfExists(tempFile);
            } catch (IOException e) {
                log.warn("Failed to delete temp file: {}", tempFile, e);
                tempFile.toFile().deleteOnExit();
            }
        }
        if (tempDir != null) {
            try {
                Files.deleteIfExists(tempDir);
            } catch (IOException e) {
                log.warn("Failed to delete temp dir: {}", tempDir, e);
                tempDir.toFile().deleteOnExit();
            }
        }
    }

    /**
     * Internal handler for extracting header names only.
     */
    private static class HeaderExtractor extends DefaultHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final int headerRowIndex;
        private final List<String> headers = new ArrayList<>();
        private final List<String> currentRow = new ArrayList<>();
        private boolean done = false;

        HeaderExtractor(int headerRowIndex) {
            this.headerRowIndex = headerRowIndex;
        }

        @Override
        public void startRow(int rowNum) {
            currentRow.clear();
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum == headerRowIndex) {
                headers.addAll(currentRow);
                done = true;
            }
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (!done) {
                int colIndex = ExcelReadSupport.getColumnIndex(cellReference);
                while (currentRow.size() < colIndex) {
                    currentRow.add("");
                }
                currentRow.add(formattedValue != null ? formattedValue : "");
            }
        }

        List<String> getHeaders() {
            return headers;
        }
    }
}
