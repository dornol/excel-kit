package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadColumn;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ExcelKitException;
import io.github.dornol.excelkit.shared.RowData;
import io.github.dornol.excelkit.shared.TempResourceCreator;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;
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

import io.github.dornol.excelkit.shared.ProgressCallback;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * Builder-style class for configuring Excel row readers.
 * <p>
 * {@code ExcelReader} allows you to define how each Excel cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, use {@link #build(InputStream)} to create a {@link ExcelReadHandler}.
 *
 * @param <T> The type of the object that represents one Excel row
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelReader<T> {
    private static final int DEFAULT_MAX_FILE_COUNT = 1_000_000;
    private static final int DEFAULT_MAX_BYTE_ARRAY_SIZE = 500_000_000;

    private final List<ReadColumn<T>> columns = new ArrayList<>();
    private final @Nullable Supplier<T> instanceSupplier;
    private final @Nullable Function<RowData, T> rowMapper;
    private final @Nullable Validator validator;
    private int sheetIndex = 0;
    private int headerRowIndex = 0;
    private @Nullable ProgressCallback progressCallback;
    private int progressInterval;
    private boolean mapMode = false;

    /**
     * Configures Apache POI's internal limits for reading large Excel files.
     * <p>
     * This adjusts:
     * <ul>
     *     <li>{@code ZipSecureFile.setMaxFileCount(1_000_000)} — max internal zip entries</li>
     *     <li>{@code IOUtils.setByteArrayMaxOverride(500_000_000)} — max in-memory byte array size</li>
     * </ul>
     * <p>
     * <b>Note:</b> These are JVM-global settings and affect all POI operations in the same process.
     * Call this method once at application startup if you need to read large files.
     */
    public static void configureLargeFileSupport() {
        configureLargeFileSupport(DEFAULT_MAX_FILE_COUNT, DEFAULT_MAX_BYTE_ARRAY_SIZE);
    }

    /**
     * Configures Apache POI's internal limits with custom values.
     *
     * @param maxFileCount       Maximum number of zip entries (default: 1,000,000)
     * @param maxByteArraySize   Maximum byte array size in bytes (default: 500,000,000)
     * @see #configureLargeFileSupport()
     */
    public static void configureLargeFileSupport(int maxFileCount, int maxByteArraySize) {
        ZipSecureFile.setMaxFileCount(maxFileCount);
        IOUtils.setByteArrayMaxOverride(maxByteArraySize);
    }

    /**
     * Constructs an ExcelReader in setter mode with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public ExcelReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.rowMapper = null;
        this.validator = validator;
    }

    private ExcelReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        this.instanceSupplier = null;
        this.rowMapper = Objects.requireNonNull(rowMapper, "rowMapper cannot be null");
        this.validator = validator;
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
     * )).build(inputStream).read(result -> { ... });
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
     *     .build(inputStream)
     *     .read(result -> {
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
     *     .build(inputStream)
     *     .read(result -> {
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
                if (selectedColumns != null && !selectedColumns.contains(header)) continue;
                map.put(header, row.get(i).formattedValue());
            }
            return map;
        };
        ExcelReader<Map<String, String>> reader = ExcelReader.mapping(mapMapper);
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
     * Sets the zero-based sheet index to read from.
     * Defaults to 0 (the first sheet).
     *
     * @param sheetIndex The zero-based index of the sheet to read
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> sheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * Sets the zero-based row index of the header row.
     * Rows before this index will be skipped during reading.
     * Defaults to 0 (the first row).
     *
     * @param headerRowIndex The zero-based index of the header row
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    /**
     * Adds a column mapping to the internal list.
     *
     * @param column An Excel column with setter logic
     */
    void addColumn(ReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Registers a positional column mapping. Columns are matched to the spreadsheet in
     * the order they are registered (after {@link #headerRowIndex(int)} is accounted for).
     *
     * @param setter a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public ExcelReader<T> column(BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(BiConsumer)");
        columns.add(new ReadColumn<>(setter));
        return this;
    }

    /**
     * Registers a name-based column mapping. The column is matched to the spreadsheet
     * column whose header cell equals {@code headerName} in the header row.
     *
     * @param headerName the header name to match
     * @param setter     a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public ExcelReader<T> column(String headerName, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(String, BiConsumer)");
        columns.add(new ReadColumn<>(headerName, setter));
        return this;
    }

    /**
     * Registers an index-based column mapping. The column is matched to the spreadsheet
     * column at the given 0-based index, regardless of header or registration order.
     *
     * @param columnIndex 0-based column index in the Excel file
     * @param setter      a {@code BiConsumer} that writes a cell value into the row object
     * @return this reader for chaining
     */
    public ExcelReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        requireNotMapMode("columnAt(int, BiConsumer)");
        columns.add(new ReadColumn<>(null, columnIndex, setter));
        return this;
    }

    /**
     * Skips one column during reading by registering a no-op mapping at the next
     * positional slot.
     *
     * @return this reader for chaining
     */
    public ExcelReader<T> skipColumn() {
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
    public ExcelReader<T> skipColumns(int count) {
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
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Finalizes the configuration and builds an {@link ExcelReadHandler} for parsing the given Excel stream.
     *
     * @param inputStream The input stream of the Excel file
     * @return A handler to execute Excel parsing
     */
    public ExcelReadHandler<T> build(InputStream inputStream) {
        if (rowMapper != null) {
            return new ExcelReadHandler<>(inputStream, rowMapper, validator,
                    sheetIndex, headerRowIndex, progressInterval, progressCallback);
        }
        return new ExcelReadHandler<>(inputStream, columns, instanceSupplier, validator,
                sheetIndex, headerRowIndex, progressInterval, progressCallback);
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
            try (InputStream is = inputStream) {
                Files.copy(is, tempFile, StandardCopyOption.REPLACE_EXISTING);
            }

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

    /**
     * Returns the header names from a specific sheet.
     *
     * @param inputStream    The input stream of the Excel file (will be consumed)
     * @param sheetIndex     The 0-based sheet index
     * @param headerRowIndex The 0-based header row index
     * @return A list of header names
     */
    public static List<String> getSheetHeaders(InputStream inputStream, int sheetIndex, int headerRowIndex) {
        Path tempDir = null;
        Path tempFile = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            try (InputStream is = inputStream) {
                Files.copy(is, tempFile, StandardCopyOption.REPLACE_EXISTING);
            }

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
                while (sheetsData.hasNext()) {
                    try (InputStream sheet = sheetsData.next()) {
                        if (currentIndex == sheetIndex) {
                            parser.parse(new InputSource(sheet));
                            break;
                        }
                    }
                    currentIndex++;
                }
                return extractor.getHeaders();
            }
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read sheet headers", e);
        } finally {
            cleanupTemp(tempDir, tempFile);
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
