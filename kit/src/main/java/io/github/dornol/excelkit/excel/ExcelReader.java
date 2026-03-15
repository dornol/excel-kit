package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ExcelKitException;
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
import org.jspecify.annotations.NonNull;
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

    private final List<ExcelReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;
    private int sheetIndex = 0;
    private int headerRowIndex = 0;
    private ProgressCallback progressCallback;
    private int progressInterval;

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
     * Constructs an ExcelReader with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public ExcelReader(@NonNull Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.validator = validator;
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
    void addColumn(ExcelReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Adds a column mapping using a setter function.
     * Useful for schema-based column registration.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> addColumn(BiConsumer<T, CellData> setter) {
        columns.add(new ExcelReadColumn<>(setter));
        return this;
    }

    /**
     * Adds a name-based column mapping using a setter function.
     * The column is matched by header name instead of positional index.
     *
     * @param headerName The header name to match in the Excel file
     * @param setter     A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> addColumn(String headerName, BiConsumer<T, CellData> setter) {
        columns.add(new ExcelReadColumn<>(headerName, setter));
        return this;
    }

    /**
     * Adds an index-based column mapping.
     * The column is matched by explicit 0-based column index.
     *
     * @param columnIndex 0-based column index in the Excel file
     * @param setter      A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        columns.add(new ExcelReadColumn<>(null, columnIndex, setter));
        return this;
    }

    /**
     * Begins a new index-based column mapping using a setter function.
     *
     * @param columnIndex 0-based column index in the Excel file
     * @param setter      A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return A builder for further column configuration
     */
    public ExcelReadColumn.ExcelReadColumnBuilder<T> columnAtBuilder(int columnIndex, BiConsumer<T, CellData> setter) {
        return new ExcelReadColumn.ExcelReadColumnBuilder<>(this, columnIndex, setter);
    }

    /**
     * Skips one column during reading by adding a no-op column mapping.
     *
     * @return This ExcelReader instance for chaining
     */
    public ExcelReader<T> skipColumn() {
        columns.add(new ExcelReadColumn<>((instance, cellData) -> {}));
        return this;
    }

    /**
     * Skips the specified number of columns during reading by adding no-op column mappings.
     *
     * @param count The number of columns to skip (must be non-negative)
     * @return This ExcelReader instance for chaining
     * @throws IllegalArgumentException if count is negative
     */
    public ExcelReader<T> skipColumns(int count) {
        if (count < 0) {
            throw new IllegalArgumentException("skipColumns count must be non-negative");
        }
        for (int i = 0; i < count; i++) {
            columns.add(new ExcelReadColumn<>((instance, cellData) -> {}));
        }
        return this;
    }

    /**
     * Begins a new positional column mapping using a setter function.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link io.github.dornol.excelkit.shared.CellData} to the row object
     * @return A builder for further column configuration
     */
    public ExcelReadColumn.ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new ExcelReadColumn.ExcelReadColumnBuilder<>(this, setter);
    }

    /**
     * Begins a new name-based column mapping using a setter function.
     * The column is matched by header name instead of positional index.
     *
     * @param headerName The header name to match in the Excel file
     * @param setter     A {@code BiConsumer} that sets a value from {@link io.github.dornol.excelkit.shared.CellData} to the row object
     * @return A builder for further column configuration
     */
    public ExcelReadColumn.ExcelReadColumnBuilder<T> column(String headerName, BiConsumer<T, CellData> setter) {
        return new ExcelReadColumn.ExcelReadColumnBuilder<>(this, headerName, setter);
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
    public ExcelReadHandler<T> build(@NonNull InputStream inputStream) {
        return new ExcelReadHandler<>(inputStream, columns, instanceSupplier, validator,
                sheetIndex, headerRowIndex, progressInterval, progressCallback);
    }

    /**
     * Returns the list of sheet names and indices from an Excel file.
     *
     * @param inputStream The input stream of the Excel file (will be consumed)
     * @return A list of {@link ExcelSheetInfo} records containing sheet names and indices
     */
    public static List<ExcelSheetInfo> getSheetNames(@NonNull InputStream inputStream) {
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
    public static List<String> getSheetHeaders(@NonNull InputStream inputStream, int sheetIndex, int headerRowIndex) {
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
                int colIndex = 0;
                for (char c : cellReference.toCharArray()) {
                    if (!Character.isLetter(c)) break;
                    colIndex = colIndex * 26 + (Character.toUpperCase(c) - 'A' + 1);
                }
                colIndex--;
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
