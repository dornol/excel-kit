package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;
import org.jspecify.annotations.NonNull;

import io.github.dornol.excelkit.shared.ProgressCallback;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
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
}
