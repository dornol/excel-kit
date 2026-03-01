package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;

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
    private static final int DEFAULT_MAX_FILE_COUNT = 10_000_000;
    private static final int DEFAULT_MAX_BYTE_ARRAY_SIZE = 2_000_000_000;

    private final List<ExcelReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;
    private int sheetIndex = 0;

    /**
     * Configures Apache POI's internal limits for reading large Excel files.
     * <p>
     * This adjusts:
     * <ul>
     *     <li>{@code ZipSecureFile.setMaxFileCount(10_000_000)} — max internal zip entries</li>
     *     <li>{@code IOUtils.setByteArrayMaxOverride(2_000_000_000)} — max in-memory byte array size</li>
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
     * @param maxFileCount       Maximum number of zip entries (default: 10,000,000)
     * @param maxByteArraySize   Maximum byte array size in bytes (default: 2,000,000,000)
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
    public ExcelReader(Supplier<T> instanceSupplier, Validator validator) {
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
     * Adds a column mapping to the internal list.
     *
     * @param column An Excel column with setter logic
     */
    void addColumn(ExcelReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Begins a new column mapping using a setter function.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link io.github.dornol.excelkit.shared.CellData} to the row object
     * @return A builder for further column configuration
     */
    public ExcelReadColumn.ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new ExcelReadColumn.ExcelReadColumnBuilder<>(this, setter);
    }

    /**
     * Finalizes the configuration and builds an {@link ExcelReadHandler} for parsing the given Excel stream.
     *
     * @param inputStream The input stream of the Excel file
     * @return A handler to execute Excel parsing
     */
    public ExcelReadHandler<T> build(InputStream inputStream) {
        return new ExcelReadHandler<>(inputStream, columns, instanceSupplier, validator, sheetIndex);
    }
}
