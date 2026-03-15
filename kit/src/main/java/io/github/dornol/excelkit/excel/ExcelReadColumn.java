package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;


/**
 * Represents a single Excel column binding for reading.
 * <p>
 * Holds a setter function that maps a {@link io.github.dornol.excelkit.shared.CellData} into a field of a row object.
 * Optionally includes a header name for name-based column matching instead of positional matching.
 *
 * @param headerName optional header name for name-based column matching (null for positional)
 * @param setter     the setter function to bind a column value to a field
 * @param <T> The row data type
 * @author dhkim
 * @since 2025-07-19
 */
public record ExcelReadColumn<T>(String headerName, BiConsumer<T, CellData> setter) {

    /**
     * Creates a positional column binding (matched by column index).
     *
     * @param setter the setter function
     */
    public ExcelReadColumn(BiConsumer<T, CellData> setter) {
        this(null, setter);
    }

    /**
     * Builder for defining multiple Excel read columns fluently.
     *
     * @param <T> The row data type
     */
    public static class ExcelReadColumnBuilder<T> {
        private final ExcelReader<T> reader;
        private final String headerName;
        private final BiConsumer<T, CellData> setter;

        /**
         * Constructs a new column builder for positional matching.
         *
         * @param reader The parent {@link ExcelReader}
         * @param setter The setter function to bind a column value to a field
         */
        ExcelReadColumnBuilder(ExcelReader<T> reader, BiConsumer<T, CellData> setter) {
            this(reader, null, setter);
        }

        /**
         * Constructs a new column builder with optional header name matching.
         *
         * @param reader     The parent {@link ExcelReader}
         * @param headerName The header name to match (null for positional)
         * @param setter     The setter function to bind a column value to a field
         */
        ExcelReadColumnBuilder(ExcelReader<T> reader, String headerName, BiConsumer<T, CellData> setter) {
            if (setter == null) {
                throw new IllegalArgumentException("setter must not be null");
            }
            this.reader = reader;
            this.headerName = headerName;
            this.setter = setter;
        }

        /**
         * Adds the current column binding to the reader and begins a new positional column definition.
         *
         * @param setter The setter function for the next column
         * @return A new builder instance for chaining the next column
         */
        public ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, setter);
        }

        /**
         * Adds the current column binding to the reader and begins a new name-based column definition.
         *
         * @param headerName The header name to match in the Excel file
         * @param setter     The setter function for the next column
         * @return A new builder instance for chaining the next column
         */
        public ExcelReadColumnBuilder<T> column(String headerName, BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, headerName, setter);
        }

        /**
         * Flushes the current column and skips one column by adding a no-op column mapping.
         *
         * @return The parent ExcelReader for chaining
         */
        public ExcelReader<T> skipColumn() {
            buildCurrentAndAddToReader();
            return reader.skipColumn();
        }

        /**
         * Flushes the current column and skips the specified number of columns.
         *
         * @param count The number of columns to skip (must be non-negative)
         * @return The parent ExcelReader for chaining
         * @throws IllegalArgumentException if count is negative
         */
        public ExcelReader<T> skipColumns(int count) {
            buildCurrentAndAddToReader();
            return reader.skipColumns(count);
        }

        /**
         * Finalizes the column definitions and builds an {@link ExcelReadHandler} for reading.
         *
         * @param inputStream The input stream of the Excel (.xlsx) file
         * @return A configured {@code ExcelReadHandler} instance
         */
        public ExcelReadHandler<T> build(InputStream inputStream) {
            buildCurrentAndAddToReader();
            return this.reader.build(inputStream);
        }

        /**
         * Finalizes the column definitions and builds an {@link ExcelReadHandler} for reading
         * from a specific sheet.
         *
         * @param inputStream The input stream of the Excel (.xlsx) file
         * @param sheetIndex  The zero-based index of the sheet to read
         * @return A configured {@code ExcelReadHandler} instance
         */
        public ExcelReadHandler<T> build(InputStream inputStream, int sheetIndex) {
            buildCurrentAndAddToReader();
            this.reader.sheetIndex(sheetIndex);
            return this.reader.build(inputStream);
        }

        /**
         * Internal method to add the current column definition to the reader.
         */
        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new ExcelReadColumn<>(this.headerName, this.setter));
        }
    }

}
