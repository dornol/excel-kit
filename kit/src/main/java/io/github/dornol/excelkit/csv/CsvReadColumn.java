package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;

/**
 * Represents a single CSV column binding for reading.
 * <p>
 * Holds a setter function that maps a {@link CellData} into a field of a row object.
 *
 * @param <T> The row data type
 * @author dhkim
 * @since 2025-07-19
 */
public record CsvReadColumn<T>(BiConsumer<T, CellData> setter) {

    /**
     * Builder for defining multiple CSV read columns fluently.
     *
     * @param <T> The row data type
     */
    public static class CsvReadColumnBuilder<T> {
        private final CsvReader<T> reader;
        private final BiConsumer<T, CellData> setter;

        /**
         * Constructs a new column builder.
         *
         * @param reader The parent {@link CsvReader}
         * @param setter The setter function to bind a column value to a field
         */
        CsvReadColumnBuilder(CsvReader<T> reader, BiConsumer<T, CellData> setter) {
            if (setter == null) {
                throw new IllegalArgumentException("setter must not be null");
            }
            this.reader = reader;
            this.setter = setter;
        }

        /**
         * Adds the current column binding to the reader and begins a new column definition.
         *
         * @param setter The setter function for the next column
         * @return A new builder instance for chaining the next column
         */
        public CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumn.CsvReadColumnBuilder<>(reader, setter);
        }

        /**
         * Flushes the current column and skips one column by adding a no-op column mapping.
         *
         * @return The parent CsvReader for chaining
         */
        public CsvReader<T> skipColumn() {
            buildCurrentAndAddToReader();
            return reader.skipColumn();
        }

        /**
         * Flushes the current column and skips the specified number of columns.
         *
         * @param count The number of columns to skip (must be non-negative)
         * @return The parent CsvReader for chaining
         * @throws IllegalArgumentException if count is negative
         */
        public CsvReader<T> skipColumns(int count) {
            buildCurrentAndAddToReader();
            return reader.skipColumns(count);
        }

        /**
         * Finalizes the column definitions and builds a {@link CsvReadHandler} for reading.
         *
         * @param inputStream The input stream of the CSV file
         * @return A configured {@code CsvReadHandler} instance
         */
        public CsvReadHandler<T> build(InputStream inputStream) {
            buildCurrentAndAddToReader();
            return this.reader.build(inputStream);
        }

        /**
         * Internal method to add the current column definition to the reader.
         */
        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new CsvReadColumn<>(this.setter));
        }
    }

}
