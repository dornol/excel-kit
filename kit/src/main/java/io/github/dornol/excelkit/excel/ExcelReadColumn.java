package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;


/**
 * Represents a single Excel column binding for reading.
 * <p>
 * Holds a setter function that maps a {@link io.github.dornol.excelkit.shared.CellData} into a field of a row object.
 *
 * @param <T> The row data type
 * @author dhkim
 * @since 2025-07-19
 */
public record ExcelReadColumn<T>(BiConsumer<T, CellData> setter) {

    /**
     * Builder for defining multiple Excel read columns fluently.
     *
     * @param <T> The row data type
     */
    public static class ExcelReadColumnBuilder<T> {
        private final ExcelReader<T> reader;
        private final BiConsumer<T, CellData> setter;

        /**
         * Constructs a new column builder.
         *
         * @param reader The parent {@link ExcelReader}
         * @param setter The setter function to bind a column value to a field
         */
        ExcelReadColumnBuilder(ExcelReader<T> reader, BiConsumer<T, CellData> setter) {
            this.reader = reader;
            this.setter = setter;
        }

        /**
         * Adds the current column binding to the reader and begins a new column definition.
         * <p>
         * Note: This method adds the current column to the reader and returns a new builder
         * instance for the next column. The current builder instance should not be reused.
         * </p>
         *
         * @param setter The setter function for the next column
         * @return A new builder instance for chaining the next column
         */
        public ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, setter);
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
         * Internal method to add the current column definition to the reader.
         */
        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new ExcelReadColumn<>(this.setter));
        }
    }

}
