package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.util.function.BiConsumer;


/**
 * Represents a single Excel column binding for reading.
 * <p>
 * Supports three matching modes:
 * <ul>
 *     <li>Positional (default) — matched by insertion order</li>
 *     <li>Name-based — matched by header name via {@code headerName}</li>
 *     <li>Index-based — matched by explicit column index via {@code columnIndex}</li>
 * </ul>
 *
 * @param headerName  optional header name for name-based column matching (null for positional/index)
 * @param columnIndex explicit 0-based column index (-1 for positional/name-based)
 * @param setter      the setter function to bind a column value to a field
 * @param <T> The row data type
 * @author dhkim
 * @since 2025-07-19
 */
public record ExcelReadColumn<T>(@Nullable String headerName, int columnIndex, BiConsumer<T, CellData> setter) {

    /**
     * Creates a positional column binding (matched by column index order).
     */
    public ExcelReadColumn(BiConsumer<T, CellData> setter) {
        this(null, -1, setter);
    }

    /**
     * Creates a name-based column binding.
     */
    public ExcelReadColumn(String headerName, BiConsumer<T, CellData> setter) {
        this(headerName, -1, setter);
    }

    /**
     * Builder for defining multiple Excel read columns fluently.
     *
     * @param <T> The row data type
     */
    public static class ExcelReadColumnBuilder<T> {
        private final ExcelReader<T> reader;
        private final String headerName;
        private final int columnIndex;
        private final BiConsumer<T, CellData> setter;

        ExcelReadColumnBuilder(ExcelReader<T> reader, BiConsumer<T, CellData> setter) {
            this(reader, null, -1, setter);
        }

        ExcelReadColumnBuilder(ExcelReader<T> reader, String headerName, BiConsumer<T, CellData> setter) {
            this(reader, headerName, -1, setter);
        }

        ExcelReadColumnBuilder(ExcelReader<T> reader, int columnIndex, BiConsumer<T, CellData> setter) {
            this(reader, null, columnIndex, setter);
        }

        ExcelReadColumnBuilder(ExcelReader<T> reader, String headerName, int columnIndex, BiConsumer<T, CellData> setter) {
            if (setter == null) {
                throw new IllegalArgumentException("setter must not be null");
            }
            this.reader = reader;
            this.headerName = headerName;
            this.columnIndex = columnIndex;
            this.setter = setter;
        }

        /**
         * Adds the current column and begins a new positional column definition.
         */
        public ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, setter);
        }

        /**
         * Adds the current column and begins a new name-based column definition.
         */
        public ExcelReadColumnBuilder<T> column(String headerName, BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, headerName, setter);
        }

        /**
         * Adds the current column and begins a new index-based column definition.
         *
         * @param columnIndex 0-based column index in the Excel file
         * @param setter      the setter function
         */
        public ExcelReadColumnBuilder<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new ExcelReadColumnBuilder<>(reader, columnIndex, setter);
        }

        public ExcelReader<T> skipColumn() {
            buildCurrentAndAddToReader();
            return reader.skipColumn();
        }

        public ExcelReader<T> skipColumns(int count) {
            buildCurrentAndAddToReader();
            return reader.skipColumns(count);
        }

        public ExcelReadHandler<T> build(InputStream inputStream) {
            buildCurrentAndAddToReader();
            return this.reader.build(inputStream);
        }

        public ExcelReadHandler<T> build(InputStream inputStream, int sheetIndex) {
            buildCurrentAndAddToReader();
            this.reader.sheetIndex(sheetIndex);
            return this.reader.build(inputStream);
        }

        /**
         * Registers a progress callback and returns the parent reader.
         */
        public ExcelReader<T> onProgress(int interval, io.github.dornol.excelkit.shared.ProgressCallback callback) {
            buildCurrentAndAddToReader();
            return this.reader.onProgress(interval, callback);
        }

        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new ExcelReadColumn<>(this.headerName, this.columnIndex, this.setter));
        }
    }
}
