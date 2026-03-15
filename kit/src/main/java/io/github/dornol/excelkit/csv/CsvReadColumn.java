package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;

/**
 * Represents a single CSV column binding for reading.
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
public record CsvReadColumn<T>(String headerName, int columnIndex, BiConsumer<T, CellData> setter) {

    public CsvReadColumn(BiConsumer<T, CellData> setter) {
        this(null, -1, setter);
    }

    public CsvReadColumn(String headerName, BiConsumer<T, CellData> setter) {
        this(headerName, -1, setter);
    }

    public static class CsvReadColumnBuilder<T> {
        private final CsvReader<T> reader;
        private final String headerName;
        private final int columnIndex;
        private final BiConsumer<T, CellData> setter;

        CsvReadColumnBuilder(CsvReader<T> reader, BiConsumer<T, CellData> setter) {
            this(reader, null, -1, setter);
        }

        CsvReadColumnBuilder(CsvReader<T> reader, String headerName, BiConsumer<T, CellData> setter) {
            this(reader, headerName, -1, setter);
        }

        CsvReadColumnBuilder(CsvReader<T> reader, int columnIndex, BiConsumer<T, CellData> setter) {
            this(reader, null, columnIndex, setter);
        }

        CsvReadColumnBuilder(CsvReader<T> reader, String headerName, int columnIndex, BiConsumer<T, CellData> setter) {
            if (setter == null) {
                throw new IllegalArgumentException("setter must not be null");
            }
            this.reader = reader;
            this.headerName = headerName;
            this.columnIndex = columnIndex;
            this.setter = setter;
        }

        public CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumnBuilder<>(reader, setter);
        }

        public CsvReadColumnBuilder<T> column(String headerName, BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumnBuilder<>(reader, headerName, setter);
        }

        /**
         * Adds the current column and begins a new index-based column definition.
         *
         * @param columnIndex 0-based column index in the CSV file
         * @param setter      the setter function
         */
        public CsvReadColumnBuilder<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumnBuilder<>(reader, columnIndex, setter);
        }

        public CsvReader<T> skipColumn() {
            buildCurrentAndAddToReader();
            return reader.skipColumn();
        }

        public CsvReader<T> skipColumns(int count) {
            buildCurrentAndAddToReader();
            return reader.skipColumns(count);
        }

        public CsvReadHandler<T> build(InputStream inputStream) {
            buildCurrentAndAddToReader();
            return this.reader.build(inputStream);
        }

        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new CsvReadColumn<>(this.headerName, this.columnIndex, this.setter));
        }
    }
}
