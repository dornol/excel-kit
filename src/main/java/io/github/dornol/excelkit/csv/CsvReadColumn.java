package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;

public record CsvReadColumn<T>(BiConsumer<T, CellData> setter) {

    public static class CsvReadColumnBuilder<T> {
        private final CsvReader<T> reader;
        private final BiConsumer<T, CellData> setter;

        CsvReadColumnBuilder(CsvReader<T> reader, BiConsumer<T, CellData> setter) {
            this.reader = reader;
            this.setter = setter;
        }

        public CsvReadColumn.CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumn.CsvReadColumnBuilder<>(reader, setter);
        }

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
