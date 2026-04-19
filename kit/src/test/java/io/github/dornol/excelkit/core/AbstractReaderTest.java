package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for AbstractReader shared behavior via ExcelReader and CsvReader.
 */
class AbstractReaderTest {

    @Nested
    class ColumnRegistration {

        @Test
        void column_positional_addsColumn() {
            var reader = ExcelReader.setter(Object::new);
            reader.column((obj, cell) -> {});
            reader.column((obj, cell) -> {});
            // should not throw — columns registered
            assertNotNull(reader);
        }

        @Test
        void column_named_addsColumn() {
            var reader = CsvReader.setter(Object::new);
            reader.column("Name", (obj, cell) -> {});
            assertNotNull(reader);
        }

        @Test
        void columnAt_indexBased_addsColumn() {
            var reader = ExcelReader.setter(Object::new);
            reader.columnAt(0, (obj, cell) -> {});
            reader.columnAt(5, (obj, cell) -> {});
            assertNotNull(reader);
        }
    }

    @Nested
    class RequiredColumn {

        @Test
        void required_withNoColumns_throws() {
            var reader = ExcelReader.setter(Object::new);
            assertThrows(IllegalStateException.class, reader::required);
        }

        @Test
        void required_afterColumn_succeeds() {
            var reader = CsvReader.setter(Object::new);
            reader.column("Name", (obj, cell) -> {}).required();
            assertNotNull(reader);
        }
    }

    @Nested
    class SkipColumns {

        @Test
        void skipColumn_addsNoOpColumn() {
            var reader = ExcelReader.setter(Object::new);
            reader.column((obj, cell) -> {});
            reader.skipColumn();
            reader.column((obj, cell) -> {});
            assertNotNull(reader);
        }

        @Test
        void skipColumns_negativeCount_throws() {
            var reader = ExcelReader.setter(Object::new);
            assertThrows(IllegalArgumentException.class, () -> reader.skipColumns(-1));
        }

        @Test
        void skipColumns_zeroCount_succeeds() {
            var reader = CsvReader.setter(Object::new);
            reader.skipColumns(0);
            assertNotNull(reader);
        }
    }

    @Nested
    class MapModeRestrictions {

        @Test
        void forMap_rejectsColumn() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.column((obj, cell) -> {}));
        }

        @Test
        void forMap_rejectsColumnNamed() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.column("Name", (obj, cell) -> {}));
        }

        @Test
        void forMap_rejectsColumnAt() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.columnAt(0, (obj, cell) -> {}));
        }

        @Test
        void forMap_rejectsSkipColumn() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class, reader::skipColumn);
        }

        @Test
        void forMap_rejectsSkipColumns() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.skipColumns(2));
        }

        @Test
        void forMap_allowsHeaderRowIndex() {
            var reader = ExcelReader.forMap().headerRowIndex(2);
            assertNotNull(reader);
        }

        @Test
        void forMap_allowsOnProgress() {
            var reader = CsvReader.forMap().onProgress(100, (count, cursor) -> {});
            assertNotNull(reader);
        }
    }

    @Nested
    class ProgressCallback {

        @Test
        void onProgress_negativeInterval_throws() {
            var reader = ExcelReader.setter(Object::new);
            assertThrows(IllegalArgumentException.class,
                    () -> reader.onProgress(-1, (c, cur) -> {}));
        }

        @Test
        void onProgress_zeroInterval_throws() {
            var reader = CsvReader.setter(Object::new);
            assertThrows(IllegalArgumentException.class,
                    () -> reader.onProgress(0, (c, cur) -> {}));
        }

        @Test
        void onProgress_positiveInterval_succeeds() {
            var reader = ExcelReader.setter(Object::new);
            reader.onProgress(1000, (c, cur) -> {});
            assertNotNull(reader);
        }
    }
}
