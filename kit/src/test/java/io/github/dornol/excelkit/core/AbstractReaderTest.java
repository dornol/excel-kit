package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import jakarta.validation.Validation;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;

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
    class ValidatorConfiguration {

        @Test
        void validator_isFluentOnExcelReader() {
            try (var factory = Validation.byDefaultProvider()
                    .configure()
                    .messageInterpolator(new ParameterMessageInterpolator())
                    .buildValidatorFactory()) {
                var reader = ExcelReader.setter(Object::new);
                assertSame(reader, reader.validator(factory.getValidator()));
            }
        }

        @Test
        void validator_isFluentOnCsvReader() {
            try (var factory = Validation.byDefaultProvider()
                    .configure()
                    .messageInterpolator(new ParameterMessageInterpolator())
                    .buildValidatorFactory()) {
                var reader = CsvReader.setter(Object::new);
                assertSame(reader, reader.validator(factory.getValidator()));
            }
        }

        @Test
        void validator_rejectsNull() {
            var reader = ExcelReader.setter(Object::new);
            assertThrows(NullPointerException.class, () -> reader.validator(null));
        }
    }

    @Nested
    class NullConfiguration {

        @Test
        void columnCallbacks_rejectNull() {
            var reader = ExcelReader.setter(Object::new);

            assertThrows(NullPointerException.class, () -> reader.column(null));
            assertThrows(NullPointerException.class, () -> reader.column("Name", null));
            assertThrows(NullPointerException.class, () -> reader.column(List.of("Name"), null));
            assertThrows(NullPointerException.class, () -> reader.columnAt(0, null));
        }

        @Test
        void callbacksAndConfigurers_rejectNull() {
            var reader = CsvReader.setter(Object::new);

            assertThrows(NullPointerException.class, () -> reader.onProgress(1, null));
            assertThrows(NullPointerException.class, () -> reader.onReadProgress(1, null));
            assertThrows(NullPointerException.class,
                    () -> reader.cellConversion((java.util.function.Consumer<CellConversionConfig.Builder>) null));
        }

        @Test
        void readEntryPoints_rejectNullBeforeExecution() {
            var excel = ExcelReader.forMap();
            var csv = CsvReader.forMap();
            var input = java.io.InputStream.nullInputStream();

            assertThrows(NullPointerException.class, () -> excel.read(input, null));
            assertThrows(NullPointerException.class, () -> csv.read(input, null));
            assertThrows(NullPointerException.class, () -> excel.readWhile(input, null));
            assertThrows(NullPointerException.class, () -> csv.readWhile(input, null));
            assertThrows(NullPointerException.class, () -> excel.read((InputStreamSource) null, result -> {}));
            assertThrows(IllegalArgumentException.class, () -> csv.readReport(input, -1));
        }
    }

    @Nested
    class ReadResultMessages {

        @Test
        void messages_areNonNullAndImmutable() {
            var result = ReadResult.success("value");

            assertTrue(result.messages().isEmpty());
            assertThrows(UnsupportedOperationException.class,
                    () -> result.messages().add("unexpected"));
        }

        @Test
        void messages_areCopied() {
            var messages = new java.util.ArrayList<>(List.of("initial"));
            var result = ReadResult.validationFailure("value", messages, -1, List.of(), List.of());

            messages.add("later");

            assertEquals(List.of("initial"), result.messages());
        }

        @Test
        void legacyConstructor_infersErrorTypeFromCause() {
            var result = ReadResult.mappingFailure("value", List.of("mapping failed"),
                    new IllegalArgumentException("bad value"), 2L, List.of(), List.of("bad"));

            assertEquals(RowError.Type.MAPPING, result.errorType());
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
