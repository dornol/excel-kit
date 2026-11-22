package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
import jakarta.validation.Validator;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import jakarta.validation.Validation;

import java.util.Locale;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Read settings given to an {@link ExcelKitSchema} must configure the readers it creates exactly as
 * if they had been chained on a reader built by hand.
 */
class SchemaReaderSettingsTest {

    static class Row {
        String name;
    }

    private static final CellConversionConfig CONVERSION =
            CellConversionConfig.builder().locale(Locale.GERMANY).build();
    private static final Validator VALIDATOR = Validation.byDefaultProvider().configure()
            .messageInterpolator(new ParameterMessageInterpolator())
            .buildValidatorFactory()
            .getValidator();

    /** A schema-level setting and the equivalent chained on a reader. */
    record Case(String name, Consumer<ExcelKitSchema.Builder<Row>> onSchema,
                Consumer<AbstractReader<Row, ?>> onReader) {
        @Override
        public String toString() {
            return name;
        }
    }

    static Stream<Case> settings() {
        return Stream.of(
                new Case("strictHeaders", ExcelKitSchema.Builder::strictHeaders, r -> r.strictHeaders()),
                new Case("duplicateHeaderPolicy",
                        b -> b.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST),
                        r -> r.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)),
                new Case("cellConversion", b -> b.cellConversion(CONVERSION), r -> r.cellConversion(CONVERSION)),
                new Case("maxRows", b -> b.maxRows(7), r -> r.maxRows(7)),
                new Case("skipBlankRows", ExcelKitSchema.Builder::skipBlankRows, r -> r.skipBlankRows()),
                new Case("stopAtBlankRows", b -> b.stopAtBlankRows(2), r -> r.stopAtBlankRows(2))
        );
    }

    private static ExcelKitSchema.Builder<Row> schemaBuilder() {
        return ExcelKitSchema.<Row>builder().column("Name", row -> row.name, (row, cell) -> row.name = cell.asString());
    }

    @ParameterizedTest
    @MethodSource("settings")
    void csvReaderFromTheSchemaMatchesAHandConfiguredReader(Case setting) {
        ExcelKitSchema.Builder<Row> builder = schemaBuilder();
        setting.onSchema().accept(builder);
        CsvReader<Row> fromSchema = builder.build().csvReader(Row::new);

        CsvReader<Row> byHand = CsvReader.setter(Row::new);
        setting.onReader().accept(byHand);

        assertEquals(byHand.snapshotReadOptions(), fromSchema.snapshotReadOptions());
        assertNotEquals(CsvReader.setter(Row::new).snapshotReadOptions(), fromSchema.snapshotReadOptions(),
                setting.name() + " did not change anything, so the comparison proves nothing");
    }

    @ParameterizedTest
    @MethodSource("settings")
    void excelReaderFromTheSchemaMatchesAHandConfiguredReader(Case setting) {
        ExcelKitSchema.Builder<Row> builder = schemaBuilder();
        setting.onSchema().accept(builder);
        ExcelReader<Row> fromSchema = builder.build().excelReader(Row::new);

        ExcelReader<Row> byHand = ExcelReader.setter(Row::new);
        setting.onReader().accept(byHand);

        assertEquals(byHand.snapshotReadOptions(), fromSchema.snapshotReadOptions());
    }

    @ParameterizedTest
    @MethodSource("settings")
    void mappingReadersGetTheSameSettings(Case setting) {
        ExcelKitSchema.Builder<Row> builder = schemaBuilder();
        setting.onSchema().accept(builder);
        ExcelKitSchema<Row> schema = builder.build();

        CsvReader<Row> csv = CsvReader.mapping(row -> new Row());
        setting.onReader().accept(csv);
        ExcelReader<Row> excel = ExcelReader.mapping(row -> new Row());
        setting.onReader().accept(excel);

        assertEquals(csv.snapshotReadOptions(), schema.csvReader(row -> new Row()).snapshotReadOptions());
        assertEquals(excel.snapshotReadOptions(), schema.excelReader(row -> new Row()).snapshotReadOptions());
    }

    @Test
    void theValidatorGivenWhenTheReaderIsCreatedIsKept() {
        ExcelKitSchema.Builder<Row> builder = schemaBuilder().maxRows(3).strictHeaders();
        ExcelKitSchema<Row> schema = builder.build();

        assertSame(VALIDATOR, schema.csvReader(Row::new, VALIDATOR).configuredValidator());
        assertSame(VALIDATOR, schema.excelReader(Row::new, VALIDATOR).configuredValidator());
        assertSame(VALIDATOR, schema.csvReader(row -> new Row(), VALIDATOR).configuredValidator());
        assertSame(VALIDATOR, schema.excelReader(row -> new Row(), VALIDATOR).configuredValidator());
        assertEquals(null, schema.csvReader(Row::new).configuredValidator());
    }

    @Test
    void settingsFromTheSchemaAreNotSharedBetweenItsReaders() {
        ExcelKitSchema<Row> schema = schemaBuilder().maxRows(5).build();

        CsvReader<Row> first = schema.csvReader(Row::new);
        first.maxRows(99).skipBlankRows();
        CsvReader<Row> second = schema.csvReader(Row::new);

        assertEquals(5, second.snapshotReadOptions().maxRows());
        assertEquals(false, second.snapshotReadOptions().skipBlankRows());
    }

    @Test
    void settingsGivenOnTheReaderAfterwardsStillWin() {
        ExcelKitSchema<Row> schema = schemaBuilder().maxRows(5).build();

        assertEquals(20, schema.csvReader(Row::new).maxRows(20).snapshotReadOptions().maxRows());
    }

    @Test
    void theBuilderRejectsInvalidValues() {
        assertThrows(IllegalArgumentException.class, () -> schemaBuilder().maxRows(-1));
        assertThrows(IllegalArgumentException.class, () -> schemaBuilder().stopAtBlankRows(-1));
        assertThrows(NullPointerException.class, () -> schemaBuilder().duplicateHeaderPolicy(null));
        assertThrows(NullPointerException.class, () -> schemaBuilder().cellConversion(null));
    }
}
