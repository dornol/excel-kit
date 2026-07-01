package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriteErrorPolicy;
import io.github.dornol.excelkit.excel.ExcelWriteException;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class NewFeaturesV019Test {

    @Test
    void readerScopedCellConversion_shouldParseCustomDateWithoutGlobalMutation() {
        byte[] csv = "Date\n25.12.2024\n".getBytes(StandardCharsets.UTF_8);
        List<LocalDate> dates = new ArrayList<>();

        CsvReader.<LocalDate>mapping(row -> row.get("Date").asLocalDate())
                .cellConversion(c -> c.addDateFormat("dd.MM.yyyy"))
                .build(new ByteArrayInputStream(csv))
                .read(result -> dates.add(result.data()));

        assertEquals(List.of(LocalDate.of(2024, 12, 25)), dates);
    }

    @Test
    void readerScopedCellConversion_shouldUseConfiguredLocale() {
        byte[] csv = "Amount\n1.234,5\n".getBytes(StandardCharsets.UTF_8);
        List<Double> values = new ArrayList<>();

        CsvReader.<Double>mapping(row -> row.get("Amount").asDouble())
                .delimiter(';')
                .cellConversion(c -> c.locale(Locale.GERMANY))
                .build(new ByteArrayInputStream(csv))
                .read(result -> values.add(result.data()));

        assertEquals(1234.5d, values.getFirst(), 0.0001);
    }

    @Test
    void csvReader_shouldExposeQuoteParserOptions() {
        byte[] csv = "Name\n'Alice, A'\n".getBytes(StandardCharsets.UTF_8);
        List<String> values = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .quoteChar('\'')
                .build(new ByteArrayInputStream(csv))
                .read(result -> values.add(result.data()));

        assertEquals(List.of("Alice, A"), values);
    }

    @Test
    void excelWriter_failFast_shouldThrowWhenValueExtractorFails() {
        ExcelWriter<String> writer = ExcelWriter.<String>create()
                .writeErrorPolicy(ExcelWriteErrorPolicy.FAIL_FAST)
                .column("Name", value -> {
                    throw new IllegalStateException("boom");
                });

        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("x")));
    }

    @Test
    void excelWriter_failFast_shouldThrowWhenCellWriteFails() {
        ExcelWriter<String> writer = ExcelWriter.<String>create()
                .writeErrorPolicy(ExcelWriteErrorPolicy.FAIL_FAST)
                .column("Age", value -> "not-number", c -> c.type(ExcelDataType.INTEGER));

        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("x")));
    }

    @Test
    void schemaExcelWriter_shouldAcceptInitOptions() {
        ExcelKitSchema<String> schema = ExcelKitSchema.<String>builder()
                .column("Name", value -> value, (value, cell) -> {})
                .build();

        ExcelWriter<String> writer = schema.excelWriter(opts -> opts.rowAccessWindowSize(10));

        writer.write(Stream.of("Alice"));
    }

    @Test
    void csvReader_shouldLimitRowsAndSkipBlankRows() {
        byte[] csv = "Name\nAlice\n\nBob\nCarol\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .skipBlankRows()
                .maxRows(2)
                .build(new ByteArrayInputStream(csv))
                .read(result -> names.add(result.data()));

        assertEquals(List.of("Alice", "Bob"), names);
    }

    @Test
    void csvReader_shouldStopAtConsecutiveBlankRows() {
        byte[] csv = "Name\nAlice\n\n\nBob\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .skipBlankRows()
                .stopAtBlankRows(2)
                .build(new ByteArrayInputStream(csv))
                .read(result -> names.add(result.data()));

        assertEquals(List.of("Alice"), names);
    }

    @Test
    void readResult_shouldCarryRawValues() {
        byte[] csv = "Name,Age\nAlice,30\n".getBytes(StandardCharsets.UTF_8);
        List<List<String>> rawRows = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .build(new ByteArrayInputStream(csv))
                .read(result -> rawRows.add(result.rawValues()));

        assertEquals(List.of(List.of("Alice", "30")), rawRows);
    }

    @Test
    void schemaDefaults_shouldApplyToReaders() {
        ExcelKitSchema<String> schema = ExcelKitSchema.<String>builder()
                .column("Name", value -> value, (value, cell) -> {})
                .strictHeaders()
                .maxRows(1)
                .build();
        byte[] csv = "Name\nAlice\nBob\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        schema.csvReader(() -> "")
                .build(new ByteArrayInputStream(csv))
                .read(result -> names.add(result.data()));

        assertEquals(1, names.size());
    }
}
