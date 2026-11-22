package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class CsvReaderRowOptionsTest {

    @Test
    void csvReader_shouldExposeQuoteParserOptions() {
        byte[] csv = "Name\n'Alice, A'\n".getBytes(StandardCharsets.UTF_8);
        List<String> values = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .quoteChar('\'')
                .read(new ByteArrayInputStream(csv), result -> values.add(result.data()));

        assertEquals(List.of("Alice, A"), values);
    }

    @Test
    void csvReader_shouldLimitRowsAndSkipBlankRows() {
        byte[] csv = "Name\nAlice\n\nBob\nCarol\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .skipBlankRows()
                .maxRows(2)
                .read(new ByteArrayInputStream(csv), result -> names.add(result.data()));

        assertEquals(List.of("Alice", "Bob"), names);
    }

    @Test
    void csvReader_shouldStopAtConsecutiveBlankRows() {
        byte[] csv = "Name\nAlice\n\n\nBob\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .skipBlankRows()
                .stopAtBlankRows(2)
                .read(new ByteArrayInputStream(csv), result -> names.add(result.data()));

        assertEquals(List.of("Alice"), names);
    }
}
