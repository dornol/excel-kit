package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class UsabilityImprovementsTest {

    @Test
    void rowDataCopiesInputsAndSupportsOptionalAccess() {
        var cells = new java.util.ArrayList<>(List.of(new CellData(0, "1")));
        var headers = new java.util.ArrayList<>(List.of("Id"));
        var index = new java.util.LinkedHashMap<>(Map.of("Id", 0));
        RowData row = new RowData(cells, headers, index);
        cells.clear();
        headers.set(0, "Changed");
        index.clear();

        assertEquals("1", row.get("Id").asString());
        assertNull(row.getOrNull("Missing"));
        assertTrue(row.get("Id").asOptionalInt().isPresent());
    }

    @Test
    void readerOptionsAndReadAllCompose() {
        List<String> values = CsvReader.<String>mapping(row -> row.get("Name").asString())
                .options(ReaderOptions.builder().skipBlankRows(true).build())
                .readAll(new ByteArrayInputStream("Name\nA\n\nB\n".getBytes(StandardCharsets.UTF_8)));

        assertEquals(List.of("A", "B"), values);
    }

    @Test
    void detectedSourceAndSummaryAreAvailable() {
        InputStreamSource source = () -> new ByteArrayInputStream(
                "Name\nA\n".getBytes(StandardCharsets.UTF_8));
        var summary = CsvReader.<String>mapping(row -> row.get("Name").asString())
                .readDetectedWithSummary(source, result -> {});
        assertEquals(1, summary.successRows());
    }

    @Test
    void duplicateHeaderMessageIncludesPositions() {
        var exception = assertThrows(ExcelKitException.class, () ->
                CsvReader.<String>mapping(row -> row.get("Name").asString())
                        .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                        .read(new ByteArrayInputStream("Name,Name\nA,B\n".getBytes(StandardCharsets.UTF_8)), r -> {}));
        assertNotNull(exception.getCause());
        assertTrue(exception.getCause().getMessage().contains("[0, 1]"));
    }
}
