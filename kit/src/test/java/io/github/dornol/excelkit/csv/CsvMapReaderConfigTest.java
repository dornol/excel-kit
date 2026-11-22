package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class CsvMapReaderConfigTest {

    @Test
    void csvMapReader_dialect_TSV_withCharset() {
        String tsv = "Name\tAge\nAlice\t30\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .dialect(CsvDialect.TSV)
                .charset(StandardCharsets.UTF_8)
                .read(new ByteArrayInputStream(tsv.getBytes()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
    }

    @Test
    void csvMapReader_read_insufficientRowsForHeader_throwsCsvReadException() {
        String csv = "only one line";
        assertThrows(CsvReadException.class, () ->
                CsvReader.forMap()
                        .headerRowIndex(5)
                        .read(new ByteArrayInputStream(csv.getBytes()), r -> {}));
    }
}
