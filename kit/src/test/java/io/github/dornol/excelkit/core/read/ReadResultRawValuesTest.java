package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ReadResultRawValuesTest {

    @Test
    void readResult_shouldCarryRawValues() {
        byte[] csv = "Name,Age\nAlice,30\n".getBytes(StandardCharsets.UTF_8);
        List<List<String>> rawRows = new ArrayList<>();

        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .read(new ByteArrayInputStream(csv), result -> rawRows.add(result.rawValues()));

        assertEquals(List.of(List.of("Alice", "30")), rawRows);
    }
}
