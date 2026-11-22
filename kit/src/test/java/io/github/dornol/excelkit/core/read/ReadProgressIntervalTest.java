package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ReadProgressIntervalTest {

    @Test void detailedProgressReportsCounts() {
        var progress = new java.util.ArrayList<ReadProgress>();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .onReadProgress(10, progress::add)
                .read(csv("Name\nA\n"), result -> {});
        assertEquals(1, progress.size(), "completion event is required even below interval");
        assertTrue(progress.get(0).completed());
        assertEquals(1, progress.get(0).successRows());
    }

    @Test void legacyAndDetailedProgressIntervalsAreIndependent() {
        var legacy = new java.util.ArrayList<Long>();
        var detailed = new java.util.ArrayList<ReadProgress>();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .onProgress(2, (count, cursor) -> legacy.add(count))
                .onReadProgress(3, detailed::add)
                .read(csv("Name\nA\nB\nC\nD\nE\nF\n"), result -> {});

        assertEquals(java.util.List.of(2L, 4L, 6L), legacy);
        assertEquals(2, detailed.stream().filter(progress -> !progress.completed()).count());
    }

    record Person(String name) {}

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
