package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for CSV read progress callback.
 */
class CsvReadProgressTest {

    @Test
    void readProgress_shouldFireAtCorrectIntervals() {
        String csv = "Name\nA\nB\nC\nD\nE\nF\nG\nH\nI\n";
        var is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<Long> counts = new ArrayList<>();
        new CsvReader<>(Holder::new, null)
                .addColumn((h, c) -> h.value = c.asString())
                .onProgress(3, (count, cursor) -> counts.add(count))
                .build(is)
                .read(r -> {});

        assertEquals(List.of(3L, 6L, 9L), counts);
    }

    @Test
    void readProgress_shouldNotFireWhenIntervalNotReached() {
        String csv = "Name\nA\nB\n";
        var is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<Long> counts = new ArrayList<>();
        new CsvReader<>(Holder::new, null)
                .addColumn((h, c) -> h.value = c.asString())
                .onProgress(100, (count, cursor) -> counts.add(count))
                .build(is)
                .read(r -> {});

        assertTrue(counts.isEmpty());
    }

    @Test
    void readProgress_viaReadAsStream_shouldAlsoFire() {
        String csv = "Name\nA\nB\nC\nD\n";
        var is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<Long> counts = new ArrayList<>();
        new CsvReader<>(Holder::new, null)
                .addColumn((h, c) -> h.value = c.asString())
                .onProgress(2, (count, cursor) -> counts.add(count))
                .build(is)
                .readAsStream()
                .forEach(r -> {});

        assertEquals(List.of(2L, 4L), counts);
    }

    @Test
    void readProgress_invalidInterval_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new CsvReader<>(Holder::new, null).onProgress(0, (c, cur) -> {}));
    }

    public static class Holder {
        String value;
    }
}
