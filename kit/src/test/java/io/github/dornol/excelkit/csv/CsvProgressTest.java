package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CsvProgressTest {

    // ========================================================================
    // CSV onProgress
    // ========================================================================
    @Test
    void csvProgress_shouldFireAtCorrectIntervals() {
        List<Long> counts = new ArrayList<>();
        CsvWriter.<Integer>create()
                .column("Value", i -> i)
                .onProgress(3, (count, cursor) -> counts.add(count))
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9))
                .writeTo(java.io.OutputStream.nullOutputStream());

        assertEquals(List.of(3L, 6L, 9L), counts);
    }

    @Test
    void csvProgress_shouldNotFireWhenIntervalNotReached() {
        List<Long> counts = new ArrayList<>();
        CsvWriter.<Integer>create()
                .column("Value", i -> i)
                .onProgress(100, (count, cursor) -> counts.add(count))
                .write(Stream.of(1, 2, 3))
                .writeTo(java.io.OutputStream.nullOutputStream());

        assertTrue(counts.isEmpty());
    }
}
