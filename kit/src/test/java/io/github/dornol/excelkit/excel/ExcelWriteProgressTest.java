package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelWriteProgressTest {

    // ========================================================================
    // Feature 8: Progress callback
    // ========================================================================
    @Test
    void progress_shouldFireAtCorrectIntervals() {
        List<Long> progressCounts = new ArrayList<>();

        ExcelWriter.<Integer>create()
                .column("Value", i -> i, cfg -> cfg.type(ExcelDataType.INTEGER))
                .onProgress(3, (count, cursor) -> progressCounts.add(count))
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
                .writeTo(java.io.OutputStream.nullOutputStream());

        // Should fire at 3, 6, 9
        assertEquals(List.of(3L, 6L, 9L), progressCounts);
    }

    @Test
    void progress_shouldNotFireWhenIntervalNotReached() {
        List<Long> progressCounts = new ArrayList<>();

        ExcelWriter.<Integer>create()
                .column("Value", i -> i, cfg -> cfg.type(ExcelDataType.INTEGER))
                .onProgress(100, (count, cursor) -> progressCounts.add(count))
                .write(Stream.of(1, 2, 3))
                .writeTo(java.io.OutputStream.nullOutputStream());

        assertTrue(progressCounts.isEmpty());
    }

    @Test
    void progress_shouldWorkInExcelSheetWriter() {
        List<Long> progressCounts = new ArrayList<>();

        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Test")
                    .column("Value", i -> i)
                    .onProgress(2, (count, cursor) -> progressCounts.add(count))
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().writeTo(java.io.OutputStream.nullOutputStream());
        }

        assertEquals(List.of(2L, 4L), progressCounts);
    }

    @Test
    void progress_shouldWorkWithColumnBuilderChain() {
        AtomicInteger callCount = new AtomicInteger();

        ExcelWriter.<Integer>create()
                .column("A", i -> i)
                .column("B", i -> i * 2)
                .onProgress(5, (count, cursor) -> callCount.incrementAndGet())
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
                .writeTo(java.io.OutputStream.nullOutputStream());

        assertEquals(2, callCount.get()); // fires at 5 and 10
    }
}
