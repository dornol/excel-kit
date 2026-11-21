package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.testsupport.TempFiles;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.time.Duration;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.fail;

/**
 * A handler that is dropped without being consumed or discarded releases its temp files once it is
 * garbage collected. The wait is bounded; the collection itself is requested, not guaranteed, so
 * these tests poll instead of assuming a single {@code System.gc()} is enough.
 */
class HandlerCleanerTest {
    private static final Duration LIMIT = Duration.ofSeconds(20);

    private static void awaitLeftoverCount(long expected) throws IOException, InterruptedException {
        long deadline = System.nanoTime() + LIMIT.toNanos();
        while (TempFiles.leftoverCount() != expected) {
            if (System.nanoTime() > deadline) {
                fail("expected " + expected + " temp entries but found " + TempFiles.leftoverCount());
            }
            System.gc();
            Thread.sleep(50);
        }
    }

    @Test
    void anAbandonedExcelHandlerReleasesItsWorkbook() throws Exception {
        long before = TempFiles.leftoverCount();
        ExcelHandler handler = ExcelWriter.<String>create().column("A", s -> s).write(List.of("x"));
        assertEquals(before + 1, TempFiles.leftoverCount(), "the handler owns a temp file while referenced");

        handler = null;

        awaitLeftoverCount(before);
    }

    @Test
    void anAbandonedCsvHandlerDeletesItsStagedFile() throws Exception {
        long before = TempFiles.leftoverCount();
        CsvHandler handler = CsvWriter.<String>create().column("A", s -> s).write(List.of("x"));
        assertEquals(before + 1, TempFiles.leftoverCount(), "the handler owns a temp file while referenced");

        handler = null;

        awaitLeftoverCount(before);
    }
}
