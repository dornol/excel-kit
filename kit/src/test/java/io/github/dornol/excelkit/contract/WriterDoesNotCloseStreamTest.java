package io.github.dornol.excelkit.contract;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The writers read a row stream but never close it; the caller that opened it (a JPA or JDBC
 * stream, for example) closes it. The javadoc of every {@code write(Stream)} says so.
 */
class WriterDoesNotCloseStreamTest {

    private static Stream<String> trackedStream(AtomicBoolean closed) {
        return Stream.of("a", "b").onClose(() -> closed.set(true));
    }

    @Test
    void csvWriterLeavesTheStreamOpen() {
        AtomicBoolean closed = new AtomicBoolean();
        try (Stream<String> rows = trackedStream(closed)) {
            CsvWriter.<String>create().column("v", s -> s).write(rows).writeTo(OutputStream.nullOutputStream());
            assertFalse(closed.get());
        }
        assertTrue(closed.get());
    }

    @Test
    void excelWriterLeavesTheStreamOpen() {
        AtomicBoolean closed = new AtomicBoolean();
        try (Stream<String> rows = trackedStream(closed)) {
            ExcelWriter.<String>create().column("v", s -> s).write(rows).writeTo(OutputStream.nullOutputStream());
            assertFalse(closed.get());
        }
        assertTrue(closed.get());
    }
}
