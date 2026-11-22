package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

class ReadInputLimitTest {

    @Test void inputLimitRejectsOversizedInput() {
        ReadLimitExceededException error = assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(3, -1, -1, -1))
                .read(csv("A\n123\n"), result -> {}));
        assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, error.limit());
        assertEquals(3, error.configured());
        assertTrue(error.actual() > 3);
    }

    @Test void inputLimitStopsBeforeConsumingEntireStream() {
        byte[] content = new byte[1_000_000];
        java.util.concurrent.atomic.AtomicInteger consumed = new java.util.concurrent.atomic.AtomicInteger();
        var input = new ByteArrayInputStream(content) {
            @Override public synchronized int read(byte[] bytes, int off, int len) {
                int read = super.read(bytes, off, len);
                if (read > 0) consumed.addAndGet(read);
                return read;
            }
        };
        assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(100, -1, -1, -1)).read(input, result -> {}));
        assertTrue(consumed.get() < content.length);
    }

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
