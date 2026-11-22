package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertThrows;

/** A writer builds one workbook; a second write must be rejected clearly instead of reusing stale state. */
class ExcelWriterReuseTest {

    @Test
    void secondWriteOnTheSameWriter_isRejected() {
        ExcelWriter<Integer> writer = ExcelWriter.<Integer>create().column("N", n -> n);
        writer.write(List.of(1, 2)).writeTo(OutputStream.nullOutputStream());

        assertThrows(ExcelWriteException.class, () -> writer.write(List.of(3)));
    }

    @Test
    void secondWriteBeforeFirstIsConsumed_isRejected() {
        ExcelWriter<Integer> writer = ExcelWriter.<Integer>create().column("N", n -> n);
        ExcelHandler first = writer.write(List.of(1, 2));
        try {
            assertThrows(ExcelWriteException.class, () -> writer.write(List.of(3)));
        } finally {
            first.writeTo(OutputStream.nullOutputStream());
        }
    }
}
