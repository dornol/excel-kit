package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.PrintStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Encrypting holds the whole workbook in heap, so a large one is announced with a warning. */
class EncryptionHeapWarningTest {

    @Test
    void thresholdIsExclusive() {
        assertFalse(ExcelHandler.isTooLargeToEncryptInMemory(ExcelHandler.ENCRYPTION_HEAP_WARNING_BYTES));
        assertTrue(ExcelHandler.isTooLargeToEncryptInMemory(ExcelHandler.ENCRYPTION_HEAP_WARNING_BYTES + 1));
    }

    @Test
    void aSmallEncryptedWorkbook_logsNoWarning() {
        PrintStream original = System.err;
        ByteArrayOutputStream captured = new ByteArrayOutputStream();
        System.setErr(new PrintStream(captured, true));
        try {
            ExcelWriter.<String>create().password("pw").column("A", s -> s).write(List.of("x"))
                    .writeTo(new ByteArrayOutputStream());
        } finally {
            System.setErr(original);
        }
        assertFalse(captured.toString().contains("holds about that much in heap"), captured.toString());
    }
}
