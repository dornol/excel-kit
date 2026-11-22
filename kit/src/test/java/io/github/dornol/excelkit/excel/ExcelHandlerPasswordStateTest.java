package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** {@link ExcelHandler#isPasswordSet()} tells a caller whether the writer already chose the password. */
class ExcelHandlerPasswordStateTest {

    @Test
    void aWriterPassword_isReported() {
        ExcelHandler handler = ExcelWriter.<String>create().password("writer-level").column("A", s -> s)
                .write(List.of("x"));
        try {
            assertTrue(handler.isPasswordSet());
        } finally {
            handler.discard();
        }
    }

    @Test
    void noWriterPassword_isReportedAsNone_andLateBindingDoesNotChangeThat() {
        ExcelHandler handler = ExcelWriter.<String>create().column("A", s -> s).write(List.of("x"));

        assertFalse(handler.isPasswordSet());
        handler.writeTo(OutputStream.nullOutputStream(), "late");
    }

    @Test
    void aWorkbookPassword_isReported() {
        try (ExcelWorkbook workbook = ExcelWorkbook.create().password("workbook-level")) {
            workbook.<String>sheet("S").column("A", s -> s).write(List.of("x"));
            ExcelHandler handler = workbook.finish();
            try {
                assertTrue(handler.isPasswordSet());
            } finally {
                handler.discard();
            }
        }
    }
}
