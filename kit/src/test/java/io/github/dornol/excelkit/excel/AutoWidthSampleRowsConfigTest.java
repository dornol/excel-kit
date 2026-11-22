package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class AutoWidthSampleRowsConfigTest {

    // ========================================================================
    // autoWidthSampleRows
    // ========================================================================
    @Test
    void autoWidthSampleRows_shouldBeConfigurable() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .autoWidthSampleRows(5)
                .column("Name", s -> s)
                .write(Stream.of("short", "a very long column value that should affect width"))
                .writeTo(out);

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoWidthSampleRows_zero_shouldDisable() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .autoWidthSampleRows(0)
                .column("Name", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoWidthSampleRows_negative_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                ExcelWriter.<String>create().autoWidthSampleRows(-1));
    }

    @Test
    void autoWidthSampleRows_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .autoWidthSampleRows(10)
                    .column("Name", s -> s)
                    .write(Stream.of("test"));
            wb.finish().writeTo(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }
}
