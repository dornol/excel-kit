package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class AutoFilterTest {

    // ========================================================================
    // ExcelSheetWriter autoFilter(boolean)
    // ========================================================================
    @Test
    void autoFilterBoolean_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .autoFilter(true)
                    .column("Name", s -> s)
                    .write(Stream.of("A", "B"));
            wb.finish().writeTo(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoFilterFalse_shouldNotApplyFilter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .autoFilter(false)
                    .column("Name", s -> s)
                    .write(Stream.of("A"));
            wb.finish().writeTo(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }
}
