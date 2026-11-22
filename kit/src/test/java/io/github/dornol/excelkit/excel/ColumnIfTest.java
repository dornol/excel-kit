package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ColumnIfTest {

    // ========================================================================
    // ExcelSheetWriter columnIf()
    // ========================================================================
    @Test
    void columnIf_shouldAddColumnWhenTrue() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .columnIf("Age", true, s -> "30")
                    .write(Stream.of("Alice"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, xwb.getSheetAt(0).getRow(0).getLastCellNum()); // 2 columns
        }
    }

    @Test
    void columnIf_shouldSkipColumnWhenFalse() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .columnIf("Age", false, s -> "30")
                    .write(Stream.of("Alice"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, xwb.getSheetAt(0).getRow(0).getLastCellNum()); // 1 column
        }
    }
}
