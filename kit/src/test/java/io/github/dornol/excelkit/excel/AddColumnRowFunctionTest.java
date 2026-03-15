package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for ExcelWriter.addColumn with ExcelRowFunction (cursor access).
 */
class AddColumnRowFunctionTest {

    @Test
    void addColumn_withRowFunction_shouldProvideCursorAccess() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("No.", (ExcelRowFunction<String, Object>) (row, cursor) -> cursor.getCurrentTotal(),
                        c -> c.type(ExcelDataType.LONG))
                .addColumn("Name", s -> s)
                .write(Stream.of("Alice", "Bob", "Charlie"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1.0, sheet.getRow(1).getCell(0).getNumericCellValue());
            assertEquals(2.0, sheet.getRow(2).getCell(0).getNumericCellValue());
            assertEquals(3.0, sheet.getRow(3).getCell(0).getNumericCellValue());
            assertEquals("Alice", sheet.getRow(1).getCell(1).getStringCellValue());
        }
    }

    @Test
    void addColumn_withRowFunctionAndConfigurer_shouldApplyType() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("No.", (row, cursor) -> cursor.getCurrentTotal(),
                        c -> c.type(ExcelDataType.LONG))
                .addColumn("Name", s -> s)
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1.0, sheet.getRow(1).getCell(0).getNumericCellValue());
        }
    }
}
