package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelTemplateWriter} to cover:
 * - Invalid sheet index/name
 * - Row order enforcement
 * - checkNotFinished after finish()
 * - setCellValue type dispatch (LocalTime, Boolean, null, fallback)
 * - Writing to template rows vs new rows
 */
class ExcelTemplateWriterEdgeCaseTest {

    private byte[] createTemplate() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            var row0 = sheet.createRow(0);
            row0.createCell(0).setCellValue("Label");
            row0.createCell(1).setCellValue("Value");
            var row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue("Name:");
            sheet.createRow(2).createCell(0).setCellValue("Date:");
            wb.write(out);
        }
        return out.toByteArray();
    }

    private byte[] createMultiSheetTemplate() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("First").createRow(0).createCell(0).setCellValue("A");
            wb.createSheet("Second").createRow(0).createCell(0).setCellValue("B");
            wb.write(out);
        }
        return out.toByteArray();
    }

    // ============================================================
    // Invalid sheet index
    // ============================================================
    @Test
    void sheet_negativeIndex_throws() throws IOException {
        byte[] template = createTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            assertThrows(ExcelWriteException.class, () -> writer.sheet(-1));
        }
    }

    @Test
    void sheet_outOfRangeIndex_throws() throws IOException {
        byte[] template = createTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            assertThrows(ExcelWriteException.class, () -> writer.sheet(5));
        }
    }

    // ============================================================
    // Invalid sheet name
    // ============================================================
    @Test
    void sheet_nonExistentName_throws() throws IOException {
        byte[] template = createTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            assertThrows(ExcelWriteException.class, () -> writer.sheet("NonExistent"));
        }
    }

    @Test
    void sheet_validName_works() throws IOException {
        byte[] template = createMultiSheetTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            assertDoesNotThrow(() -> writer.sheet("Second"));
        }
    }

    // ============================================================
    // Row order enforcement
    // ============================================================
    @Test
    void cell_reverseRowOrder_throws() throws IOException {
        byte[] template = createTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B3", "value1"); // row 2
            assertThrows(ExcelWriteException.class, () -> writer.cell("B1", "value2")); // row 0 < row 2
        }
    }

    @Test
    void cell_sameRow_allowed() throws IOException {
        byte[] template = createTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("A1", "x");
            assertDoesNotThrow(() -> writer.cell("B1", "y")); // same row
        }
    }

    // ============================================================
    // checkNotFinished after finish()
    // ============================================================
    @Test
    void cell_afterFinish_throws() throws IOException {
        byte[] template = createTemplate();
        var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template));
        writer.finish();
        assertThrows(ExcelWriteException.class, () -> writer.cell("A1", "x"));
    }

    @Test
    void list_afterFinish_throws() throws IOException {
        byte[] template = createTemplate();
        var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template));
        writer.finish();
        assertThrows(ExcelWriteException.class, () -> writer.list(5));
    }

    @Test
    void finish_afterFinish_throws() throws IOException {
        byte[] template = createTemplate();
        var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template));
        writer.finish();
        assertThrows(ExcelWriteException.class, () -> writer.finish());
    }

    // ============================================================
    // setCellValue type dispatch — verify actual cell values
    // ============================================================
    @Test
    void cell_withString_shouldWriteStringValue() throws IOException {
        byte[] result = writeTemplateCell("B4", "hello");

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            assertEquals("hello", wb.getSheetAt(0).getRow(3).getCell(1).getStringCellValue());
        }
    }

    @Test
    void cell_withNumber_shouldWriteNumericValue() throws IOException {
        byte[] result = writeTemplateCell("B4", 42.5);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            assertEquals(42.5, wb.getSheetAt(0).getRow(3).getCell(1).getNumericCellValue(), 0.001);
        }
    }

    @Test
    void cell_withBoolean_shouldWriteBooleanValue() throws IOException {
        byte[] result = writeTemplateCell("B4", true);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            assertTrue(wb.getSheetAt(0).getRow(3).getCell(1).getBooleanCellValue());
        }
    }

    @Test
    void cell_withLocalDate_shouldWriteDateValue() throws IOException {
        byte[] result = writeTemplateCell("B4", LocalDate.of(2025, 6, 15));

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(1);
            LocalDate read = cell.getLocalDateTimeCellValue().toLocalDate();
            assertEquals(LocalDate.of(2025, 6, 15), read);
        }
    }

    @Test
    void cell_withLocalDateTime_shouldWriteDateTimeValue() throws IOException {
        LocalDateTime ldt = LocalDateTime.of(2025, 1, 15, 10, 30);
        byte[] result = writeTemplateCell("B4", ldt);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(1);
            assertEquals(ldt, cell.getLocalDateTimeCellValue());
        }
    }

    @Test
    void cell_withLocalTime_shouldWriteAsEpochDate() throws IOException {
        byte[] result = writeTemplateCell("B4", LocalTime.of(14, 30));

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(1);
            // LocalTime is written as LocalTime.atDate(LocalDate.EPOCH)
            LocalDateTime expected = LocalTime.of(14, 30).atDate(LocalDate.EPOCH);
            assertEquals(expected, cell.getLocalDateTimeCellValue());
        }
    }

    @Test
    void cell_withNull_shouldSetBlankCell() throws IOException {
        byte[] result = writeTemplateCell("B4", null);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(1);
            assertEquals(org.apache.poi.ss.usermodel.CellType.BLANK, cell.getCellType());
        }
    }

    @Test
    void cell_withUnknownType_shouldWriteToStringValue() throws IOException {
        byte[] result = writeTemplateCell("B4", new Object() {
            @Override
            public String toString() {
                return "custom-value";
            }
        });

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(result))) {
            assertEquals("custom-value", wb.getSheetAt(0).getRow(3).getCell(1).getStringCellValue());
        }
    }

    private byte[] writeTemplateCell(String cellRef, Object value) throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell(cellRef, value);
            writer.finish().consumeOutputStream(out);
        }
        return out.toByteArray();
    }

    // ============================================================
    // Writing to template rows (row <= templateLastRow)
    // ============================================================
    @Test
    void cell_withinTemplateRow_shouldOverwrite() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            // Row 1 exists in template
            writer.cell("B2", "Alice"); // row 1, col 1
            writer.finish().consumeOutputStream(out);
        }

        // Verify overwritten value
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        }
    }

    // ============================================================
    // Writing beyond template rows (row > templateLastRow)
    // ============================================================
    @Test
    void cell_beyondTemplateRow_shouldCreateNew() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("A10", "New Data"); // row 9, far beyond template
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 100, "Output should contain valid Excel data, got " + out.size() + " bytes");
    }

    // ============================================================
    // close() without finish()
    // ============================================================
    @Test
    void close_withoutFinish_shouldNotThrow() throws IOException {
        byte[] template = createTemplate();
        var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template));
        assertDoesNotThrow(writer::close);
    }

    // ============================================================
    // sheet(int) valid index
    // ============================================================
    @Test
    void sheet_validIndex_works() throws IOException {
        byte[] template = createMultiSheetTemplate();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.sheet(1);
            writer.cell("B1", "data");
            writer.finish();
        }
    }
}
