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
    // setCellValue type dispatch
    // ============================================================
    @Test
    void cell_withLocalTime_shouldWork() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", LocalTime.of(14, 30));
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withLocalDateTime_shouldWork() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", LocalDateTime.of(2025, 1, 15, 10, 30));
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withLocalDate_shouldWork() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", LocalDate.of(2025, 6, 15));
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withBoolean_shouldWork() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", true);
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withNull_shouldSetBlank() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", null);
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withNumber_shouldWork() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", 42.5);
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void cell_withUnknownType_shouldCallToString() throws IOException {
        byte[] template = createTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B1", new Object() {
                @Override
                public String toString() {
                    return "custom-value";
                }
            });
            writer.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
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
        assertTrue(out.size() > 0);
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
