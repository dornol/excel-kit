package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Cell-level writes into a template: the formats cells keep or get, the rows they may go to, the values they take. */
class ExcelTemplateWriterDetailsTest {

    /**
     * Row 0 holds a money, a date and a text formatted cell; row 1 a formula; sheet "B" has a title.
     * Everything below row 1 of sheet "A" is new.
     */
    private static byte[] template() throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("A");
            CellStyle money = workbook.createCellStyle();
            money.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00 \"KRW\""));
            CellStyle date = workbook.createCellStyle();
            date.setDataFormat(workbook.createDataFormat().getFormat("dd/mm/yyyy"));
            CellStyle text = workbook.createCellStyle();
            text.setDataFormat(workbook.createDataFormat().getFormat("@"));
            CellStyle bold = workbook.createCellStyle();
            var font = workbook.createFont();
            font.setBold(true);
            bold.setFont(font);
            var row = sheet.createRow(0);
            row.createCell(0).setCellStyle(money);
            row.createCell(1).setCellStyle(date);
            row.createCell(2).setCellStyle(text);
            row.createCell(3).setCellStyle(bold); // General format, but a bold font
            sheet.createRow(1).createCell(0).setCellFormula("1+1");
            workbook.createSheet("B").createRow(0).createCell(0).setCellValue("B title");
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static XSSFWorkbook fill(Consumer<ExcelTemplateWriter> writes) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            writes.accept(writer);
            writer.finish().writeTo(out);
        }
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static XSSFCell cell(XSSFWorkbook workbook, String sheet, int row, int col) {
        return workbook.getSheet(sheet).getRow(row).getCell(col);
    }

    // ---- formats

    @Test
    void aTemplateCellKeepsTheFormatTheTemplateGaveIt() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A1", 1234.5).cell("B1", LocalDate.of(2024, 3, 5))
                .cell("C1", "007").cell("D1", 5))) {
            assertEquals("#,##0.00 \"KRW\"", cell(workbook, "A", 0, 0).getCellStyle().getDataFormatString());
            assertEquals(1234.5, cell(workbook, "A", 0, 0).getNumericCellValue());
            assertEquals("dd/mm/yyyy", cell(workbook, "A", 0, 1).getCellStyle().getDataFormatString());
            assertEquals("@", cell(workbook, "A", 0, 2).getCellStyle().getDataFormatString());
            assertEquals("007", cell(workbook, "A", 0, 2).getStringCellValue());
        }
    }

    @Test
    void aDateWithoutAFormatGetsOneInsteadOfShowingAsANumber() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A5", LocalDate.of(2024, 3, 5))
                .cell("B5", LocalDateTime.of(2024, 3, 5, 13, 45, 30)).cell("C5", LocalTime.of(8, 30)))) {
            assertEquals("yyyy-mm-dd", cell(workbook, "A", 4, 0).getCellStyle().getDataFormatString());
            assertEquals("yyyy-mm-dd hh:mm:ss", cell(workbook, "A", 4, 1).getCellStyle().getDataFormatString());
            assertEquals("hh:mm:ss", cell(workbook, "A", 4, 2).getCellStyle().getDataFormatString());
            assertTrue(DateUtil.isCellDateFormatted(cell(workbook, "A", 4, 0)));
            assertEquals(LocalDate.of(2024, 3, 5), cell(workbook, "A", 4, 0).getLocalDateTimeCellValue().toLocalDate());
        }
    }

    @Test
    void aTemplateCellWithOnlyAFontKeepsItAndGainsTheDateFormat() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("D1", LocalDate.of(2024, 3, 5)))) {
            XSSFCell cell = cell(workbook, "A", 0, 3);
            assertEquals("yyyy-mm-dd", cell.getCellStyle().getDataFormatString());
            assertTrue(cell.getCellStyle().getFont().getBold(), "the template's font stays");
        }
    }

    @Test
    void numbersAndTextBelowTheTemplateStayGeneral() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A5", 42).cell("B5", "text"))) {
            assertEquals("General", cell(workbook, "A", 4, 0).getCellStyle().getDataFormatString());
            assertEquals("General", cell(workbook, "A", 4, 1).getCellStyle().getDataFormatString());
        }
    }

    // ---- values

    @Test
    void nullBlanksACellAndWritingOverAFormulaRemovesIt() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A2", null))) {
            assertEquals(CellType.BLANK, cell(workbook, "A", 1, 0).getCellType());
        }
        try (XSSFWorkbook workbook = fill(w -> w.cell("A2", 7))) {
            assertEquals(CellType.NUMERIC, cell(workbook, "A", 1, 0).getCellType());
            assertEquals(7, cell(workbook, "A", 1, 0).getNumericCellValue());
        }
    }

    @Test
    void textThatLooksLikeAFormulaStaysText() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A5", "=SUM(1,2)").cell("B5", "+1").cell("C5", "@x"))) {
            for (int column = 0; column < 3; column++) {
                assertEquals(CellType.STRING, cell(workbook, "A", 4, column).getCellType());
            }
            assertEquals("=SUM(1,2)", cell(workbook, "A", 4, 0).getStringCellValue());
        }
    }

    @Test
    void valuesExcelCannotHoldAsANumberAreKeptAsText() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A5", Double.NaN).cell("B5", new BigDecimal("12345678901234567890.5"))
                .cell("C5", Long.MAX_VALUE).cell("D5", Double.POSITIVE_INFINITY))) {
            assertEquals("NaN", cell(workbook, "A", 4, 0).getStringCellValue());
            assertEquals("12345678901234567890.5", cell(workbook, "A", 4, 1).getStringCellValue());
            assertEquals(String.valueOf(Long.MAX_VALUE), cell(workbook, "A", 4, 2).getStringCellValue());
            assertEquals("Infinity", cell(workbook, "A", 4, 3).getStringCellValue());
        }
    }

    @Test
    void aDateOutsideTheExcelDateSystemIsKeptAsText() throws Exception {
        try (XSSFWorkbook workbook = fill(w -> w.cell("A5", LocalDate.of(1850, 6, 15)))) {
            assertEquals(CellType.STRING, cell(workbook, "A", 4, 0).getCellType());
            assertEquals("1850-06-15", cell(workbook, "A", 4, 0).getStringCellValue());
        }
    }

    @Test
    void aTextOverTheCellLimitIsRefused() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            assertThrows(IllegalArgumentException.class, () -> writer.cell("A5", "a".repeat(32_768)));
            assertDoesNotThrow(() -> writer.cell("A6", "a".repeat(32_767)));
        }
    }

    // ---- where cells go

    @Test
    void aRefusedWriteLeavesTheRowOrderAsItWas() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            writer.cell("A5", "written");
            assertThrows(IllegalArgumentException.class, () -> writer.cell("A9", "a".repeat(32_768)));
            assertThrows(IllegalArgumentException.class, () -> writer.cell(1_048_576, 0, "off the sheet"));
            assertThrows(IllegalArgumentException.class, () -> writer.list(1_048_576));
            // none of those counts as written: row 6 is still the next row, not "above row 9" or beyond the sheet
            assertDoesNotThrow(() -> writer.cell("A6", "still fine"));
        }
    }

    @Test
    void cellsOfOneRowMayBeWrittenInAnyOrderButRowsOnlyDownwards() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            assertDoesNotThrow(() -> writer.cell("C5", "c").cell("A5", "a").cell("B5", "b"));
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> writer.cell("A4", "late"));
            assertTrue(error.getMessage().contains("top-down"), error.getMessage());
            assertTrue(error.getMessage().contains("Attempted row 3") && error.getMessage().contains("last written row was 4"),
                    error.getMessage());
        }
    }

    @Test
    void eachSheetKeepsItsOwnRowOrder() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            writer.cell("A9", "first sheet, row 9");
            assertDoesNotThrow(() -> writer.sheet("B").cell("A3", "second sheet, row 3"));
            assertDoesNotThrow(() -> writer.sheet(0).cell("A10", "back on the first"));
            assertThrows(ExcelWriteException.class, () -> writer.sheet(0).cell("A2", "above the first sheet's row 10"));
        }
    }

    @Test
    void referencesAndPositionsAreChecked() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            assertThrows(IllegalArgumentException.class, () -> writer.cell("not a reference", "x"));
            assertThrows(IllegalArgumentException.class, () -> writer.cell(-1, 0, "x"));
            assertThrows(IllegalArgumentException.class, () -> writer.cell(5, -1, "x"));
            assertThrows(IllegalArgumentException.class, () -> writer.cell(5, 16_384, "x"));
            assertThrows(IllegalArgumentException.class, () -> writer.cell(1_048_576, 0, "x"));
            assertDoesNotThrow(() -> writer.cell(1_048_575, 16_383, "the last cell"));
        }
    }

    @Test
    void sheetsAreSelectedByIndexOrNameAndOnlyIfTheyExist() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            assertEquals("Sheet index out of range: 2 (workbook has 2 sheets)",
                    assertThrows(ExcelWriteException.class, () -> writer.sheet(2)).getMessage());
            assertThrows(ExcelWriteException.class, () -> writer.sheet(-1));
            assertEquals("Sheet not found: nope", assertThrows(ExcelWriteException.class, () -> writer.sheet("nope")).getMessage());
            assertDoesNotThrow(() -> writer.sheet(1));
            assertDoesNotThrow(() -> writer.sheet("A"));
        }
    }

    @Test
    void aFinishedWriterTakesNothingMore() throws Exception {
        ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()));
        writer.finish().writeTo(new ByteArrayOutputStream());
        assertThrows(ExcelWriteException.class, () -> writer.cell("A5", "x"));
        assertThrows(ExcelWriteException.class, () -> writer.finish());
        assertThrows(ExcelWriteException.class, () -> writer.list(5));
        writer.close();
    }

    @Test
    void theTemplateMustBeAnXlsxAndTheWindowPositive() throws Exception {
        assertThrows(Exception.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream("hello".getBytes())));
        assertThrows(Exception.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(new byte[0])));
        assertThrows(IllegalArgumentException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(template()), 0));
        assertThrows(IllegalArgumentException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(template()), -5));
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()), -1)) {
            assertDoesNotThrow(() -> writer.cell("A5", "unbounded window"));
        }
    }
}
