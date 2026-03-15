package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for Formula and Hyperlink data types.
 */
class FormulaAndHyperlinkTest {

    @Test
    void formula_shouldWriteFormulaCells() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<TestRow>()
                .column("Value", TestRow::getValue)
                    .type(ExcelDataType.INTEGER)
                .column("Formula", TestRow::getFormula)
                    .type(ExcelDataType.FORMULA)
                .write(Stream.of(
                        new TestRow(100, "A2*2"),
                        new TestRow(200, "A3*2")
                ))
                .consumeOutputStream(out);

        // Read back with XSSFWorkbook to verify formulas
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);

            // Row 1 (data row 0) - header is row 0
            Row row1 = sheet.getRow(1);
            Cell formulaCell1 = row1.getCell(1);
            assertEquals(CellType.FORMULA, formulaCell1.getCellType());
            assertEquals("A2*2", formulaCell1.getCellFormula());

            Row row2 = sheet.getRow(2);
            Cell formulaCell2 = row2.getCell(1);
            assertEquals(CellType.FORMULA, formulaCell2.getCellType());
            assertEquals("A3*2", formulaCell2.getCellFormula());
        }
    }

    @Test
    void formula_inAfterData_shouldWriteSummaryRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<TestRow>()
                .column("Value", TestRow::getValue)
                    .type(ExcelDataType.INTEGER)
                .afterData(ctx -> {
                    var sheet = ctx.getSheet();
                    var row = sheet.createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellFormula(
                            "SUM(" + SheetContext.columnLetter(0) + "2:" +
                                    SheetContext.columnLetter(0) + ctx.getCurrentRow() + ")");
                    return ctx.getCurrentRow() + 1;
                })
                .write(Stream.of(new TestRow(100, null), new TestRow(200, null)))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Summary row should be at row index 3 (header=0, data=1,2, summary=3)
            Row summaryRow = sheet.getRow(3);
            assertNotNull(summaryRow);
            Cell sumCell = summaryRow.getCell(0);
            assertEquals(CellType.FORMULA, sumCell.getCellType());
            assertEquals("SUM(A2:A3)", sumCell.getCellFormula());
        }
    }

    @Test
    void hyperlink_shouldWriteHyperlinkCells() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<TestLink>()
                .column("Name", TestLink::getName)
                .column("URL", TestLink::getUrl)
                    .type(ExcelDataType.HYPERLINK)
                .write(Stream.of(
                        new TestLink("Google", "https://google.com"),
                        new TestLink("GitHub", "https://github.com")
                ))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);

            Row row1 = sheet.getRow(1);
            Cell urlCell1 = row1.getCell(1);
            assertEquals("https://google.com", urlCell1.getStringCellValue());
            assertNotNull(urlCell1.getHyperlink());
            assertEquals("https://google.com", urlCell1.getHyperlink().getAddress());

            Row row2 = sheet.getRow(2);
            Cell urlCell2 = row2.getCell(1);
            assertEquals("https://github.com", urlCell2.getStringCellValue());
            assertNotNull(urlCell2.getHyperlink());
            assertEquals("https://github.com", urlCell2.getHyperlink().getAddress());
        }
    }

    @Test
    void hyperlink_withExcelHyperlink_shouldUseSeparateLabelAndUrl() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<TestLink>()
                .column("Name", TestLink::getName)
                .column("Link", t -> new ExcelHyperlink(t.getUrl(), "Click Here"))
                    .type(ExcelDataType.HYPERLINK)
                .write(Stream.of(new TestLink("Google", "https://google.com")))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row1 = sheet.getRow(1);
            Cell linkCell = row1.getCell(1);
            assertEquals("Click Here", linkCell.getStringCellValue());
            assertNotNull(linkCell.getHyperlink());
            assertEquals("https://google.com", linkCell.getHyperlink().getAddress());
        }
    }

    @Test
    void columnLetter_shouldConvertCorrectly() {
        assertEquals("A", SheetContext.columnLetter(0));
        assertEquals("B", SheetContext.columnLetter(1));
        assertEquals("Z", SheetContext.columnLetter(25));
        assertEquals("AA", SheetContext.columnLetter(26));
        assertEquals("AB", SheetContext.columnLetter(27));
        assertEquals("AZ", SheetContext.columnLetter(51));
        assertEquals("BA", SheetContext.columnLetter(52));
    }

    public static class TestRow {
        private final int value;
        private final String formula;

        public TestRow(int value, String formula) {
            this.value = value;
            this.formula = formula;
        }

        public int getValue() { return value; }
        public String getFormula() { return formula; }
    }

    public static class TestLink {
        private final String name;
        private final String url;

        public TestLink(String name, String url) {
            this.name = name;
            this.url = url;
        }

        public String getName() { return name; }
        public String getUrl() { return url; }
    }
}
