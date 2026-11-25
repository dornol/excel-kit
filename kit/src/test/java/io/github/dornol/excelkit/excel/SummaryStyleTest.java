package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/** The summary rows look like the numbers above them, and like the subtotal rows unless told otherwise. */
class SummaryStyleTest {

    private record Row(String dept, double amount, int qty) {
    }

    private static final List<Row> ROWS = List.of(new Row("A", 1234567.891, 1500000), new Row("B", 2.5, 2500000));

    private static ExcelWriter<Row> writer() {
        return ExcelWriter.<Row>create().column("Dept", Row::dept)
                .column("Amount", Row::amount, c -> c.type(ExcelDataType.DOUBLE))
                .column("Qty", Row::qty, c -> c.type(ExcelDataType.INTEGER));
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws Exception {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static boolean bold(XSSFWorkbook wb, XSSFCell cell) {
        return wb.getFontAt(cell.getCellStyle().getFontIndex()).getBold();
    }

    @Test
    void aSummaryCellTakesTheNumberFormatOfItsColumn() throws Exception {
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Amount").sum("Qty")).write(ROWS))) {
            XSSFSheet sheet = wb.getSheetAt(0);
            var total = sheet.getRow(3);
            assertEquals("Total", total.getCell(0).getStringCellValue());
            assertEquals("#,##0.00", total.getCell(1).getCellStyle().getDataFormatString(), "the Amount cells' format, not General");
            assertEquals("#,##0", total.getCell(2).getCellStyle().getDataFormatString());
            assertEquals(CellType.FORMULA, total.getCell(1).getCellType());
        }
    }

    @Test
    void theDefaultIsBoldOnLightGrayAcrossTheWholeRow() throws Exception {
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Qty")).write(ROWS))) {
            var total = wb.getSheetAt(0).getRow(3);
            for (int col = 0; col < 3; col++) {
                var cell = total.getCell(col);
                assertNotNull(cell, "column " + col + " has a cell, so that the band runs across the row");
                assertTrue(bold(wb, cell), "bold in column " + col);
                assertEquals(FillPatternType.SOLID_FOREGROUND, cell.getCellStyle().getFillPattern(), "filled in column " + col);
            }
            byte[] rgb = ((org.apache.poi.xssf.usermodel.XSSFColor) total.getCell(0).getCellStyle().getFillForegroundColorColor()).getRGB();
            assertEquals((byte) ExcelColor.LIGHT_GRAY.getR(), rgb[0]);
            assertFalse(bold(wb, wb.getSheetAt(0).getRow(1).getCell(0)), "the data rows are not touched");
        }
    }

    @Test
    void styleChangesTheDefaultAndClearingItGivesPlainRows() throws Exception {
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Qty")
                .style(r -> r.backgroundColor(ExcelColor.LIGHT_YELLOW).italic(true))).write(ROWS))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(2);
            byte[] rgb = ((org.apache.poi.xssf.usermodel.XSSFColor) cell.getCellStyle().getFillForegroundColorColor()).getRGB();
            assertEquals((byte) ExcelColor.LIGHT_YELLOW.getR(), rgb[0]);
            assertTrue(wb.getFontAt(cell.getCellStyle().getFontIndex()).getItalic());
        }
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Qty")
                .style(r -> r.bold(false).backgroundColor(null))).write(ROWS))) {
            var cell = wb.getSheetAt(0).getRow(3).getCell(2);
            assertFalse(bold(wb, cell));
            assertEquals(FillPatternType.NO_FILL, cell.getCellStyle().getFillPattern());
            assertEquals("#,##0", cell.getCellStyle().getDataFormatString(), "the number format stays");
        }
    }

    @Test
    void everyRowOfAMultiRowSummaryIsStyled() throws Exception {
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Amount").average("Amount")).write(ROWS))) {
            for (int r = 3; r <= 4; r++) {
                assertTrue(bold(wb, wb.getSheetAt(0).getRow(r).getCell(1)), "row " + r);
            }
        }
    }

    @Test
    void theSummaryCellsAreLockedUnderProtectionEvenInAnEditableColumn() throws Exception {
        var writer = ExcelWriter.<Row>create().defaultStyle(d -> d.locked(false)).protectSheet("pw")
                .column("Dept", Row::dept)
                .column("Amount", Row::amount, c -> c.type(ExcelDataType.DOUBLE))
                .column("Qty", Row::qty, c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Qty"));
        try (var wb = open(writer.write(ROWS))) {
            var sheet = wb.getSheetAt(0);
            assertFalse(sheet.getRow(1).getCell(2).getCellStyle().getLocked(), "a data cell stays editable");
            for (int col = 0; col < 3; col++) {
                assertTrue(sheet.getRow(3).getCell(col).getCellStyle().getLocked(), "summary column " + col);
            }
        }
    }

    @Test
    void eachRolledOverSheetGetsAStyledSummary() throws Exception {
        try (var wb = open(writer().maxRows(1).summary(s -> s.label("Total").sum("Qty")).write(ROWS))) {
            assertEquals(2, wb.getNumberOfSheets());
            for (int s = 0; s < 2; s++) {
                var cell = wb.getSheetAt(s).getRow(2).getCell(2);
                assertEquals(CellType.FORMULA, cell.getCellType());
                assertTrue(bold(wb, cell), "sheet " + s);
                assertEquals("#,##0", cell.getCellStyle().getDataFormatString(), "sheet " + s);
            }
        }
    }

    @Test
    void styleRefusesNull() {
        assertThrows(NullPointerException.class, () -> new ExcelSummary().style(null));
    }
}
