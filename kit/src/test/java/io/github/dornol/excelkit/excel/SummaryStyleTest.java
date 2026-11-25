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
    void anAverageOfWholeNumbersHasDecimalsAndTheOtherOperationsKeepTheColumnFormat() throws Exception {
        var rows = List.of(new Row("A", 1.0, 8), new Row("B", 2.0, 9));
        try (var wb = open(writer().summary(s -> s.label("Total").sum("Qty").average("Qty").min("Qty").max("Qty")).write(rows))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("#,##0", sheet.getRow(3).getCell(2).getCellStyle().getDataFormatString(), "sum");
            assertEquals("#,##0.0", sheet.getRow(4).getCell(2).getCellStyle().getDataFormatString(), "8.5 must not read 9");
            assertEquals("#,##0", sheet.getRow(5).getCell(2).getCellStyle().getDataFormatString(), "min");
            assertEquals("#,##0", sheet.getRow(6).getCell(2).getCellStyle().getDataFormatString(), "max");
            assertEquals(8.5, new org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator(wb)
                    .evaluate(sheet.getRow(4).getCell(2)).getNumberValue());
        }
    }

    @Test
    void aCountIsANumberOfCellsWhateverTheColumnHolds() throws Exception {
        try (var wb = open(writer().summary(s -> s.label("Total").count("Amount").countA("Dept")).write(ROWS))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("#,##0", sheet.getRow(3).getCell(1).getCellStyle().getDataFormatString(), "not #,##0.00 like the amounts");
            assertEquals("#,##0", sheet.getRow(4).getCell(0).getCellStyle().getDataFormatString());
        }
    }

    @Test
    void formatOfTellsWhatEachOperationTakes() {
        var sum = ExcelSummary.Op.SUM;
        assertNull(ExcelSummary.formatOf(sum, "#,##0", 1.0));
        assertNull(ExcelSummary.formatOf(ExcelSummary.Op.MIN, "#,##0.00", 1.0));
        assertEquals("#,##0", ExcelSummary.formatOf(ExcelSummary.Op.COUNT, "#,##0.00", 2.0));
        assertEquals("#,##0", ExcelSummary.formatOf(ExcelSummary.Op.COUNTA, "yyyy-mm-dd", 2.0));
        var avg = ExcelSummary.Op.AVERAGE;
        // the decimals the average has, up to two: none for 3, one for 8.5, two for 8.25
        assertNull(ExcelSummary.formatOf(avg, "#,##0", 3.0), "3 stays #,##0: no .00");
        assertEquals("#,##0.0", ExcelSummary.formatOf(avg, "#,##0", 8.5));
        assertEquals("#,##0.00", ExcelSummary.formatOf(avg, "#,##0", 8.25));
        assertEquals("#,##0.00", ExcelSummary.formatOf(avg, "#,##0", 8.333333333));
        assertNull(ExcelSummary.formatOf(avg, "#,##0", 8.001), "8.001 rounds to 8.00, which is 8");
        assertEquals("0.0", ExcelSummary.formatOf(avg, "0", 2.5));
        assertEquals("0.0%", ExcelSummary.formatOf(avg, "0%", 0.125), "12.5%");
        assertNull(ExcelSummary.formatOf(avg, "0%", 0.12), "12%");
        assertEquals("#,##0.00", ExcelSummary.formatOf(avg, "#,##0", null), "unknown value: two decimals");
        assertNull(ExcelSummary.formatOf(avg, "#,##0.00", 3.0), "decimals already");
        assertNull(ExcelSummary.formatOf(avg, "General", 8.5));
        assertNull(ExcelSummary.formatOf(avg, "yyyy-mm-dd", 8.5));
    }

    @Test
    void styleRefusesNull() {
        assertThrows(NullPointerException.class, () -> new ExcelSummary().style(null));
    }
}
