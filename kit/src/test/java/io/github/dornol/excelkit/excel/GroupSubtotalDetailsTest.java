package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

class GroupSubtotalDetailsTest {
    record Row(String dept, String name, int amount, int score) {}

    private static final List<Row> ROWS = List.of(
            new Row("A", "x", 10, 1), new Row("A", "y", 20, 5), new Row("A", "z", 60, 3),
            new Row("B", "w", 4, 9));

    private static ExcelWriter<Row> base() {
        return ExcelWriter.<Row>create()
                .column("Dept", Row::dept)
                .column("Name", Row::name)
                .column("Amount", Row::amount)
                .column("Score", Row::score);
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static double eval(XSSFWorkbook wb, int row, int col) {
        return wb.getCreationHelper().createFormulaEvaluator().evaluate(wb.getSheetAt(0).getRow(row).getCell(col)).getNumberValue();
    }

    private static String formula(XSSFSheet sheet, int row, int col) {
        return sheet.getRow(row).getCell(col).getCellFormula();
    }

    @Nested
    class Functions {
        @Test
        void everyFunctionCalculatesOverItsGroup() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Name", k -> "t")
                    .sum("Amount").average("Score"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(90, eval(wb, 4, 2));
                assertEquals(3, eval(wb, 4, 3), 1e-9);
                assertEquals(4, eval(wb, 6, 2));
                assertEquals(9, eval(wb, 6, 3));
            }
        }

        @Test
        void aColumnGivenTwoFunctionsIsRejected() {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").min("Score").max("Score"));
            assertThrows(ExcelWriteException.class, () -> w.write(ROWS));
        }

        @Test
        void countMinMaxOnTheirOwnColumns() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t")
                    .count("Name").min("Amount").max("Score"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("SUBTOTAL(2,B2:B4)", formula(sheet, 4, 1));
                assertEquals("SUBTOTAL(5,C2:C4)", formula(sheet, 4, 2));
                assertEquals("SUBTOTAL(4,D2:D4)", formula(sheet, 4, 3));
                assertEquals(0, eval(wb, 4, 1), "COUNT counts numbers only: Name is text");
                assertEquals(10, eval(wb, 4, 2));
                assertEquals(5, eval(wb, 4, 3));
            }
        }

        @Test
        void functionNumbersAreExcelsSubtotalCodes() {
            assertEquals(9, ExcelSubtotal.functionNumber(ExcelSummary.Op.SUM));
            assertEquals(1, ExcelSubtotal.functionNumber(ExcelSummary.Op.AVERAGE));
            assertEquals(2, ExcelSubtotal.functionNumber(ExcelSummary.Op.COUNT));
            assertEquals(5, ExcelSubtotal.functionNumber(ExcelSummary.Op.MIN));
            assertEquals(4, ExcelSubtotal.functionNumber(ExcelSummary.Op.MAX));
            for (var op : ExcelSummary.Op.values()) assertDoesNotThrow(() -> ExcelSubtotal.functionNumber(op));
        }

        @Test
        void aColumnCanBeSubtotaledByColumnsOfDifferentFunctionsSideBySide() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount").sum("Score"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(90, eval(wb, 4, 2));
                assertEquals(9, eval(wb, 4, 3));
            }
        }
    }

    @Nested
    class SummaryWithSubtotals {
        @Test
        void everySummaryOperationSwitchesToItsSubtotalCode() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount"))
                    .summary(s -> s.label("All").sum("Amount").average("Amount").count("Amount").min("Amount").max("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                // header, A x3, A total, B x1, B total = 7 rows (0..6); the summary follows at 7..11
                assertEquals("SUBTOTAL(9,C2:C7)", formula(sheet, 7, 2));
                assertEquals("SUBTOTAL(1,C2:C7)", formula(sheet, 8, 2));
                assertEquals("SUBTOTAL(2,C2:C7)", formula(sheet, 9, 2));
                assertEquals("SUBTOTAL(5,C2:C7)", formula(sheet, 10, 2));
                assertEquals("SUBTOTAL(4,C2:C7)", formula(sheet, 11, 2));
                assertEquals(94, eval(wb, 7, 2), "subtotals are skipped by SUBTOTAL");
                assertEquals(4, eval(wb, 9, 2), "COUNT of the four data rows, not the subtotal cells");
                assertEquals(4, eval(wb, 10, 2));
                assertEquals(60, eval(wb, 11, 2));
            }
        }

        @Test
        void aSummaryWithoutSubtotalsKeepsThePlainFunctions() throws IOException {
            var w = base().summary(s -> s.label("All").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("SUM(C2:C5)", formula(wb.getSheetAt(0), 5, 2));
            }
        }
    }

    @Nested
    class Labels {
        @Test
        void theDefaultLabelIsInTheFirstColumnAndNamesTheKey() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("A Subtotal", wb.getSheetAt(0).getRow(4).getCell(0).getStringCellValue());
                assertEquals("B Subtotal", wb.getSheetAt(0).getRow(6).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aNullKeyLeavesJustTheWord() throws IOException {
            var rows = List.of(new Row(null, "x", 1, 1));
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(rows))) {
                assertEquals("Subtotal", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aCustomLabelGetsTheKeyAndCanSitInAnyColumn() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Name", k -> "Sum of " + k).sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var row = wb.getSheetAt(0).getRow(4);
                assertEquals("Sum of A", row.getCell(1).getStringCellValue());
                assertEquals(CellType.BLANK, row.getCell(0).getCellType());
            }
        }

        @Test
        void theKeyCanBeAnyObject() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(r -> r.amount() > 15, s -> s.label("Dept", k -> "big=" + k).sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                // amounts 10, 20, 60, 4 give the keys false, true, true, false
                assertEquals("big=false", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("big=true", sheet.getRow(5).getCell(0).getStringCellValue());
                assertEquals("big=false", sheet.getRow(7).getCell(0).getStringCellValue());
            }
        }
    }

    @Nested
    class Styling {
        private XSSFColor fill(XSSFWorkbook wb, int row, int col) {
            return ((XSSFCellStyle) wb.getSheetAt(0).getRow(row).getCell(col).getCellStyle()).getFillForegroundXSSFColor();
        }

        @Test
        void theDefaultIsBoldOnLightGrayAcrossTheWholeRow() throws IOException {
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(ROWS))) {
                for (int col = 0; col < 4; col++) {
                    var style = wb.getSheetAt(0).getRow(4).getCell(col).getCellStyle();
                    assertTrue(wb.getFontAt(style.getFontIndex()).getBold(), "bold in column " + col);
                    assertArrayEquals(new byte[]{(byte) 217, (byte) 217, (byte) 217}, fill(wb, 4, col).getRGB());
                }
            }
        }

        @Test
        void theStyleCanBeChanged() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.sum("Amount")
                    .style(st -> st.italic(true).backgroundColor(ExcelColor.LIGHT_YELLOW).bold(false)));
            try (var wb = open(w.write(ROWS))) {
                var style = wb.getSheetAt(0).getRow(4).getCell(2).getCellStyle();
                assertTrue(wb.getFontAt(style.getFontIndex()).getItalic());
                assertFalse(wb.getFontAt(style.getFontIndex()).getBold());
                assertArrayEquals(new byte[]{(byte) ExcelColor.LIGHT_YELLOW.getR(), (byte) ExcelColor.LIGHT_YELLOW.getG(),
                        (byte) ExcelColor.LIGHT_YELLOW.getB()}, fill(wb, 4, 2).getRGB());
            }
        }

        @Test
        void aSubtotalKeepsTheNumberFormatOfItsColumn() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .column("Amount", Row::amount, c -> c.format("#,##0.00"))
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("#,##0.00", sheet.getRow(1).getCell(1).getCellStyle().getDataFormatString());
                assertEquals("#,##0.00", sheet.getRow(4).getCell(1).getCellStyle().getDataFormatString());
            }
        }

        @Test
        void dataRowsAreNotTouchedByTheSubtotalStyle() throws IOException {
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(ROWS))) {
                for (int row : new int[]{1, 2, 3, 5}) {
                    var style = wb.getSheetAt(0).getRow(row).getCell(2).getCellStyle();
                    assertFalse(wb.getFontAt(style.getFontIndex()).getBold(), "row " + row);
                    assertNotEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
                }
            }
        }

        @Test
        void aSubtotalRowHasTheRowHeightOfTheSheet() throws IOException {
            try (var wb = open(base().rowHeight(31f).groupSubtotal(Row::dept, s -> s.sum("Amount")).write(ROWS))) {
                assertEquals(31f, wb.getSheetAt(0).getRow(4).getHeightInPoints());
            }
        }

        @Test
        void columnsWithoutAFunctionAreBlank() throws IOException {
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount")).write(ROWS))) {
                var row = wb.getSheetAt(0).getRow(4);
                assertEquals(CellType.BLANK, row.getCell(1).getCellType());
                assertEquals(CellType.BLANK, row.getCell(3).getCellType());
                assertEquals(CellType.FORMULA, row.getCell(2).getCellType());
            }
        }
    }

    @Nested
    class Stripes {
        /** The fill of the first cell of every row, as hex, or "-" without a fill. */
        private List<String> fills(XSSFWorkbook wb, int sheet) {
            var out = new java.util.ArrayList<String>();
            var s = wb.getSheetAt(sheet);
            for (int r = 1; r <= s.getLastRowNum(); r++) {
                var style = (XSSFCellStyle) s.getRow(r).getCell(0).getCellStyle();
                var color = style.getFillPattern() == org.apache.poi.ss.usermodel.FillPatternType.NO_FILL
                        ? null : style.getFillForegroundXSSFColor();
                out.add(color == null ? "-" : String.format("%02X%02X%02X",
                        color.getRGB()[0] & 0xFF, color.getRGB()[1] & 0xFF, color.getRGB()[2] & 0xFF));
            }
            return out;
        }

        @Test
        void dataRowsAlternateAsIfThereWereNoSubtotalsAndSubtotalRowsKeepTheirOwnColor() throws IOException {
            // data rows 1..4 stripe on the 2nd and 4th; the subtotal rows sit after the 3rd and the 4th
            var w = base().stripedRows().groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(List.of("-", "F2F2F2", "-", "D9D9D9", "F2F2F2", "D9D9D9"), fills(wb, 0));
            }
        }

        @Test
        void aRolledOverSheetStartsThePatternAgain() throws IOException {
            var w = base().stripedRows().maxRows(2).groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(List.of("-", "F2F2F2", "D9D9D9"), fills(wb, 0));
                // z (data row 1 of the sheet), the subtotal of A, w (data row 2: striped), the subtotal of B
                assertEquals(List.of("-", "D9D9D9", "F2F2F2", "D9D9D9"), fills(wb, 1));
            }
        }

        @Test
        void aRowColorOrCellColorStillWinsOverTheStripeOnDataRows() throws IOException {
            var w = base().stripedRows().rowColor(r -> r.amount() == 20 ? ExcelColor.LIGHT_GREEN : null)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertNotEquals("F2F2F2", fills(wb, 0).get(1));
                assertNotEquals("-", fills(wb, 0).get(1));
            }
        }

        @Test
        void aSubtotalStyleWithoutABackgroundShowsNoStripe() throws IOException {
            var w = base().stripedRows().groupSubtotal(Row::dept, s -> s.sum("Amount")
                    .style(st -> st.backgroundColor(null)));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("-", fills(wb, 0).get(3), "a subtotal row is never striped");
            }
        }
    }

    @Nested
    class Layout {
        @Test
        void theRangeStartsBelowARowsWrittenBeforeTheHeader() throws IOException {
            var w = base().beforeHeader(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Title");
                        return ctx.getCurrentRow() + 2;
                    })
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Dept", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("SUBTOTAL(9,C4:C6)", formula(sheet, 6, 2));
                assertEquals("SUBTOTAL(9,C8:C8)", formula(sheet, 8, 2));
                assertEquals(90, eval(wb, 6, 2));
            }
        }

        @Test
        void theRangeStartsBelowGroupHeaderRows() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept, c -> c.group("G"))
                    .column("Amount", Row::amount, c -> c.group("G"))
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Dept", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("SUBTOTAL(9,B3:B5)", formula(sheet, 5, 1));
                assertEquals(90, eval(wb, 5, 1));
            }
        }

        @Test
        void aHiddenColumnIsStillSubtotaled() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .column("Amount", Row::amount, c -> c.hidden())
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertTrue(wb.getSheetAt(0).isColumnHidden(1));
                assertEquals(90, eval(wb, 4, 1));
            }
        }

        @Test
        void aSkippedConditionalColumnShiftsTheLetters() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .columnIf("Name", false, Row::name)
                    .column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("SUBTOTAL(9,B2:B4)", formula(wb.getSheetAt(0), 4, 1));
            }
        }

        @Test
        void aColumnBeyondZIsLetteredCorrectly() throws IOException {
            var w = ExcelWriter.<Row>create();
            for (int i = 0; i < 27; i++) w.column("c" + i, r -> r.dept());
            w.column("Amount", Row::amount).groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("SUBTOTAL(9,AB2:AB4)", formula(wb.getSheetAt(0), 4, 27));
                assertEquals(90, eval(wb, 4, 27));
            }
        }

        @Test
        void theFreezePaneAndFilterStayOnTheHeader() throws IOException {
            var w = base().freezeRows(1).autoFilter(true).groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals(1, sheet.getPaneInformation().getHorizontalSplitPosition());
                assertEquals("A1:D7", sheet.getCTWorksheet().getAutoFilter().getRef());
            }
        }

        @Test
        void rowBreakEveryCountsSheetRows() throws IOException {
            var w = base().printSetup(p -> p.rowBreakEvery(2)).groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                // rows 0..6: a break below index 2 and 4
                assertArrayEquals(new int[]{2, 4}, wb.getSheetAt(0).getRowBreaks());
            }
        }

        @Test
        void aCustomPageBreakConditionBreaksBelowTheSubtotal() throws IOException {
            var w = base().pageBreakWhen((prev, cur) -> !prev.dept().equals(cur.dept()))
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertArrayEquals(new int[]{4}, wb.getSheetAt(0).getRowBreaks());
            }
        }
    }

    @Nested
    class Grouping {
        @Test
        void everyRowItsOwnGroup() throws IOException {
            var rows = List.of(new Row("A", "1", 1, 0), new Row("B", "2", 2, 0), new Row("C", "3", 3, 0));
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(rows))) {
                var sheet = wb.getSheetAt(0);
                assertEquals(7, sheet.getLastRowNum() + 1);
                assertEquals("SUBTOTAL(9,C2:C2)", formula(sheet, 2, 2));
                assertEquals("SUBTOTAL(9,C4:C4)", formula(sheet, 4, 2));
                assertEquals("SUBTOTAL(9,C6:C6)", formula(sheet, 6, 2));
            }
        }

        @Test
        void oneGroupForAllRows() throws IOException {
            var rows = IntStream.range(0, 50).mapToObj(i -> new Row("A", "n" + i, 2, 0)).toList();
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(rows))) {
                assertEquals("SUBTOTAL(9,C2:C51)", formula(wb.getSheetAt(0), 51, 2));
                assertEquals(100, eval(wb, 51, 2));
                assertEquals(52, wb.getSheetAt(0).getLastRowNum() + 1);
            }
        }

        @Test
        void aKeyThatComesBackStartsANewGroup() throws IOException {
            var rows = List.of(new Row("A", "1", 1, 0), new Row("B", "2", 2, 0), new Row("A", "3", 3, 0));
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(rows))) {
                assertEquals(1, eval(wb, 2, 2));
                assertEquals(2, eval(wb, 4, 2));
                assertEquals(3, eval(wb, 6, 2));
            }
        }

        @Test
        void keysAreComparedByEqualsNotByIdentity() throws IOException {
            var rows = List.of(new Row(new String("A"), "1", 1, 0), new Row(new String("A"), "2", 2, 0),
                    new Row(new String("A"), "3", 3, 0));
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(rows))) {
                assertEquals(5, wb.getSheetAt(0).getLastRowNum() + 1, "header, 3 rows, one subtotal");
                assertEquals(6, eval(wb, 4, 2));
            }
        }

        @Test
        void equalBoxedNumbersFormOneGroup() throws IOException {
            var rows = List.of(new Row("a", "1", 1, 0), new Row("b", "2", 2, 0), new Row("c", "3", 3, 0));
            int[] n = {1000, 1000, 1000};
            int[] i = {0};
            var w = base().groupSubtotal(r -> Integer.valueOf(n[i[0]++]), s -> s.sum("Amount"));
            try (var wb = open(w.write(rows))) {
                assertEquals(5, wb.getSheetAt(0).getLastRowNum() + 1);
                assertEquals(6, eval(wb, 4, 2));
            }
        }

        @Test
        void aGroupSubtotalIsNotWrittenForAnEmptyStream() throws IOException {
            try (var wb = open(base().groupSubtotal(Row::dept, s -> s.sum("Amount")).write(List.<Row>of()))) {
                assertEquals(0, wb.getSheetAt(0).getLastRowNum());
            }
        }

        @Test
        void aStreamIsReadOnceAndKeysAreComputedOncePerRow() throws IOException {
            int[] calls = {0};
            var w = base().groupSubtotal(r -> { calls[0]++; return r.dept(); }, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(ROWS.size(), calls[0]);
            }
        }
    }

    @Nested
    class Rollover {
        @Test
        void maxRowsOfOneGivesEveryRowItsOwnSheetWithASubtotal() throws IOException {
            try (var wb = open(base().maxRows(1).groupSubtotal(Row::dept, s -> s.sum("Amount")).write(ROWS))) {
                assertEquals(4, wb.getNumberOfSheets());
                for (int i = 0; i < 4; i++) {
                    assertEquals(3, wb.getSheetAt(i).getLastRowNum() + 1, "sheet " + i);
                    assertEquals("SUBTOTAL(9,C2:C2)", formula(wb.getSheetAt(i), 2, 2));
                }
            }
        }

        @Test
        void rowNumbersContinueOverSheetsAndSkipSubtotals() throws IOException {
            var w = ExcelWriter.<Row>create().rowNumberColumn("No").column("Dept", Row::dept).column("Amount", Row::amount)
                    .maxRows(2).groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(1, wb.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue());
                assertEquals(2, wb.getSheetAt(0).getRow(2).getCell(0).getNumericCellValue());
                assertEquals(3, wb.getSheetAt(1).getRow(1).getCell(0).getNumericCellValue());
                assertEquals(4, wb.getSheetAt(1).getRow(3).getCell(0).getNumericCellValue());
            }
        }

        @Test
        void aMergedKeyColumnStaysWithinItsSheet() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept, c -> c.mergeSameValues())
                    .column("Name", Row::name).column("Amount", Row::amount)
                    .maxRows(2).groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var first = wb.getSheetAt(0).getMergedRegions();
                assertEquals(1, first.size());
                assertEquals(1, first.get(0).getFirstRow());
                assertEquals(2, first.get(0).getLastRow());
                assertEquals(0, wb.getSheetAt(1).getMergedRegions().size(), "z alone and then w alone: no run");
            }
        }

        @Test
        void aRolledOverSheetHasItsOwnSheetNameAndSubtotals() throws IOException {
            var w = base().sheetName("Data").maxRows(3).groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("Data", wb.getSheetName(0));
                assertEquals("Data (2)", wb.getSheetName(1));
            }
        }

        @Test
        void aSheetThatCannotEvenHoldItsFirstSubtotalRowFailsClearly() {
            // the preamble leaves one row below the header: the first data row fits, its subtotal row does not
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .beforeHeader(ctx -> 1_048_576 - 2)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(List.of(new Row("A", "n", 1, 0))));
            assertTrue(e.getMessage().contains("no room for the subtotal row"), e.getMessage());
        }
    }

    @Nested
    class Rejections {
        private ExcelWriter<Row> writer() {
            return ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
        }

        @Test
        void tableChartAndNamedRangeFailBeforeAnythingIsWritten() {
            assertThrows(ExcelWriteException.class, () -> writer().table("Sales").write(ROWS));
            assertThrows(ExcelWriteException.class,
                    () -> writer().chart(c -> c.categoryColumn(0).valueColumn(1, "v")).write(ROWS));
            assertThrows(ExcelWriteException.class, () -> writer().namedRange("Amounts", 1).write(ROWS));
        }

        @Test
        void aWorkbookSheetRejectsThemToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.sum("Amount")).table("Sales");
                assertThrows(ExcelWriteException.class, () -> sheet.write(ROWS));
            }
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.sum("Amount")).chart(c -> c.categoryColumn(0).valueColumn(1, "v"));
                assertThrows(ExcelWriteException.class, () -> sheet.write(ROWS));
            }
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.sum("Amount")).namedRange("Amounts", 1);
                assertThrows(ExcelWriteException.class, () -> sheet.write(ROWS));
            }
        }

        @Test
        void aSheetWithoutSubtotalsStillTakesTableChartAndNamedRange() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .namedRange("Amounts", 1).chart(c -> c.categoryColumn(0).valueColumn(1, "v"));
            w.write(ROWS).writeTo(OutputStream.nullOutputStream());
            ExcelWriter.<Row>create().column("Dept", Row::dept).table("Sales").write(ROWS).writeTo(OutputStream.nullOutputStream());
        }

        @Test
        void unknownAndClashingColumnsAreNamed() {
            var unknownOp = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .groupSubtotal(Row::dept, s -> s.sum("Nope"));
            assertTrue(assertThrows(ExcelWriteException.class, () -> unknownOp.write(ROWS)).getMessage().contains("Nope"));

            var unknownLabel = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .groupSubtotal(Row::dept, s -> s.label("Nope", k -> "x"));
            assertTrue(assertThrows(ExcelWriteException.class, () -> unknownLabel.write(ROWS)).getMessage().contains("Nope"));

            var clash = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Amount", k -> "x").sum("Amount"));
            assertThrows(ExcelWriteException.class, () -> clash.write(ROWS));

            var twice = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount").average("Amount"));
            assertThrows(ExcelWriteException.class, () -> twice.write(ROWS));

            var firstColumnIsTheDefaultLabel = ExcelWriter.<Row>create().column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            assertThrows(ExcelWriteException.class, () -> firstColumnIsTheDefaultLabel.write(ROWS));
        }

        @Test
        void aSecondCallReplacesTheFirst() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.sum("Amount"))
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "second").sum("Score"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals("second", wb.getSheetAt(0).getRow(4).getCell(0).getStringCellValue());
                assertEquals(CellType.BLANK, wb.getSheetAt(0).getRow(4).getCell(2).getCellType());
                assertEquals(CellType.FORMULA, wb.getSheetAt(0).getRow(4).getCell(3).getCellType());
            }
        }

        @Test
        void nullArgumentsAreRejectedByTheConfiguration() {
            var subtotal = new ExcelSubtotal();
            assertThrows(NullPointerException.class, () -> subtotal.sum(null));
            assertThrows(NullPointerException.class, () -> subtotal.average(null));
            assertThrows(NullPointerException.class, () -> subtotal.count(null));
            assertThrows(NullPointerException.class, () -> subtotal.min(null));
            assertThrows(NullPointerException.class, () -> subtotal.max(null));
            assertThrows(NullPointerException.class, () -> subtotal.label(null, k -> "x"));
            assertThrows(NullPointerException.class, () -> subtotal.label("A", null));
            assertThrows(NullPointerException.class, () -> subtotal.style(null));
        }
    }
}
