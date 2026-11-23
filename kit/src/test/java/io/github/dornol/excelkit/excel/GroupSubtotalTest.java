package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class GroupSubtotalTest {
    record Line(String dept, String name, int amount) {}

    private static final List<Line> LINES = List.of(
            new Line("A", "x", 10), new Line("A", "y", 20), new Line("B", "z", 5));

    private static ExcelWriter<Line> writer() {
        return ExcelWriter.<Line>create()
                .column("Dept", Line::dept)
                .column("Name", Line::name)
                .column("Amount", Line::amount)
                .groupSubtotal(Line::dept, s -> s.label("Dept", k -> k + " total").sum("Amount"));
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static String text(XSSFSheet sheet, int row, int col) {
        var cell = sheet.getRow(row).getCell(col);
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case FORMULA -> cell.getCellFormula();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            default -> cell.getStringCellValue();
        };
    }

    @Nested
    class Rows {
        @Test
        void aSubtotalRowFollowsEveryGroup() throws IOException {
            try (var wb = open(writer().write(LINES))) {
                var sheet = wb.getSheetAt(0);
                assertEquals(6, sheet.getLastRowNum() + 1, "header + 3 data + 2 subtotals");
                assertEquals("A total", text(sheet, 3, 0));
                assertEquals("SUBTOTAL(9,C2:C3)", text(sheet, 3, 2));
                assertEquals("B", text(sheet, 4, 0));
                assertEquals("B total", text(sheet, 5, 0));
                assertEquals("SUBTOTAL(9,C5:C5)", text(sheet, 5, 2));
            }
        }

        @Test
        void theSubtotalsCalculateInExcel() throws IOException {
            try (var wb = open(writer().summary(s -> s.label("Total").sum("Amount")).write(LINES))) {
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                var sheet = wb.getSheetAt(0);
                assertEquals(30, evaluator.evaluate(sheet.getRow(3).getCell(2)).getNumberValue());
                assertEquals(5, evaluator.evaluate(sheet.getRow(5).getCell(2)).getNumberValue());
                var summary = sheet.getRow(6).getCell(2);
                assertEquals("SUBTOTAL(9,C2:C6)", summary.getCellFormula());
                assertEquals(35, evaluator.evaluate(summary).getNumberValue(), "subtotals are not counted twice");
            }
        }

        @Test
        void aSubtotalRowIsStyled() throws IOException {
            try (var wb = open(writer().write(LINES))) {
                var sheet = wb.getSheetAt(0);
                var subtotal = sheet.getRow(3).getCell(0).getCellStyle();
                assertTrue(wb.getFontAt(subtotal.getFontIndex()).getBold());
                assertFalse(wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex()).getBold());
            }
        }

        @Test
        void anEmptyStreamHasNoSubtotal() throws IOException {
            try (var wb = open(writer().write(List.<Line>of()))) {
                assertEquals(0, wb.getSheetAt(0).getLastRowNum());
            }
        }

        @Test
        void aKeyThatReturnsGetsASubtotalOfItsOwn() throws IOException {
            var rows = List.of(new Line("A", "x", 1), new Line("B", "y", 2), new Line("A", "z", 3));
            try (var wb = open(writer().write(rows))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("A total", text(sheet, 2, 0));
                assertEquals("B total", text(sheet, 4, 0));
                assertEquals("A total", text(sheet, 6, 0));
            }
        }
    }

    @Nested
    class WithOtherFeatures {
        @Test
        void rowNumbersCountDataRowsOnly() throws IOException {
            var w = ExcelWriter.<Line>create().rowNumberColumn("No")
                    .column("Dept", Line::dept).column("Amount", Line::amount)
                    .groupSubtotal(Line::dept, s -> s.label("Dept", k -> "sub").sum("Amount"));
            try (var wb = open(w.write(LINES))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("1", text(sheet, 1, 0));
                assertEquals("2", text(sheet, 2, 0));
                assertEquals("3", text(sheet, 4, 0));
            }
        }

        @Test
        void aMergedRunDoesNotReachTheSubtotalRow() throws IOException {
            var w = ExcelWriter.<Line>create()
                    .column("Dept", Line::dept, c -> c.mergeSameValues(true))
                    .column("Name", Line::name).column("Amount", Line::amount)
                    .groupSubtotal(Line::dept, s -> s.label("Name", k -> "sub").sum("Amount"));
            try (var wb = open(w.write(LINES))) {
                var regions = wb.getSheetAt(0).getMergedRegions();
                assertEquals(1, regions.size());
                assertEquals(1, regions.get(0).getFirstRow());
                assertEquals(2, regions.get(0).getLastRow());
            }
        }

        @Test
        void pageBreakGoesBelowTheSubtotalRow() throws IOException {
            var w = writer().pageBreakOnChange(Line::dept);
            try (var wb = open(w.write(LINES))) {
                int[] breaks = wb.getSheetAt(0).getRowBreaks();
                assertArrayEquals(new int[]{3}, breaks, "the break is below the subtotal row (index 3)");
            }
        }

        @Test
        void aRolledOverSheetClosesItsGroup() throws IOException {
            var rows = new ArrayList<Line>();
            for (int i = 0; i < 5; i++) rows.add(new Line("A", "n" + i, 1));
            try (var wb = open(writer().maxRows(3).write(rows))) {
                assertEquals(2, wb.getNumberOfSheets());
                var first = wb.getSheetAt(0);
                assertEquals("A total", text(first, 4, 0));
                assertEquals("SUBTOTAL(9,C2:C4)", text(first, 4, 2));
                var second = wb.getSheetAt(1);
                assertEquals("A total", text(second, 3, 0));
                assertEquals("SUBTOTAL(9,C2:C3)", text(second, 3, 2));
            }
        }

        @Test
        void aWorkbookSheetWritesSubtotalsToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Line>sheet("S")
                        .column("Dept", Line::dept).column("Amount", Line::amount)
                        .groupSubtotal(Line::dept, s -> s.label("Dept", k -> "sub").sum("Amount"))
                        .write(LINES);
                try (var wb = open(workbook.finish())) {
                    assertEquals("SUBTOTAL(9,B2:B3)", text(wb.getSheetAt(0), 3, 1));
                }
            }
        }
    }

    @Nested
    class MoreCombinations {
        @Test
        void theAutoFilterCoversTheSubtotalRows() throws IOException {
            try (var wb = open(writer().autoFilter(true).write(LINES))) {
                assertEquals("A1:C6", wb.getSheetAt(0).getCTWorksheet().getAutoFilter().getRef());
            }
        }

        @Test
        void eachRolledOverSheetGetsItsOwnSummaryOverItsOwnRows() throws IOException {
            var rows = new ArrayList<Line>();
            for (int i = 0; i < 4; i++) rows.add(new Line("A", "n" + i, 1));
            try (var wb = open(writer().maxRows(2).summary(s -> s.label("Total").sum("Amount")).write(rows))) {
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                assertEquals(2, wb.getNumberOfSheets());
                for (int i = 0; i < 2; i++) {
                    var sheet = wb.getSheetAt(i);
                    assertEquals("A total", text(sheet, 3, 0));
                    assertEquals("SUBTOTAL(9,C2:C4)", text(sheet, 4, 2), "summary of sheet " + i);
                    assertEquals(2, evaluator.evaluate(sheet.getRow(4).getCell(2)).getNumberValue());
                }
            }
        }

        @Test
        void aGroupThatEndsExactlyAtTheRolloverIsClosedOnce() throws IOException {
            var rows = List.of(new Line("A", "a", 1), new Line("A", "b", 2), new Line("B", "c", 3), new Line("B", "d", 4));
            try (var wb = open(writer().maxRows(2).write(rows))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(4, wb.getSheetAt(0).getLastRowNum() + 1, "header, 2 rows, 1 subtotal");
                assertEquals("A total", text(wb.getSheetAt(0), 3, 0));
                assertEquals(4, wb.getSheetAt(1).getLastRowNum() + 1);
                assertEquals("B total", text(wb.getSheetAt(1), 3, 0));
            }
        }

        @Test
        void afterDataRowsComeBelowTheLastSubtotal() throws IOException {
            var w = writer().afterData(ctx -> {
                ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("footer");
                return ctx.getCurrentRow() + 1;
            });
            try (var wb = open(w.write(LINES))) {
                assertEquals("B total", text(wb.getSheetAt(0), 5, 0));
                assertEquals("footer", text(wb.getSheetAt(0), 6, 0));
            }
        }

        @Test
        void rowStylesStillApplyToDataRowsOnly() throws IOException {
            var w = writer().rowStyle(l -> l.amount() > 10, s -> s.italic(true));
            try (var wb = open(w.write(LINES))) {
                var sheet = wb.getSheetAt(0);
                assertTrue(wb.getFontAt(sheet.getRow(2).getCell(0).getCellStyle().getFontIndex()).getItalic());
                assertFalse(wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex()).getItalic());
            }
        }

        @Test
        void stripesRowColorAndConditionalFormattingDoNotBreakTheWrite() throws IOException {
            var w = writer().stripedRows().rowColor(l -> l.amount() == 5 ? io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_GREEN : null)
                    .freezeRows(1).autoFilter(true)
                    .conditionalFormatting(c -> c.columns(2).greaterThan("10", io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW));
            try (var wb = open(w.write(LINES))) {
                assertEquals(6, wb.getSheetAt(0).getLastRowNum() + 1);
            }
        }

        @Test
        void aNullKeyFormsAGroup() throws IOException {
            var rows = List.of(new Line(null, "a", 1), new Line(null, "b", 2), new Line("A", "c", 3));
            try (var wb = open(writer().write(rows))) {
                assertEquals(" total", text(wb.getSheetAt(0), 3, 0).replace("null", ""));
                assertEquals("A total", text(wb.getSheetAt(0), 5, 0));
            }
        }

        @Test
        void theRowCallbackStillSeesEveryDataRowOnce() throws IOException {
            var seen = new ArrayList<String>();
            try (var wb = open(writer().write(LINES.stream(), (row, cursor) -> seen.add(row.name())))) {
                assertEquals(List.of("x", "y", "z"), seen);
            }
        }

        @Test
        void theProgressCallbackCountsDataRows() throws IOException {
            var counts = new ArrayList<Long>();
            var w = writer().onProgress(1, (total, cursor) -> counts.add(total));
            try (var wb = open(w.write(LINES))) {
                assertEquals(List.of(1L, 2L, 3L), counts);
            }
        }

        @Test
        void aWorkbookWithTwoSubtotalSheetsKeepsThemApart() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                for (String name : List.of("S1", "S2")) {
                    workbook.<Line>sheet(name).column("Dept", Line::dept).column("Amount", Line::amount)
                            .groupSubtotal(Line::dept, s -> s.label("Dept", k -> "sub").sum("Amount"))
                            .write(LINES);
                }
                try (var wb = open(workbook.finish())) {
                    assertEquals("SUBTOTAL(9,B2:B3)", text(wb.getSheetAt(1), 3, 1));
                    assertEquals("SUBTOTAL(9,B2:B3)", text(wb.getSheetAt(0), 3, 1));
                }
            }
        }
    }

    @Nested
    class Rejections {
        @Test
        void aTableIsRejected() {
            assertThrows(ExcelWriteException.class, () -> writer().table("Sales").write(LINES));
        }

        @Test
        void aChartIsRejected() {
            assertThrows(ExcelWriteException.class,
                    () -> writer().chart(c -> c.categoryColumn(0).valueColumn(2, "v")).write(LINES));
        }

        @Test
        void aNamedRangeIsRejected() {
            assertThrows(ExcelWriteException.class, () -> writer().namedRange("N", 2).write(LINES));
        }

        @Test
        void anUnknownColumnIsRejected() {
            var w = ExcelWriter.<Line>create().column("Dept", Line::dept)
                    .groupSubtotal(Line::dept, s -> s.sum("Nope"));
            assertThrows(ExcelWriteException.class, () -> w.write(LINES));
        }

        @Test
        void theLabelColumnCannotBeSubtotaledToo() {
            var w = ExcelWriter.<Line>create().column("Dept", Line::dept).column("Amount", Line::amount)
                    .groupSubtotal(Line::dept, s -> s.label("Amount", k -> "x").sum("Amount"));
            assertThrows(ExcelWriteException.class, () -> w.write(LINES));
        }

        @Test
        void nullArgumentsAreRejected() {
            assertThrows(NullPointerException.class, () -> ExcelWriter.<Line>create().groupSubtotal(null, s -> {}));
            assertThrows(NullPointerException.class, () -> ExcelWriter.<Line>create().groupSubtotal(Line::dept, null));
        }
    }
}
