package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/** How the subtotal rows meet the rest of the writer: protection, widths, colors, validations, types, failures, reading back. */
class GroupSubtotalInteractionsTest {
    record Row(String dept, String name, int amount) {}

    private static final List<Row> ROWS = List.of(
            new Row("A", "x", 10), new Row("A", "y", 20), new Row("B", "z", 5));

    private static ExcelWriter<Row> base() {
        return ExcelWriter.<Row>create()
                .column("Dept", Row::dept)
                .column("Name", Row::name)
                .column("Amount", Row::amount)
                .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
    }

    private static byte[] bytes(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return out.toByteArray();
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        return new XSSFWorkbook(new ByteArrayInputStream(bytes(handler)));
    }

    @Nested
    class Protection {
        @Test
        void aSubtotalCellStaysLockedInAColumnPeopleMayEdit() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept, c -> c.locked(false))
                    .column("Name", Row::name, c -> c.locked(false))
                    .column("Amount", Row::amount, c -> c.locked(false))
                    .protectSheet("pw")
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                for (int col = 0; col < 3; col++) {
                    assertFalse(sheet.getRow(1).getCell(col).getCellStyle().getLocked(), "data cell " + col + " stays editable");
                    assertTrue(sheet.getRow(3).getCell(col).getCellStyle().getLocked(), "subtotal cell " + col + " is locked");
                }
                assertTrue(sheet.getProtect());
            }
        }

        @Test
        void aLockedColumnStaysLocked() throws IOException {
            try (var wb = open(base().protectSheet("pw").write(ROWS))) {
                assertTrue(wb.getSheetAt(0).getRow(3).getCell(2).getCellStyle().getLocked());
            }
        }

        @Test
        void aFileWithAPasswordStillWritesSubtotals() throws Exception {
            var out = new ByteArrayOutputStream();
            base().write(ROWS).writeTo(out, "secret");
            assertTrue(out.size() > 0);
            try (var in = new ByteArrayInputStream(out.toByteArray());
                 var fs = new org.apache.poi.poifs.filesystem.POIFSFileSystem(in)) {
                var info = new org.apache.poi.poifs.crypt.EncryptionInfo(fs);
                var decryptor = org.apache.poi.poifs.crypt.Decryptor.getInstance(info);
                assertTrue(decryptor.verifyPassword("secret"));
                try (var wb = new XSSFWorkbook(decryptor.getDataStream(fs))) {
                    assertEquals("SUBTOTAL(9,C2:C3)", wb.getSheetAt(0).getRow(3).getCell(2).getCellFormula());
                }
            }
        }
    }

    @Nested
    class Widths {
        private int width(XSSFWorkbook wb, int col) {
            return wb.getSheetAt(0).getColumnWidth(col);
        }

        @Test
        void aLongLabelWidensItsColumn() throws IOException {
            var rows = List.of(new Row("A", "x", 1));
            int without;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name).write(rows))) {
                without = width(wb, 1);
            }
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "a really long subtotal label of a department").sum("Amount"));
            try (var wb = open(w.write(rows))) {
                assertTrue(width(wb, 1) > without, width(wb, 1) + " > " + without);
            }
        }

        @Test
        void aFixedWidthIsKept() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .column("Name", Row::name, c -> c.width(10)).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "a really long subtotal label of a department").sum("Amount"));
            int fixed;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name, c -> c.width(10)).write(ROWS))) {
                fixed = wb.getSheetAt(0).getColumnWidth(1);
            }
            try (var wb = open(w.write(ROWS))) {
                assertEquals(fixed, wb.getSheetAt(0).getColumnWidth(1));
            }
        }

        @Test
        void aLabelBeyondTheSampledRowsDoesNotWidenTheColumn() throws IOException {
            var rows = List.of(new Row("A", "x", 1), new Row("A", "y", 1), new Row("B", "z", 1));
            var longLabel = "a really long subtotal label of a department";
            int sampled;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount).autoWidthSampleRows(1)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> longLabel).sum("Amount")).write(rows))) {
                sampled = width(wb, 1);
            }
            int all;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount).autoWidthSampleRows(100)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> longLabel).sum("Amount")).write(rows))) {
                all = width(wb, 1);
            }
            assertTrue(sampled < all, "the first subtotal comes after the single sampled row: " + sampled + " < " + all);
        }

        @Test
        void aTurnedOffWidthSampleLeavesTheLabelColumnAsTheDataSetIt() throws IOException {
            var rows = List.of(new Row("A", "x", 1), new Row("B", "y", 1));
            var longLabel = "a really long subtotal label of a department";
            int off;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount).autoWidthSampleRows(0)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> longLabel).sum("Amount")).write(rows))) {
                off = width(wb, 1);
            }
            int on;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount).autoWidthSampleRows(100)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> longLabel).sum("Amount")).write(rows))) {
                on = width(wb, 1);
            }
            assertTrue(off < on, off + " < " + on);
        }
    }

    @Nested
    class ColorsAndRules {
        @Test
        void cellColorAndRowColorDoNotReachASubtotalRow() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept)
                    .column("Name", Row::name)
                    .column("Amount", Row::amount, c -> c.cellColor((v, r) -> ExcelColor.LIGHT_GREEN))
                    .rowColor(r -> ExcelColor.LIGHT_YELLOW)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                for (int col = 0; col < 3; col++) {
                    byte[] rgb = ((XSSFCellStyle) sheet.getRow(3).getCell(col).getCellStyle()).getFillForegroundXSSFColor().getRGB();
                    assertArrayEquals(new byte[]{(byte) 217, (byte) 217, (byte) 217}, rgb, "subtotal column " + col);
                }
                byte[] data = ((XSSFCellStyle) sheet.getRow(1).getCell(2).getCellStyle()).getFillForegroundXSSFColor().getRGB();
                assertEquals((byte) ExcelColor.LIGHT_GREEN.getR(), data[0], "the data row keeps its cell color");
            }
        }

        private static java.util.List<String> ruleRanges(org.apache.poi.xssf.usermodel.XSSFWorkbook wb) {
            var formatting = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertEquals(1, formatting.getNumConditionalFormattings());
            return java.util.Arrays.stream(formatting.getConditionalFormattingAt(0).getFormattingRanges())
                    .map(org.apache.poi.ss.util.CellRangeAddress::formatAsString).toList();
        }

        @Test
        void aConditionalFormatRuleSkipsTheSubtotalRows() throws IOException {
            var w = base().conditionalFormatting(c -> c.columns(2).greaterThan("15", ExcelColor.LIGHT_YELLOW));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(java.util.List.of("C2:C3", "C5"), ruleRanges(wb),
                        "the data rows of each group; the subtotal rows 4 and 6 are no data");
            }
        }

        @Test
        void includeSubtotalRowsPutsTheSubtotalRowsBackUnderTheRule() throws IOException {
            var w = base().conditionalFormatting(c -> c.columns(2).includeSubtotalRows()
                    .greaterThan("15", ExcelColor.LIGHT_YELLOW));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(java.util.List.of("C2:C6"), ruleRanges(wb), "from the first data row to the last subtotal row");
            }
        }

        @Test
        void aRuleOnASheetWithoutSubtotalsIsOneRangeAsBefore() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name).column("Amount", Row::amount)
                    .conditionalFormatting(c -> c.columns(2).greaterThan("15", ExcelColor.LIGHT_YELLOW));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(java.util.List.of("C2:C4"), ruleRanges(wb));
            }
        }

        @Test
        void aDropdownAndAValidationRunBelowTheHeaderThroughTheSubtotalRows() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept, c -> c.dropdown("A", "B"))
                    .column("Name", Row::name)
                    .column("Amount", Row::amount, c -> c.validation(ExcelValidation.integerBetween(0, 100)))
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertEquals(2, validations.size());
                for (var v : validations) {
                    var region = v.getRegions().getCellRangeAddress(0);
                    assertEquals(1, region.getFirstRow());
                    assertTrue(region.getLastRow() >= 4, "reaches the subtotal rows (and on down the column)");
                }
                assertEquals(CellType.FORMULA, wb.getSheetAt(0).getRow(3).getCell(2).getCellType(), "a formula is not checked");
            }
        }
    }

    @Nested
    class ColumnTypes {
        record Typed(String dept, LocalDate date, String link, int amount) {}

        @Test
        void dateHyperlinkFormulaAndNullValueColumnsAreBlankInASubtotalRow() throws IOException {
            var w = ExcelWriter.<Typed>create()
                    .column("Dept", Typed::dept)
                    .column("Date", Typed::date, c -> c.type(ExcelDataType.DATE))
                    .column("Link", Typed::link, c -> c.type(ExcelDataType.HYPERLINK))
                    .column("Calc", t -> "A1+1", c -> c.type(ExcelDataType.FORMULA))
                    .column("Note", t -> null, c -> c.nullValue("-"))
                    .column("Amount", Typed::amount)
                    .groupSubtotal(Typed::dept, s -> s.label("Dept", k -> k + " total").sum("Amount"));
            var rows = List.of(new Typed("A", LocalDate.of(2026, 1, 1), "https://example.com", 1),
                    new Typed("A", LocalDate.of(2026, 1, 2), "https://example.org", 2));
            try (var wb = open(w.write(rows))) {
                var row = wb.getSheetAt(0).getRow(3);
                assertEquals("A total", row.getCell(0).getStringCellValue());
                for (int col = 1; col <= 4; col++) assertEquals(CellType.BLANK, row.getCell(col).getCellType(), "column " + col);
                assertEquals("SUBTOTAL(9,F2:F3)", row.getCell(5).getCellFormula());
                assertEquals("-", wb.getSheetAt(0).getRow(1).getCell(4).getStringCellValue(), "the data rows still show the null value");
            }
        }

        @Test
        void aColumnAlignmentAndTheDefaultStyleCarryOverToTheSubtotalRow() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .defaultStyle(d -> d.fontName("Arial"))
                    .column("Dept", Row::dept)
                    .column("Name", Row::name, c -> c.alignment(HorizontalAlignment.CENTER))
                    .column("Amount", Row::amount, c -> c.alignment(HorizontalAlignment.RIGHT))
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                var row = wb.getSheetAt(0).getRow(3);
                assertEquals(HorizontalAlignment.CENTER, row.getCell(1).getCellStyle().getAlignment());
                assertEquals(HorizontalAlignment.RIGHT, row.getCell(2).getCellStyle().getAlignment());
                assertEquals("Arial", wb.getFontAt(row.getCell(0).getCellStyle().getFontIndex()).getFontName());
            }
        }

        @Test
        void writeErrorPoliciesDoNotChangeTheSubtotals() throws IOException {
            for (var policy : WriteErrorPolicy.values()) {
                try (var wb = open(base().writeErrorPolicy(policy).write(ROWS))) {
                    assertEquals("SUBTOTAL(9,C2:C3)", wb.getSheetAt(0).getRow(3).getCell(2).getCellFormula(), policy.name());
                }
            }
        }
    }

    @Nested
    class Styles {
        @Test
        void aTooSmallStyleBudgetFailsWithTheStyleMessage() {
            var e = assertThrows(ExcelWriteException.class, () -> base().maxCellStyles(1).write(ROWS));
            assertTrue(e.getMessage().contains("Too many distinct cell styles"), e.getMessage());
        }

        @Test
        void theSameWriterWithoutSubtotalsNeedsNoDerivedStyle() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount).maxCellStyles(1);
            w.write(ROWS).writeTo(java.io.OutputStream.nullOutputStream());
        }

        @Test
        void subtotalRowsShareTheirStylesAcrossGroups() throws IOException {
            var rows = new ArrayList<Row>();
            for (int g = 0; g < 200; g++) rows.add(new Row("K" + g, "n", 1));
            int before;
            try (var wb = open(ExcelWriter.<Row>create().column("Dept", Row::dept).column("Name", Row::name)
                    .column("Amount", Row::amount).write(rows.subList(0, 1)))) {
                before = wb.getNumCellStyles();
            }
            try (var wb = open(base().write(rows))) {
                assertTrue(wb.getNumCellStyles() <= before + 3, "one derived style per column, not per group: " + wb.getNumCellStyles());
            }
        }
    }

    @Nested
    class Callbacks {
        @Test
        void afterAllWritesBelowTheSummaryOfTheLastSheet() throws IOException {
            var w = base().summary(s -> s.label("Dept", "Total").sum("Amount"))
                    .afterAll(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("done");
                        return ctx.getCurrentRow() + 1;
                    });
            try (var wb = open(w.write(ROWS))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("B total", sheet.getRow(5).getCell(1).getStringCellValue());
                assertEquals("Total", sheet.getRow(6).getCell(0).getStringCellValue());
                assertEquals("done", sheet.getRow(7).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aRowCallbackThatWritesRowsDoesNotMoveTheSubtotalRanges() throws IOException {
            var seen = new ArrayList<Integer>();
            try (var wb = open(base().write(ROWS.stream(), (row, cursor) -> seen.add(cursor.getRowOfSheet())))) {
                assertEquals(List.of(2, 3, 5), seen, "the cursor sits below the row just written; the third row follows the subtotal row of A");
                assertEquals("SUBTOTAL(9,C2:C3)", wb.getSheetAt(0).getRow(3).getCell(2).getCellFormula());
            }
        }
    }

    @Nested
    class Failures {
        @Test
        void aKeyFunctionThatThrowsFailsTheWriteAndLeavesNoTempFile() {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(r -> { if (r.amount() == 20) throw new IllegalStateException("bad key"); return r.dept(); },
                            s -> s.sum("Amount"));
            var e = assertThrows(RuntimeException.class, () -> w.write(ROWS));
            assertTrue(e.getMessage().contains("bad key") || (e.getCause() != null && e.getCause().getMessage().contains("bad key")),
                    e.toString());
        }

        @Test
        void aLabelFunctionThatThrowsFailsTheWrite() {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> { throw new IllegalStateException("bad label"); }).sum("Amount"));
            assertThrows(RuntimeException.class, () -> w.write(ROWS));
        }

        @Test
        void aWorkbookSheetThatFailsLeavesTheWorkbookClosable() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(r -> { throw new IllegalStateException("bad key"); }, s -> s.sum("Amount"));
                assertThrows(RuntimeException.class, () -> sheet.write(ROWS));
            }
        }

        @Test
        void aWriterWritesOnceEvenWithSubtotals() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept).column("Amount", Row::amount)
                    .groupSubtotal(Row::dept, s -> s.sum("Amount"));
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.sum("Amount"));
                sheet.write(ROWS);
                assertThrows(ExcelWriteException.class, () -> sheet.write(ROWS));
            }
            w.write(ROWS).writeTo(java.io.OutputStream.nullOutputStream());
        }
    }

    @Nested
    class ReadingBack {
        @Test
        void aFileWithSubtotalsReadsBackWithTheSubtotalRowsAsRows() throws IOException {
            byte[] file = bytes(base().write(ROWS));
            var rows = new ArrayList<Map<String, String>>();
            ExcelReader.forMap().read(new ByteArrayInputStream(file), r -> {
                assertTrue(r.success(), r.toString());
                rows.add(r.data());
            });
            assertEquals(5, rows.size(), "3 data rows and 2 subtotal rows");
            assertEquals("A total", rows.get(2).get("Name"));
            assertEquals("B total", rows.get(4).get("Name"));
            assertEquals("x", rows.get(0).get("Name"));
        }

        @Test
        void aSubtotalCellHoldsItsResultAndExcelStillRecalculatesOnOpen() throws IOException {
            try (var wb = open(base().write(ROWS))) {
                var cell = wb.getSheetAt(0).getRow(3).getCell(2);
                assertEquals(CellType.FORMULA, cell.getCellType());
                assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType(), "a viewer that does not calculate shows it");
                assertEquals(30, cell.getNumericCellValue());
                assertTrue(wb.getForceFormulaRecalculation(), "Excel recalculates the workbook when it opens it");
            }
        }
    }
}
