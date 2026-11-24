package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** A sheet without rows says so: a line below the header that is no data row. */
class EmptyMessageTest {
    record Row(String name, int amount, int qty) {}

    private static ExcelWriter<Row> writer() {
        return ExcelWriter.<Row>create()
                .column("Name", Row::name).column("Amount", Row::amount).column("Qty", Row::qty);
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static final List<Row> NONE = List.of();

    @Nested
    class TheLine {
        @Test
        void anEmptySheetSaysSoBelowTheHeaderOverAllTheColumns() throws IOException {
            try (var wb = open(writer().emptyMessage("No data").write(NONE))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("No data", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(1, sheet.getNumMergedRegions());
                assertEquals("A2:C2", sheet.getMergedRegion(0).formatAsString());
                assertEquals(1, sheet.getLastRowNum(), "the header and the line, nothing else");
            }
        }

        @Test
        void theLineIsCenteredGrayAndItalic() throws IOException {
            try (var wb = open(writer().emptyMessage("No data").write(NONE))) {
                var style = (XSSFCellStyle) wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
                var font = wb.getFontAt(style.getFontIndex());
                assertTrue(font.getItalic());
                var color = ((org.apache.poi.xssf.usermodel.XSSFFont) font).getXSSFColor().getRGB();
                assertArrayEquals(new byte[]{(byte) 128, (byte) 128, (byte) 128}, color);
            }
        }

        @Test
        void aSingleColumnSheetHasNothingToMerge() throws IOException {
            var w = ExcelWriter.<Row>create().column("Name", Row::name).emptyMessage("Nothing here");
            try (var wb = open(w.write(NONE))) {
                assertEquals(0, wb.getSheetAt(0).getNumMergedRegions());
                assertEquals("Nothing here", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aSheetWithRowsIsNotTouched() throws IOException {
            try (var wb = open(writer().emptyMessage("No data").write(List.of(new Row("a", 1, 2))))) {
                var sheet = wb.getSheetAt(0);
                assertEquals(1, sheet.getLastRowNum());
                assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(0, sheet.getNumMergedRegions());
            }
        }

        @Test
        void theRowHeightOfTheSheetApplies() throws IOException {
            try (var wb = open(writer().rowHeight(33f).emptyMessage("No data").write(NONE))) {
                assertEquals(33f, wb.getSheetAt(0).getRow(1).getHeightInPoints());
            }
        }

        @Test
        void aTextWithEscapesAndNonLatinCharactersIsKeptAsItIs() throws IOException {
            String text = "_x0041_ 데이터가 없습니다 & <none>";
            try (var wb = open(writer().emptyMessage(text).write(NONE))) {
                assertEquals(text, wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aStreamThatYieldsNothingCountsAsEmptyWhateverItIs() throws IOException {
            try (var wb = open(writer().emptyMessage("No data").write(Stream.<Row>empty()))) {
                assertEquals("No data", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
            try (var wb = open(writer().emptyMessage("No data").write((Iterable<Row>) new ArrayList<Row>()))) {
                assertEquals("No data", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aPreambleLeavesTheLineDirectlyBelowTheHeader() throws IOException {
            var w = writer().emptyMessage("No data").beforeHeader(ctx -> {
                ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Title");
                return ctx.getCurrentRow() + 2;
            });
            try (var wb = open(w.write(NONE))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("No data", sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals("A4:C4", sheet.getMergedRegion(0).formatAsString());
            }
        }

        @Test
        void aGroupHeaderLeavesTheLineBelowTheColumnNames() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Name", Row::name, c -> c.group("G"))
                    .column("Amount", Row::amount, c -> c.group("G"))
                    .emptyMessage("No data");
            try (var wb = open(w.write(NONE))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("No data", sheet.getRow(2).getCell(0).getStringCellValue());
            }
        }
    }

    @Nested
    class NotADataRow {
        @Test
        void theFilterCoversTheHeaderOnly() throws IOException {
            try (var wb = open(writer().autoFilter(true).emptyMessage("No data").write(NONE))) {
                assertEquals("A1:C1", wb.getSheetAt(0).getCTWorksheet().getAutoFilter().getRef());
            }
        }

        @Test
        void aSummaryBelowTheLineCountsNothingAndKeepsItsLabel() throws IOException {
            var w = writer().emptyMessage("No data").summary(s -> s.label("Name", "Total").sum("Amount").count("Qty"));
            try (var wb = open(w.write(NONE))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("No data", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("Total - Sum", sheet.getRow(2).getCell(0).getStringCellValue());
                for (int r = 2; r <= 3; r++) {
                    for (int c = 1; c <= 2; c++) {
                        var cell = sheet.getRow(r).getCell(c);
                        assertTrue(cell == null || cell.getCellType() != CellType.FORMULA, "no formula over a range that holds only the line: row " + r);
                    }
                }
            }
        }

        @Test
        void afterDataAndAfterAllComeBelowTheLine() throws IOException {
            var w = writer().emptyMessage("No data")
                    .afterData(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("after");
                        return ctx.getCurrentRow() + 1;
                    });
            try (var wb = open(w.write(NONE))) {
                assertEquals("after", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
            }
        }

        @Test
        void aChartAndANamedRangeHaveNothingToCoverAndDoNotFail() throws IOException {
            var w = writer().emptyMessage("No data").namedRange("Amounts", 1)
                    .chart(c -> c.categoryColumn(0).valueColumn(1, "Amount"));
            try (var wb = open(w.write(NONE))) {
                assertNull(wb.getSheetAt(0).getDrawingPatriarch(), "no chart over a line of text");
                assertEquals(0, wb.getAllNames().size(), "no name over a line of text");
            }
        }

        @Test
        void conditionalFormattingAndFreezeDoNotFail() throws IOException {
            var w = writer().emptyMessage("No data").freezeRows(1)
                    .conditionalFormatting(c -> c.columns(1).greaterThan("5", io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW));
            try (var wb = open(w.write(NONE))) {
                assertEquals("No data", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void subtotalsAndMergedColumnsHaveNothingToDoWithNoRows() throws IOException {
            var w = ExcelWriter.<Row>create()
                    .column("Name", Row::name, c -> c.mergeSameValues())
                    .column("Amount", Row::amount)
                    .groupSubtotal(Row::name, s -> s.label("Name", k -> "t").sum("Amount"))
                    .emptyMessage("No data");
            try (var wb = open(w.write(NONE))) {
                assertEquals(1, wb.getSheetAt(0).getNumMergedRegions());
                assertEquals(1, wb.getSheetAt(0).getLastRowNum());
            }
        }
    }

    @Nested
    class Refusals {
        @Test
        void aTableIsRefusedInBothWriters() throws IOException {
            assertThrows(ExcelWriteException.class, () -> writer().table("Data").emptyMessage("No data").write(NONE));
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S").column("Name", Row::name).table("Data").emptyMessage("No data");
                assertThrows(ExcelWriteException.class, () -> sheet.write(NONE));
            }
        }

        @Test
        void aSheetWithoutTheMessageStillTakesATable() throws IOException {
            try (var wb = open(writer().table("Data").write(List.of(new Row("a", 1, 2))))) {
                assertEquals(1, wb.getSheetAt(0).getTables().size());
            }
        }

        @Test
        void theTextMustNotBeNull() {
            assertThrows(NullPointerException.class, () -> ExcelWriter.<Row>create().emptyMessage(null));
        }

        @Test
        void aStyleBudgetTooSmallForTheLineFailsWithTheStyleMessage() {
            var e = assertThrows(ExcelWriteException.class, () -> writer().maxCellStyles(1).emptyMessage("No data").write(NONE));
            assertTrue(e.getMessage().contains("Too many distinct cell styles"), e.getMessage());
        }
    }

    @Nested
    class WorkbookAndReading {
        @Test
        void everySheetOfAWorkbookSaysItOnItsOwn() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("Empty").column("Name", Row::name).column("Amount", Row::amount).emptyMessage("Nothing for Empty").write(NONE);
                workbook.<Row>sheet("Full").column("Name", Row::name).column("Amount", Row::amount).emptyMessage("Nothing for Full")
                        .write(List.of(new Row("a", 1, 2)));
                workbook.<Row>sheet("Plain").column("Name", Row::name).column("Amount", Row::amount).write(NONE);
                try (var wb = open(workbook.finish())) {
                    assertEquals("Nothing for Empty", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                    assertEquals("a", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
                    assertEquals(1, wb.getSheetAt(1).getLastRowNum(), "a sheet with a row gets no line below it");
                    assertEquals(0, wb.getSheetAt(1).getNumMergedRegions());
                    assertEquals(0, wb.getSheetAt(2).getLastRowNum(), "no message asked: the header alone, as before");
                }
            }
        }

        @Test
        void aWorkbookSheetsSummaryBelowTheLineCountsNothing() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("S").column("Name", Row::name).column("Amount", Row::amount)
                        .emptyMessage("No data").summary(s -> s.label("Name", "Total").sum("Amount")).write(NONE);
                try (var wb = open(workbook.finish())) {
                    var sheet = wb.getSheetAt(0);
                    assertEquals("No data", sheet.getRow(1).getCell(0).getStringCellValue());
                    assertEquals("Total", sheet.getRow(2).getCell(0).getStringCellValue());
                    var cell = sheet.getRow(2).getCell(1);
                    assertTrue(cell == null || cell.getCellType() != CellType.FORMULA, "no formula over a range that holds only the line");
                }
            }
        }

        @Test
        void theSheetsOfAWorkbookShareOneStyleForTheLine() throws IOException {
            int one = cellFormats(workbookOfEmptySheets(1));
            int fifty = cellFormats(workbookOfEmptySheets(50));
            assertEquals(one, fifty, "fifty sheets with the line need no more cell formats than one");
        }

        @Test
        void sharedStyleStillGivesEverySheetItsOwnHeightColumnsAndText() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("Two").column("Name", Row::name).column("Amount", Row::amount)
                        .rowHeight(30).emptyMessage("two columns").write(NONE);
                workbook.<Row>sheet("Three").column("Name", Row::name).column("Amount", Row::amount).column("Qty", Row::qty)
                        .rowHeight(18).emptyMessage("three columns").write(NONE);
                workbook.<Row>sheet("One").column("Name", Row::name).emptyMessage("one column").write(NONE);
                try (var wb = open(workbook.finish())) {
                    assertEquals("A2:B2", wb.getSheetAt(0).getMergedRegion(0).formatAsString());
                    assertEquals(30f, wb.getSheetAt(0).getRow(1).getHeightInPoints(), 0.01);
                    assertEquals("A2:C2", wb.getSheetAt(1).getMergedRegion(0).formatAsString());
                    assertEquals(18f, wb.getSheetAt(1).getRow(1).getHeightInPoints(), 0.01);
                    assertEquals(0, wb.getSheetAt(2).getNumMergedRegions());
                    assertEquals("one column", wb.getSheetAt(2).getRow(1).getCell(0).getStringCellValue());
                    var a = (XSSFCellStyle) wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                    var c = (XSSFCellStyle) wb.getSheetAt(1).getRow(1).getCell(0).getCellStyle();
                    assertEquals(a.getIndex(), c.getIndex(), "one shared style");
                    assertEquals(HorizontalAlignment.CENTER, c.getAlignment());
                    assertTrue(wb.getFontAt(c.getFontIndex()).getItalic());
                }
            }
        }

        @Test
        void aFileWithTheLineReadsBackWithItAsARow() throws IOException {
            var out = new ByteArrayOutputStream();
            writer().emptyMessage("No data").write(NONE).writeTo(out);
            var rows = new ArrayList<String>();
            ExcelReader.forMap().read(new ByteArrayInputStream(out.toByteArray()), r -> rows.add(r.data().get("Name")));
            assertEquals(List.of("No data"), rows, "documented: leave the line off a sheet people fill in and upload");
        }

        @Test
        void withoutTheMessageAnEmptySheetIsTheHeaderAlone() throws IOException {
            try (var wb = open(writer().write(NONE))) {
                assertEquals(0, wb.getSheetAt(0).getLastRowNum());
            }
        }
    }

    private static byte[] workbookOfEmptySheets(int sheets) throws IOException {
        try (var workbook = ExcelWorkbook.create()) {
            for (int i = 0; i < sheets; i++) {
                workbook.<Row>sheet("S" + i).column("Name", Row::name).column("Amount", Row::amount)
                        .emptyMessage("No data").write(NONE);
            }
            var out = new ByteArrayOutputStream();
            workbook.finish().writeTo(out);
            return out.toByteArray();
        }
    }

    private static int cellFormats(byte[] file) throws IOException {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            return wb.getNumCellStyles();
        }
    }
}
