package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.ExcelPrintSetup.Orientation;
import io.github.dornol.excelkit.excel.ExcelPrintSetup.PaperSize;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The page setup a sheet is written with, read back from the sheet XML, and the setups that are refused. */
class PrintSetupDetailsTest {

    record Row(String name) {
    }

    private static XSSFSheet sheet(Consumer<ExcelPrintSetup> setup) throws Exception {
        return sheet(setup, w -> { }, 2).getSheetAt(0);
    }

    private static XSSFWorkbook sheet(Consumer<ExcelPrintSetup> setup, Consumer<ExcelWriter<Row>> more, int rows) throws Exception {
        List<Row> data = new ArrayList<>();
        for (int i = 0; i < rows; i++) data.add(new Row("n" + i));
        ExcelWriter<Row> writer = ExcelWriter.<Row>create().column("name", Row::name).printSetup(setup);
        more.accept(writer);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(data).writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static String repeated(XSSFWorkbook workbook, int sheetIndex) {
        for (var name : workbook.getAllNames()) {
            if (name.getNameName().equals("_xlnm.Print_Titles") && name.getSheetIndex() == sheetIndex) {
                return name.getRefersToFormula();
            }
        }
        return null;
    }

    // ---- page

    @Test
    void orientationAndPaperAreWrittenAsExcelCodes() throws Exception {
        var landscape = sheet(p -> p.orientation(Orientation.LANDSCAPE).paperSize(PaperSize.A3)).getCTWorksheet().getPageSetup();
        assertEquals("landscape", landscape.getOrientation().toString());
        assertEquals(8, landscape.getPaperSize());

        var portrait = sheet(p -> p.orientation(Orientation.PORTRAIT).paperSize(PaperSize.A4)).getCTWorksheet().getPageSetup();
        assertEquals("portrait", portrait.getOrientation().toString());
        assertEquals(9, portrait.getPaperSize());

        int[] codes = {1, 5, 8, 9, 11, 12, 13};
        PaperSize[] papers = {PaperSize.LETTER, PaperSize.LEGAL, PaperSize.A3, PaperSize.A4, PaperSize.A5, PaperSize.B4, PaperSize.B5};
        for (int i = 0; i < papers.length; i++) {
            PaperSize paper = papers[i];
            assertEquals(codes[i], sheet(p -> p.paperSize(paper)).getCTWorksheet().getPageSetup().getPaperSize(), paper.toString());
        }
    }

    @Test
    void whatIsNotSetKeepsExcelsDefault() throws Exception {
        var setup = sheet(p -> p.leftMargin(1)).getCTWorksheet();
        assertFalse(setup.getPageSetup().isSetOrientation() && "landscape".equals(setup.getPageSetup().getOrientation().toString()));
        assertFalse(setup.getPageSetup().isSetPaperSize() && setup.getPageSetup().getPaperSize() == 8);
    }

    @Test
    void marginsAreWrittenInInches() throws Exception {
        var margins = sheet(p -> p.margins(0.25, 0.5, 1.0, 1.5)).getCTWorksheet().getPageMargins();
        assertEquals(0.25, margins.getLeft());
        assertEquals(0.5, margins.getRight());
        assertEquals(1.0, margins.getTop());
        assertEquals(1.5, margins.getBottom());

        var one = sheet(p -> p.topMargin(2.0)).getCTWorksheet().getPageMargins();
        assertEquals(2.0, one.getTop());
        assertEquals(0.7, one.getLeft(), "an unset margin stays at Excel's default");
        assertEquals(0.75, one.getBottom());
    }

    @Test
    void aMarginIsANonNegativeNumberOfInches() {
        for (double bad : new double[]{-0.01, -1, Double.NaN, Double.POSITIVE_INFINITY, Double.NEGATIVE_INFINITY}) {
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().leftMargin(bad), "left " + bad);
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().rightMargin(bad));
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().topMargin(bad));
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().bottomMargin(bad));
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().margins(1, 1, 1, bad));
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().margins(bad, 1, 1, 1));
        }
        assertDoesNotThrow(() -> new ExcelPrintSetup().margins(0, 0, 0, 0));
    }

    // ---- fit to page

    @Test
    void fitToPageWritesThePagesWideAndTall() throws Exception {
        var one = sheet(p -> p.fitToPageWidth());
        assertTrue(one.getCTWorksheet().getSheetPr().getPageSetUpPr().getFitToPage());
        assertEquals(1, one.getCTWorksheet().getPageSetup().getFitToWidth());
        assertEquals(0, one.getCTWorksheet().getPageSetup().getFitToHeight());

        var some = sheet(p -> p.fitToPage(3, 2)).getCTWorksheet().getPageSetup();
        assertEquals(3, some.getFitToWidth());
        assertEquals(2, some.getFitToHeight());

        var most = sheet(p -> p.fitToPage(32_767, 32_767)).getCTWorksheet().getPageSetup();
        assertEquals(32_767, most.getFitToWidth());
        assertEquals(32_767, most.getFitToHeight());
    }

    @Test
    void thePagesAreBetweenZeroAndThe16BitLimit() {
        assertDoesNotThrow(() -> new ExcelPrintSetup().fitToPage(0, 0));
        for (int[] bad : new int[][]{{-1, 1}, {1, -1}, {32_768, 1}, {1, 32_768}, {40_000, 1}, {65_536, 1}}) {
            assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().fitToPage(bad[0], bad[1]), bad[0] + "," + bad[1]);
        }
    }

    // ---- header and footer

    @Test
    void theSectionsOfAHeaderAndFooterAreWrittenWithTheirCodes() throws Exception {
        var all = sheet(p -> p.headerLeft("L1").headerCenter("C1").headerRight("R1")
                .footerLeft("L2").footerCenter("Page &P of &N").footerRight("&D &T &F")).getCTWorksheet().getHeaderFooter();
        assertEquals(java.util.Map.of('L', "L1", 'C', "C1", 'R', "R1"), sections(all.getOddHeader()));
        assertEquals(java.util.Map.of('L', "L2", 'C', "Page &P of &N", 'R', "&D &T &F"), sections(all.getOddFooter()));

        var onlyRight = sheet(p -> p.headerRight("only")).getCTWorksheet().getHeaderFooter();
        assertEquals(java.util.Map.of('R', "only"), sections(onlyRight.getOddHeader()));
        assertNull(onlyRight.getOddFooter());
    }

    @Test
    void aLiteralAmpersandIsWrittenTwice() throws Exception {
        // "R&&D" is what a caller writes to print R&D; "R&D" would print R and the date
        var header = sheet(p -> p.headerLeft("R&&D Report")).getCTWorksheet().getHeaderFooter();
        assertEquals(java.util.Map.of('L', "R&&D Report"), sections(header.getOddHeader()));
    }

    @Test
    void aHeaderOrFooterHoldsAtMost255CharactersIncludingItsSectionMarks() {
        // &L + 253 characters = 255
        assertDoesNotThrow(() -> validated(new ExcelPrintSetup().headerLeft("h".repeat(253))));
        assertThrows(IllegalArgumentException.class, () -> validated(new ExcelPrintSetup().headerLeft("h".repeat(254))));
        // three sections: 2 + 2 + 2 marks and 249 characters
        assertDoesNotThrow(() -> validated(new ExcelPrintSetup().footerLeft("a".repeat(100)).footerCenter("b".repeat(100)).footerRight("c".repeat(49))));
        assertThrows(IllegalArgumentException.class, () -> validated(new ExcelPrintSetup().footerLeft("a".repeat(100)).footerCenter("b".repeat(100)).footerRight("c".repeat(50))));
        // the header and the footer are counted apart
        assertDoesNotThrow(() -> validated(new ExcelPrintSetup().headerLeft("h".repeat(253)).footerLeft("f".repeat(253))));
    }

    @Test
    void aTooLongHeaderIsRefusedWhenTheSetupIsGivenToTheWriter() {
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
                () -> ExcelWriter.<Row>create().printSetup(p -> p.headerCenter("x".repeat(300))));
        assertTrue(error.getMessage().contains("header") && error.getMessage().contains("255"), error.getMessage());
    }

    /**
     * The text of each section of a header or footer string by its mark ({@code &L}, {@code &C}, {@code &R}); POI
     * writes the sections in an order of its own, which Excel does not care about. A doubled ampersand is kept as is.
     */
    private static java.util.Map<Character, String> sections(String text) {
        java.util.Map<Character, String> result = new java.util.TreeMap<>();
        char current = 0;
        StringBuilder part = new StringBuilder();
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (c == '&' && i + 1 < text.length()) {
                char next = text.charAt(i + 1);
                if (next == 'L' || next == 'C' || next == 'R') {
                    if (current != 0) result.put(current, part.toString());
                    current = next;
                    part.setLength(0);
                    i++;
                    continue;
                }
                part.append(c).append(next);
                i++;
                continue;
            }
            part.append(c);
        }
        if (current != 0) result.put(current, part.toString());
        return result;
    }

    private static void validated(ExcelPrintSetup setup) {
        setup.validate();
    }

    // ---- repeated rows

    @Test
    void repeatedRowsAreWrittenAsTheSheetsPrintTitles() throws Exception {
        try (var workbook = sheet(p -> p.repeatRows(0, 0), w -> { }, 2)) {
            assertEquals("Sheet0!$1:$1", repeated(workbook, 0));
        }
        try (var workbook = sheet(p -> p.repeatRows(1, 3), w -> { }, 2)) {
            assertEquals("Sheet0!$2:$4", repeated(workbook, 0));
        }
    }

    @Test
    void theHeaderRowsRepeatDownToTheColumnNames() throws Exception {
        try (var workbook = sheet(p -> p.repeatHeaderRows(), w -> { }, 2)) {
            assertEquals("Sheet0!$1:$1", repeated(workbook, 0));
        }
        // a preamble and a group header are part of what repeats
        try (var workbook = sheet(p -> p.repeatHeaderRows(),
                w -> w.beforeHeader(ctx -> 3).column("g", Row::name, c -> c.group("Group")), 2)) {
            assertEquals("Sheet0!$1:$5", repeated(workbook, 0), "3 rows of preamble, the group row and the column names");
        }
    }

    @Test
    void theLastWordOnRepeatedRowsWins() throws Exception {
        try (var workbook = sheet(p -> p.repeatHeaderRows().repeatRows(2, 2), w -> { }, 2)) {
            assertEquals("Sheet0!$3:$3", repeated(workbook, 0));
        }
        try (var workbook = sheet(p -> p.repeatRows(2, 2).repeatHeaderRows(), w -> { }, 2)) {
            assertEquals("Sheet0!$1:$1", repeated(workbook, 0));
        }
    }

    @Test
    void everySheetOfARolloverGetsTheSetup() throws Exception {
        try (var workbook = sheet(p -> p.repeatHeaderRows().orientation(Orientation.LANDSCAPE).headerCenter("Report"),
                w -> w.maxRows(3), 7)) {
            assertEquals(3, workbook.getNumberOfSheets());
            for (int i = 0; i < 3; i++) {
                assertEquals("landscape", workbook.getSheetAt(i).getCTWorksheet().getPageSetup().getOrientation().toString(), "sheet " + i);
                assertEquals(java.util.Map.of('C', "Report"),
                        sections(workbook.getSheetAt(i).getCTWorksheet().getHeaderFooter().getOddHeader()), "sheet " + i);
                assertTrue(repeated(workbook, i) != null && repeated(workbook, i).endsWith("$1:$1"), "sheet " + i);
            }
        }
    }

    @Test
    void repeatedRowsMustBeAValidRange() {
        assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().repeatRows(-1, 2));
        assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().repeatRows(3, 1));
        assertThrows(IllegalArgumentException.class, () -> new ExcelPrintSetup().repeatRows(0, 1_048_576));
        assertDoesNotThrow(() -> new ExcelPrintSetup().repeatRows(0, 1_048_575));
        assertDoesNotThrow(() -> new ExcelPrintSetup().repeatRows(2, 2));
    }

    // ---- valid XML

    @Test
    void everySettingWritesValidSpreadsheetML() throws Exception {
        XSSFSheet sheet = sheet(p -> p.orientation(Orientation.LANDSCAPE).paperSize(PaperSize.B5)
                .margins(0.2, 0.3, 0.4, 0.5).headerLeft("a").headerCenter("b").headerRight("c")
                .footerLeft("d").footerCenter("e").footerRight("f").fitToPage(2, 3).repeatHeaderRows());
        List<org.apache.xmlbeans.XmlError> errors = new ArrayList<>();
        assertTrue(sheet.getCTWorksheet().validate(new org.apache.xmlbeans.XmlOptions().setErrorListener(errors)), errors.toString());
    }
}
