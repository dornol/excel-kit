package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class PrintSetupTest {

    @Test
    void printSetup_landscapeOrientation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.orientation(ExcelPrintSetup.Orientation.LANDSCAPE))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getPrintSetup().getLandscape(),
                    "Landscape orientation should set landscape to true");
        }
    }

    @Test
    void printSetup_portraitOrientation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.orientation(ExcelPrintSetup.Orientation.PORTRAIT))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getPrintSetup().getLandscape(),
                    "Portrait orientation should set landscape to false");
        }
    }

    @Test
    void printSetup_paperSizeA4() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.paperSize(ExcelPrintSetup.PaperSize.A4))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(PrintSetup.A4_PAPERSIZE,
                    wb.getSheetAt(0).getPrintSetup().getPaperSize(),
                    "Paper size should be A4");
        }
    }

    @Test
    void printSetup_allMargins() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.margins(0.5, 0.6, 0.7, 0.8))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(0.5, sheet.getMargin(Sheet.LeftMargin), 0.001,
                    "Left margin should be 0.5");
            assertEquals(0.6, sheet.getMargin(Sheet.RightMargin), 0.001,
                    "Right margin should be 0.6");
            assertEquals(0.7, sheet.getMargin(Sheet.TopMargin), 0.001,
                    "Top margin should be 0.7");
            assertEquals(0.8, sheet.getMargin(Sheet.BottomMargin), 0.001,
                    "Bottom margin should be 0.8");
        }
    }

    @Test
    void printSetup_individualMargins() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps
                        .leftMargin(1.0)
                        .rightMargin(1.5)
                        .topMargin(2.0)
                        .bottomMargin(2.5))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1.0, sheet.getMargin(Sheet.LeftMargin), 0.001,
                    "Individual left margin");
            assertEquals(1.5, sheet.getMargin(Sheet.RightMargin), 0.001,
                    "Individual right margin");
            assertEquals(2.0, sheet.getMargin(Sheet.TopMargin), 0.001,
                    "Individual top margin");
            assertEquals(2.5, sheet.getMargin(Sheet.BottomMargin), 0.001,
                    "Individual bottom margin");
        }
    }

    @Test
    void printSetup_headerCenterAndFooterCenter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps
                        .headerCenter("My Report")
                        .footerCenter("Page &P of &N"))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("My Report", sheet.getHeader().getCenter(),
                    "Header center should contain the configured text");
            assertEquals("Page &P of &N", sheet.getFooter().getCenter(),
                    "Footer center should contain the configured text");
        }
    }

    @Test
    void printSetup_allHeaderFooterSections() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps
                        .headerLeft("HL")
                        .headerCenter("HC")
                        .headerRight("HR")
                        .footerLeft("FL")
                        .footerCenter("FC")
                        .footerRight("FR"))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("HL", sheet.getHeader().getLeft(), "Header left");
            assertEquals("HC", sheet.getHeader().getCenter(), "Header center");
            assertEquals("HR", sheet.getHeader().getRight(), "Header right");
            assertEquals("FL", sheet.getFooter().getLeft(), "Footer left");
            assertEquals("FC", sheet.getFooter().getCenter(), "Footer center");
            assertEquals("FR", sheet.getFooter().getRight(), "Footer right");
        }
    }

    @Test
    void printSetup_repeatHeaderRows() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.repeatHeaderRows())
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var repeatingRows = wb.getSheetAt(0).getRepeatingRows();
            assertNotNull(repeatingRows,
                    "Repeating rows should be set when repeatHeaderRows() is called");
            assertEquals(0, repeatingRows.getFirstRow(),
                    "Repeating rows should start from row 0");
        }
    }

    @Test
    void printSetup_repeatCustomRows() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.repeatRows(0, 1))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var repeatingRows = wb.getSheetAt(0).getRepeatingRows();
            assertNotNull(repeatingRows,
                    "Repeating rows should be set when repeatRows() is called");
            assertEquals(0, repeatingRows.getFirstRow(), "Custom repeat row start");
            assertEquals(1, repeatingRows.getLastRow(), "Custom repeat row end");
        }
    }

    @Test
    void printSetup_fitToPageWidth() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.fitToPageWidth())
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.getFitToPage(),
                    "Sheet should be set to fit-to-page");
            assertEquals(1, sheet.getPrintSetup().getFitWidth(),
                    "Fit width should be 1 for fitToPageWidth()");
            assertEquals(0, sheet.getPrintSetup().getFitHeight(),
                    "Fit height should be 0 (auto) for fitToPageWidth()");
        }
    }

    @Test
    void printSetup_fitToPageCustom() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s)
                .printSetup(ps -> ps.fitToPage(2, 3))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.getFitToPage(),
                    "Sheet should be set to fit-to-page");
            assertEquals(2, sheet.getPrintSetup().getFitWidth(),
                    "Custom fit width should be 2");
            assertEquals(3, sheet.getPrintSetup().getFitHeight(),
                    "Custom fit height should be 3");
        }
    }

    @Test
    void printSetup_withExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("PrintSheet")
                    .column("Col", s -> s)
                    .printSetup(ps -> ps
                            .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                            .paperSize(ExcelPrintSetup.PaperSize.A4)
                            .margins(0.5, 0.5, 0.75, 0.75)
                            .headerCenter("Report Title")
                            .footerCenter("Page &P"))
                    .write(Stream.of("data"));
            workbook.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.getPrintSetup().getLandscape(),
                    "ExcelSheetWriter should apply landscape orientation");
            assertEquals(PrintSetup.A4_PAPERSIZE, sheet.getPrintSetup().getPaperSize(),
                    "ExcelSheetWriter should apply A4 paper size");
            assertEquals(0.5, sheet.getMargin(Sheet.LeftMargin), 0.001,
                    "ExcelSheetWriter should apply left margin");
            assertEquals("Report Title", sheet.getHeader().getCenter(),
                    "ExcelSheetWriter should apply header center");
            assertEquals("Page &P", sheet.getFooter().getCenter(),
                    "ExcelSheetWriter should apply footer center");
        }
    }

    @Test
    void printSetup_withSheetRollover_allSheetsHavePrintSetup() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Data")
                    .column("Col", s -> s)
                    .maxRows(3)
                    .printSetup(ps -> ps
                            .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                            .paperSize(ExcelPrintSetup.PaperSize.LETTER)
                            .headerCenter("Multi-Sheet Report"))
                    .write(Stream.of("A", "B", "C", "D", "E", "F", "G"));
            workbook.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getNumberOfSheets() >= 2,
                    "Should have multiple sheets from rollover");
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                var sheet = wb.getSheetAt(i);
                assertTrue(sheet.getPrintSetup().getLandscape(),
                        "All rollover sheets should have landscape orientation, sheet " + i);
                assertEquals(PrintSetup.LETTER_PAPERSIZE, sheet.getPrintSetup().getPaperSize(),
                        "All rollover sheets should have LETTER paper size, sheet " + i);
                assertEquals("Multi-Sheet Report", sheet.getHeader().getCenter(),
                        "All rollover sheets should have header center, sheet " + i);
            }
        }
    }

    @Test
    void printSetup_combinedWithAutoFilterAndFreezePane() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s.toUpperCase())
                .autoFilter(true)
                .freezeRows(1)
                .printSetup(ps -> ps
                        .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                        .fitToPageWidth()
                        .margins(0.25, 0.25, 0.5, 0.5)
                        .headerCenter("Combined Features Report")
                        .footerRight("&D"))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Print setup
            assertTrue(sheet.getPrintSetup().getLandscape(),
                    "Landscape should work with autoFilter and freezePane");
            assertTrue(sheet.getFitToPage(),
                    "Fit-to-page should work with autoFilter and freezePane");
            assertEquals(0.25, sheet.getMargin(Sheet.LeftMargin), 0.001,
                    "Margins should work with autoFilter and freezePane");
            assertEquals("Combined Features Report", sheet.getHeader().getCenter(),
                    "Header should work with autoFilter and freezePane");
            assertEquals("&D", sheet.getFooter().getRight(),
                    "Footer right should work with autoFilter and freezePane");
            // Auto filter should also be present
            assertNotNull(sheet.getCTWorksheet().getAutoFilter(),
                    "Auto filter should still be set alongside print setup");
        }
    }
}
