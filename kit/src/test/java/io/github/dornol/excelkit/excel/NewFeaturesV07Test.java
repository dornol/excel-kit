package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive tests for v0.7 features:
 * 1. Column Hidden
 * 2. Rich Text (ExcelRichText + ExcelDataType.RICH_TEXT)
 * 3. Print Setup (ExcelPrintSetup)
 */
class NewFeaturesV07Test {

    // ============================================================
    // Feature 1: Column Hidden
    // ============================================================
    @Nested
    class ColumnHiddenTests {

        @Test
        void hidden_noArgs_columnIsHidden() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Visible", s -> s)
                    .addColumn("Secret", s -> "hidden-value", c -> c.hidden())
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertFalse(sheet.isColumnHidden(0), "First column should not be hidden");
                assertTrue(sheet.isColumnHidden(1), "Second column should be hidden when hidden() is called");
            }
        }

        @Test
        void hidden_true_columnIsHidden() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s, c -> c.hidden(true))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getSheetAt(0).isColumnHidden(0),
                        "Column should be hidden when hidden(true) is called");
            }
        }

        @Test
        void hidden_false_columnIsNotHidden() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s, c -> c.hidden(false))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertFalse(wb.getSheetAt(0).isColumnHidden(0),
                        "Column should not be hidden when hidden(false) is called");
            }
        }

        @Test
        void hidden_mixedColumns_someHiddenSomeNot() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s)
                    .addColumn("B", s -> s, c -> c.hidden(true))
                    .addColumn("C", s -> s)
                    .addColumn("D", s -> s, c -> c.hidden())
                    .addColumn("E", s -> s, c -> c.hidden(false))
                    .write(Stream.of("row1"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertFalse(sheet.isColumnHidden(0), "Column A should be visible");
                assertTrue(sheet.isColumnHidden(1), "Column B should be hidden");
                assertFalse(sheet.isColumnHidden(2), "Column C should be visible");
                assertTrue(sheet.isColumnHidden(3), "Column D should be hidden");
                assertFalse(sheet.isColumnHidden(4), "Column E should be visible");
            }
        }

        @Test
        void hidden_withExcelSheetWriter_columnConfigHidden() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Sheet1")
                        .column("Visible", s -> s)
                        .column("Hidden", s -> "secret", c -> c.hidden())
                        .write(Stream.of("data"));
                workbook.finish().consumeOutputStream(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertFalse(sheet.isColumnHidden(0), "First column should be visible in ExcelSheetWriter");
                assertTrue(sheet.isColumnHidden(1), "Hidden column via ColumnConfig.hidden() should be hidden");
            }
        }

        @Test
        void hidden_columnStillHasData() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Name", s -> s)
                    .addColumn("Secret", s -> "hidden-" + s, c -> c.hidden())
                    .write(Stream.of("Alice"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertTrue(sheet.isColumnHidden(1), "Column should be hidden");
                assertEquals("hidden-Alice", sheet.getRow(1).getCell(1).getStringCellValue(),
                        "Hidden column should still contain data");
            }
        }

        @Test
        void hidden_withSheetRollover_allSheetsHaveHiddenColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Data")
                        .column("Visible", s -> s)
                        .column("Hidden", s -> "h", c -> c.hidden())
                        .maxRows(3)
                        .write(Stream.of("A", "B", "C", "D", "E", "F", "G"));
                workbook.finish().consumeOutputStream(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getNumberOfSheets() >= 2,
                        "Should have multiple sheets from rollover");
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    var sheet = wb.getSheetAt(i);
                    assertFalse(sheet.isColumnHidden(0),
                            "Visible column should not be hidden on sheet " + i);
                    assertTrue(sheet.isColumnHidden(1),
                            "Hidden column should be hidden on all rollover sheets, sheet " + i);
                }
            }
        }

        @Test
        void hidden_combinedWithOtherStyling() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Styled Hidden", s -> s, c -> c
                            .hidden()
                            .bold(true)
                            .backgroundColor(ExcelColor.LIGHT_BLUE))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertTrue(sheet.isColumnHidden(0),
                        "Column should be hidden even with other styling applied");
                // Data should still be present
                assertEquals("data", sheet.getRow(1).getCell(0).getStringCellValue());
            }
        }
    }

    // ============================================================
    // Feature 2: Rich Text (ExcelRichText + ExcelDataType.RICH_TEXT)
    // ============================================================
    @Nested
    class RichTextTests {

        @Test
        void richText_plainText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText().text("Hello World"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Hello World", cell.getRichStringCellValue().getString(),
                        "Plain text segment should produce correct cell value");
            }
        }

        @Test
        void richText_boldText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText().text("Normal ").bold("Bold"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Normal Bold", cell.getRichStringCellValue().getString(),
                        "Bold rich text should produce correct concatenated string");
            }
        }

        @Test
        void richText_italicText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText().text("Normal ").italic("Italic"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Normal Italic", cell.getRichStringCellValue().getString(),
                        "Italic rich text should produce correct concatenated string");
            }
        }

        @Test
        void richText_styledWithMultipleOptions() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText()
                                    .text("Normal ")
                                    .styled("Fancy", style -> style
                                            .color(255, 0, 0)
                                            .fontSize(16)
                                            .underline(true)
                                            .strikethrough(true)),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Normal Fancy", cell.getRichStringCellValue().getString(),
                        "Styled rich text should produce correct concatenated string");
            }
        }

        @Test
        void richText_multipleSegments() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText()
                                    .text("Hello ")
                                    .bold("World")
                                    .text(" - ")
                                    .italic("End"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                var rts = (XSSFRichTextString) cell.getRichStringCellValue();
                assertEquals("Hello World - End", rts.getString(),
                        "Multiple segments should be concatenated correctly");
            }
        }

        @Test
        void richText_toStringReturnsPlainText() {
            ExcelRichText rt = new ExcelRichText()
                    .text("Hello ")
                    .bold("World")
                    .italic("!")
                    .styled("end", s -> s.color(255, 0, 0).fontSize(20));

            assertEquals("Hello World!end", rt.toString(),
                    "toString() should return plain text without formatting");
        }

        @Test
        void richText_dataTypeWritesCorrectly() throws IOException {
            ExcelRichText rt = new ExcelRichText()
                    .bold("Important: ")
                    .text("read this");

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> rt, c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertTrue(cell.getRichStringCellValue() instanceof XSSFRichTextString,
                        "RICH_TEXT data type should produce XSSFRichTextString");
                assertEquals("Important: read this", cell.getRichStringCellValue().getString());
            }
        }

        @Test
        void richText_fallbackWhenNotExcelRichText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> "plain string value",
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("plain string value", cell.getStringCellValue(),
                        "RICH_TEXT type should fall back to String.valueOf when value is not ExcelRichText");
            }
        }

        @Test
        void richText_withExcelColorPreset() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText()
                                    .styled("Red text", style -> style.color(ExcelColor.RED))
                                    .styled("Blue text", style -> style.color(ExcelColor.BLUE)),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Red textBlue text", cell.getRichStringCellValue().getString(),
                        "Rich text with ExcelColor presets should have correct content");
            }
        }

        @Test
        void richText_inExcelSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("RichSheet")
                        .column("Col", s -> new ExcelRichText()
                                        .text("Normal ")
                                        .bold("Bold"),
                                c -> c.type(ExcelDataType.RICH_TEXT))
                        .write(Stream.of("data"));
                workbook.finish().consumeOutputStream(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Normal Bold", cell.getRichStringCellValue().getString(),
                        "Rich text should work in ExcelSheetWriter (ExcelWorkbook)");
            }
        }

        @Test
        void richText_emptySegments() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> new ExcelRichText()
                                    .text("")
                                    .bold("")
                                    .text("Actual"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("row"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Actual", cell.getRichStringCellValue().getString(),
                        "Empty segments should not affect the final text content");
            }
        }
    }

    // ============================================================
    // Feature 3: Print Setup (ExcelPrintSetup)
    // ============================================================
    @Nested
    class PrintSetupTests {

        @Test
        void printSetup_landscapeOrientation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.orientation(ExcelPrintSetup.Orientation.LANDSCAPE))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getSheetAt(0).getPrintSetup().getLandscape(),
                        "Landscape orientation should set landscape to true");
            }
        }

        @Test
        void printSetup_portraitOrientation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.orientation(ExcelPrintSetup.Orientation.PORTRAIT))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertFalse(wb.getSheetAt(0).getPrintSetup().getLandscape(),
                        "Portrait orientation should set landscape to false");
            }
        }

        @Test
        void printSetup_paperSizeA4() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.paperSize(ExcelPrintSetup.PaperSize.A4))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(PrintSetup.A4_PAPERSIZE,
                        wb.getSheetAt(0).getPrintSetup().getPaperSize(),
                        "Paper size should be A4");
            }
        }

        @Test
        void printSetup_allMargins() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.margins(0.5, 0.6, 0.7, 0.8))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps
                            .leftMargin(1.0)
                            .rightMargin(1.5)
                            .topMargin(2.0)
                            .bottomMargin(2.5))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps
                            .headerCenter("My Report")
                            .footerCenter("Page &P of &N"))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps
                            .headerLeft("HL")
                            .headerCenter("HC")
                            .headerRight("HR")
                            .footerLeft("FL")
                            .footerCenter("FC")
                            .footerRight("FR"))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.repeatHeaderRows())
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.repeatRows(0, 1))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.fitToPageWidth())
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .printSetup(ps -> ps.fitToPage(2, 3))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

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
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("PrintSheet")
                        .column("Col", s -> s)
                        .printSetup(ps -> ps
                                .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                                .paperSize(ExcelPrintSetup.PaperSize.A4)
                                .margins(0.5, 0.5, 0.75, 0.75)
                                .headerCenter("Report Title")
                                .footerCenter("Page &P"))
                        .write(Stream.of("data"));
                workbook.finish().consumeOutputStream(out);
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
            try (ExcelWorkbook workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Data")
                        .column("Col", s -> s)
                        .maxRows(3)
                        .printSetup(ps -> ps
                                .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                                .paperSize(ExcelPrintSetup.PaperSize.LETTER)
                                .headerCenter("Multi-Sheet Report"))
                        .write(Stream.of("A", "B", "C", "D", "E", "F", "G"));
                workbook.finish().consumeOutputStream(out);
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
            new ExcelWriter<String>()
                    .addColumn("A", s -> s)
                    .addColumn("B", s -> s.toUpperCase())
                    .autoFilter(true)
                    .freezePane(1)
                    .printSetup(ps -> ps
                            .orientation(ExcelPrintSetup.Orientation.LANDSCAPE)
                            .fitToPageWidth()
                            .margins(0.25, 0.25, 0.5, 0.5)
                            .headerCenter("Combined Features Report")
                            .footerRight("&D"))
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
}
