package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive tests for features added in v0.7.2:
 * 1. Workbook Protection
 * 2. Header Row Styling
 * 3. Default Column Style
 * 4. Summary/Footer Row DSL
 * 5. Named Ranges
 * 6. List Validation from Range
 */
class NewFeaturesV072Test {

    // ============================================================
    // Feature 1: Workbook Protection
    // ============================================================
    @Nested
    class WorkbookProtectionTests {

        @Test
        void excelWriter_protectWorkbook_structureIsLocked() throws Exception {
            var handler = new ExcelWriter<String>()
                    .protectWorkbook("secret")
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                assertTrue(wb.isStructureLocked(),
                        "Workbook structure should be locked after protectWorkbook");
            }
        }

        @Test
        void excelWorkbook_protectWorkbook_structureIsLocked() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
                workbook.protectWorkbook("pwd");
                workbook.<String>sheet("Sheet1")
                        .column("Name", s -> s)
                        .write(Stream.of("Bob"));
                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                assertTrue(wb.isStructureLocked(),
                        "ExcelWorkbook structure should be locked after protectWorkbook");
            }
        }

        @Test
        void withoutProtectWorkbook_structureIsNotLocked() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                assertFalse(wb.isStructureLocked(),
                        "Workbook structure should NOT be locked when protectWorkbook is not called");
            }
        }

        @Test
        void protectWorkbook_combinedWithProtectSheet_bothApply() throws Exception {
            var handler = new ExcelWriter<String>()
                    .protectWorkbook("workbookPwd")
                    .protectSheet("sheetPwd")
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Data"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                assertTrue(wb.isStructureLocked(),
                        "Workbook structure should be locked");
                assertTrue(wb.getSheetAt(0).isSheetLocked(),
                        "Sheet should be protected when protectSheet is also called");
            }
        }
    }

    // ============================================================
    // Feature 2: Header Row Styling
    // ============================================================
    @Nested
    class HeaderRowStylingTests {

        @Test
        void headerFontName_only_verifiesFontNameOnHeaderCell() throws Exception {
            var handler = new ExcelWriter<String>()
                    .headerFontName("Arial")
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Arial", font.getFontName());
            }
        }

        @Test
        void headerFontSize_only_verifiesFontSizeOnHeaderCell() throws Exception {
            var handler = new ExcelWriter<String>()
                    .headerFontSize(18)
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals(18, font.getFontHeightInPoints());
            }
        }

        @Test
        void headerFontNameAndSize_bothApplied() throws Exception {
            var handler = new ExcelWriter<String>()
                    .headerFontName("Courier New")
                    .headerFontSize(14)
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Courier New", font.getFontName());
                assertEquals(14, font.getFontHeightInPoints());
            }
        }

        @Test
        void excelWorkbook_headerFontNameAndSize_appliedToSheet() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook()) {
                workbook.headerFontName("Times New Roman").headerFontSize(16);
                workbook.<String>sheet("Sheet1")
                        .column("Name", s -> s)
                        .write(Stream.of("Bob"));
                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Times New Roman", font.getFontName());
                assertEquals(16, font.getFontHeightInPoints());
            }
        }

        @Test
        void defaultHeaderStyle_isBold11pt() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertTrue(font.getBold(), "Default header font should be bold");
                assertEquals(11, font.getFontHeightInPoints(),
                        "Default header font size should be 11pt");
            }
        }

        @Test
        void headerFontSize_zeroThrowsIllegalArgumentException() {
            var writer = new ExcelWriter<String>();
            assertThrows(IllegalArgumentException.class, () -> writer.headerFontSize(0));
        }

        @Test
        void headerFontSize_negativeThrowsIllegalArgumentException() {
            var writer = new ExcelWriter<String>();
            assertThrows(IllegalArgumentException.class, () -> writer.headerFontSize(-5));
        }

        @Test
        void multipleColumns_allHeadersHaveSameFont() throws Exception {
            var handler = new ExcelWriter<String>()
                    .headerFontName("Georgia")
                    .headerFontSize(12)
                    .addColumn("Col1", s -> s)
                    .addColumn("Col2", s -> s)
                    .addColumn("Col3", s -> s)
                    .write(Stream.of("data"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var headerRow = wb.getSheetAt(0).getRow(0);
                for (int i = 0; i < 3; i++) {
                    var font = wb.getFontAt(headerRow.getCell(i).getCellStyle().getFontIndex());
                    assertEquals("Georgia", font.getFontName(),
                            "Header cell " + i + " should have font Georgia");
                    assertEquals(12, font.getFontHeightInPoints(),
                            "Header cell " + i + " should have size 12");
                }
            }
        }
    }

    // ============================================================
    // Feature 3: Default Column Style
    // ============================================================
    @Nested
    class DefaultColumnStyleTests {

        @Test
        void excelWriter_defaultStyle_allColumnsInheritBoldFontNameAlignment() throws Exception {
            var handler = new ExcelWriter<String>()
                    .defaultStyle(d -> d.bold(true).fontName("Arial").alignment(HorizontalAlignment.LEFT))
                    .addColumn("Name", s -> s)
                    .addColumn("Value", s -> s)
                    .write(Stream.of("Test"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1); // data row
                for (int i = 0; i < 2; i++) {
                    var cell = row.getCell(i);
                    var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
                    assertTrue(font.getBold(), "Column " + i + " should be bold from default");
                    assertEquals("Arial", font.getFontName(),
                            "Column " + i + " should have Arial from default");
                    assertEquals(HorizontalAlignment.LEFT, cell.getCellStyle().getAlignment(),
                            "Column " + i + " should have LEFT alignment from default");
                }
            }
        }

        @Test
        void columnLevelOverride_winsOverDefault() throws Exception {
            var handler = new ExcelWriter<String>()
                    .defaultStyle(d -> d.bold(true).fontName("Arial"))
                    .column("Name", s -> s)
                        .bold(false)  // override bold
                    .write(Stream.of("Test"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
                assertFalse(font.getBold(), "Column override (bold=false) should win over default");
                assertEquals("Arial", font.getFontName(),
                        "Non-overridden fontName should come from default");
            }
        }

        @Test
        void excelSheetWriter_defaultStyle_appliedToColumn() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Data")
                        .defaultStyle(d -> d.fontName("Courier New"))
                        .column("Name", s -> s)
                        .write(Stream.of("Test"));
                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
                assertEquals("Courier New", font.getFontName());
            }
        }

        @Test
        void defaultStyle_withVerticalAlignmentWrapTextFontSizeFontColor() throws Exception {
            var handler = new ExcelWriter<String>()
                    .defaultStyle(d -> d
                            .verticalAlignment(VerticalAlignment.TOP)
                            .wrapText(true)
                            .fontSize(14)
                            .fontColor(255, 0, 0))
                    .addColumn("Col", s -> s)
                    .write(Stream.of("Test"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                var style = cell.getCellStyle();
                assertEquals(VerticalAlignment.TOP, style.getVerticalAlignment());
                assertTrue(style.getWrapText(), "wrapText should be enabled from default");
                var font = wb.getFontAt(style.getFontIndex());
                assertEquals(14, font.getFontHeightInPoints());
            }
        }

        @Test
        void noDefaultStyle_columnsUseNormalDefaults() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Col", s -> s)
                    .write(Stream.of("Test"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                var style = cell.getCellStyle();
                // default alignment is CENTER per ColumnStyleConfig
                assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
            }
        }

        @Test
        void applyDefaults_fillsNullFields() {
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
            defaults.bold(true).fontName("Arial").fontSize(14).alignment(HorizontalAlignment.LEFT);

            var config = new ExcelSheetWriter.ColumnConfig<String>();
            // Only set bold -- everything else should come from defaults
            config.bold(false);

            config.applyDefaults(defaults);

            assertEquals(false, config.bold, "Explicit bold=false should be kept");
            assertEquals("Arial", config.fontName, "fontName should come from defaults");
            assertEquals(14, config.fontSize, "fontSize should come from defaults");
            assertEquals(HorizontalAlignment.LEFT, config.alignment,
                    "alignment should come from defaults");
        }

        @Test
        void applyDefaults_doesNotOverrideExplicitValues() {
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
            defaults.bold(true).fontName("Arial");

            var config = new ExcelSheetWriter.ColumnConfig<String>();
            config.bold(false).fontName("Courier");

            config.applyDefaults(defaults);

            assertEquals(false, config.bold, "Explicit bold=false should not be overridden");
            assertEquals("Courier", config.fontName, "Explicit fontName should not be overridden");
        }

        @Test
        void applyDefaults_alignmentSetFlagTracked() {
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
            defaults.alignment(HorizontalAlignment.RIGHT);

            var configWithoutAlignment = new ExcelSheetWriter.ColumnConfig<String>();
            // alignmentSet is false by default, so defaults should apply
            configWithoutAlignment.applyDefaults(defaults);
            assertEquals(HorizontalAlignment.RIGHT, configWithoutAlignment.alignment,
                    "Default alignment should apply when column did not set alignment");
            assertTrue(configWithoutAlignment.alignmentSet,
                    "alignmentSet flag should be propagated from defaults");

            var configWithAlignment = new ExcelSheetWriter.ColumnConfig<String>();
            configWithAlignment.alignment(HorizontalAlignment.LEFT);
            configWithAlignment.applyDefaults(defaults);
            assertEquals(HorizontalAlignment.LEFT, configWithAlignment.alignment,
                    "Explicit alignment should not be overridden by defaults");
        }

        @Test
        void applyDefaults_nullDefaultFieldsDoNotOverwrite() {
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
            // defaults has no explicit settings

            var config = new ExcelSheetWriter.ColumnConfig<String>();
            config.bold(true).fontName("Helvetica");

            config.applyDefaults(defaults);

            assertEquals(true, config.bold, "Existing bold should remain");
            assertEquals("Helvetica", config.fontName, "Existing fontName should remain");
        }
    }

    // ============================================================
    // Feature 4: Summary/Footer Row DSL
    // ============================================================
    @Nested
    class SummaryTests {

        record Item(String name, int price, int qty) {}

        @Test
        void summary_singleSum_generatesFormulaAndLabel() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Qty", i -> i.qty(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Price").sum("Qty"))
                    .write(Stream.of(new Item("A", 100, 5), new Item("B", 200, 3)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: header, Row 1-2: data, Row 3: summary
                var summaryRow = sheet.getRow(3);
                assertNotNull(summaryRow, "Summary row should exist at row 3");
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
                assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
                assertEquals("SUM(C2:C3)", summaryRow.getCell(2).getCellFormula());
            }
        }

        @Test
        void summary_multipleColumnsSum_allFormulasGenerated() throws Exception {
            record Data(String a, int b, int c, int d) {}
            var handler = new ExcelWriter<Data>()
                    .addColumn("A", Data::a)
                    .addColumn("B", x -> x.b(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("C", x -> x.c(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("D", x -> x.d(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Sum").sum("B").sum("C").sum("D"))
                    .write(Stream.of(new Data("x", 1, 2, 3), new Data("y", 4, 5, 6)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(3);
                assertEquals("Sum", summaryRow.getCell(0).getStringCellValue());
                assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
                assertEquals("SUM(C2:C3)", summaryRow.getCell(2).getCellFormula());
                assertEquals("SUM(D2:D3)", summaryRow.getCell(3).getCellFormula());
            }
        }

        @Test
        void summary_multipleOps_separateRowsWithCorrectLabels() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.sum("Price").average("Price"))
                    .write(Stream.of(new Item("A", 100, 5)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: header, Row 1: data, Row 2: SUM, Row 3: AVERAGE
                var sumRow = sheet.getRow(2);
                assertEquals("Sum", sumRow.getCell(0).getStringCellValue());
                assertEquals("SUM(B2:B2)", sumRow.getCell(1).getCellFormula());

                var avgRow = sheet.getRow(3);
                assertEquals("Average", avgRow.getCell(0).getStringCellValue());
                assertEquals("AVERAGE(B2:B2)", avgRow.getCell(1).getCellFormula());
            }
        }

        @Test
        void summary_countMinMax_generateCorrectFormulas() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.count("Price").min("Price").max("Price"))
                    .write(Stream.of(new Item("A", 10, 1), new Item("B", 20, 2), new Item("C", 30, 3)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: header, Row 1-3: data, Row 4: COUNT, Row 5: MIN, Row 6: MAX
                var countRow = sheet.getRow(4);
                assertEquals("Count", countRow.getCell(0).getStringCellValue());
                assertEquals("COUNT(B2:B4)", countRow.getCell(1).getCellFormula());

                var minRow = sheet.getRow(5);
                assertEquals("Min", minRow.getCell(0).getStringCellValue());
                assertEquals("MIN(B2:B4)", minRow.getCell(1).getCellFormula());

                var maxRow = sheet.getRow(6);
                assertEquals("Max", maxRow.getCell(0).getStringCellValue());
                assertEquals("MAX(B2:B4)", maxRow.getCell(1).getCellFormula());
            }
        }

        @Test
        void summary_customLabelText_singleOp() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Grand Total").sum("Price"))
                    .write(Stream.of(new Item("A", 100, 5)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(2);
                assertEquals("Grand Total", summaryRow.getCell(0).getStringCellValue(),
                        "Custom label should be used when only one op is configured");
            }
        }

        @Test
        void summary_labelInSpecificColumn_verifyPlacement() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Qty", i -> i.qty(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Price", "Subtotal").sum("Qty"))
                    .write(Stream.of(new Item("A", 100, 5)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(2);
                // Label should be in column index 1 (the "Price" column)
                assertEquals("Subtotal", summaryRow.getCell(1).getStringCellValue());
                assertEquals("SUM(C2:C2)", summaryRow.getCell(2).getCellFormula());
            }
        }

        @Test
        void excelSheetWriter_withSummary_generatesFormulaRow() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook()) {
                workbook.<Item>sheet("Sales")
                        .column("Name", Item::name)
                        .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                        .summary(s -> s.label("Total").sum("Price"))
                        .write(Stream.of(new Item("X", 50, 1), new Item("Y", 75, 2)));
                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                var summaryRow = sheet.getRow(3);
                assertNotNull(summaryRow);
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
                assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
            }
        }

        @Test
        void summary_withBeforeHeader_formulaRangeAccountsForOffset() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .beforeHeader(ctx -> {
                        // Write a title row before the header
                        var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                        row.createCell(0).setCellValue("Report Title");
                        return ctx.getCurrentRow() + 1;
                    })
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Price"))
                    .write(Stream.of(new Item("A", 100, 5), new Item("B", 200, 3)));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: title, Row 1: header, Row 2-3: data, Row 4: summary
                var summaryRow = sheet.getRow(4);
                assertNotNull(summaryRow, "Summary row should be at row 4 after beforeHeader offset");
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
                // Data starts at row 3 (1-based), ends at row 4 (1-based)
                assertEquals("SUM(B3:B4)", summaryRow.getCell(1).getCellFormula(),
                        "Formula range should account for the beforeHeader offset");
            }
        }

        @Test
        void summary_emptyData_summaryRowStillCreated() throws Exception {
            var handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Price"))
                    .write(Stream.empty());

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: header, Row 1: summary (no data rows)
                var summaryRow = sheet.getRow(1);
                assertNotNull(summaryRow, "Summary row should exist even with no data");
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
                // The formula will reference an empty range (B2:B1), which is valid in Excel
                assertNotNull(summaryRow.getCell(1).getCellFormula());
            }
        }
    }

    // ============================================================
    // Feature 5: Named Ranges
    // ============================================================
    @Nested
    class NamedRangeTests {

        @Test
        void namedRange_withReferenceString_verifyNameAndFormula() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Category", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("Categories", "Sheet0!$A$2:$A$4");
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("A", "B", "C"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var name = wb.getName("Categories");
                assertNotNull(name, "Named range 'Categories' should exist");
                assertEquals("Sheet0!$A$2:$A$4", name.getRefersToFormula());
            }
        }

        @Test
        void namedRange_withColumnIndices_generatesCorrectReference() throws Exception {
            var handler = new ExcelWriter<String>()
                    .sheetName("Data")
                    .addColumn("Items", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("ItemList", 0, 1, 3); // col A, rows 2-4 (0-based: 1-3)
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("X", "Y", "Z"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var name = wb.getName("ItemList");
                assertNotNull(name, "Named range 'ItemList' should exist");
                assertEquals("'Data'!$A$2:$A$4", name.getRefersToFormula());
            }
        }

        @Test
        void multipleNamedRanges_allCreated() throws Exception {
            var handler = new ExcelWriter<String>()
                    .sheetName("Ref")
                    .addColumn("Col", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("Range1", "'Ref'!$A$2:$A$3");
                        ctx.namedRange("Range2", "'Ref'!$A$2:$A$4");
                        ctx.namedRange("Range3", 0, 1, 3);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("A", "B", "C"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                assertNotNull(wb.getName("Range1"), "Range1 should exist");
                assertNotNull(wb.getName("Range2"), "Range2 should exist");
                assertNotNull(wb.getName("Range3"), "Range3 should exist");
                assertEquals("'Ref'!$A$2:$A$3", wb.getName("Range1").getRefersToFormula());
                assertEquals("'Ref'!$A$2:$A$4", wb.getName("Range2").getRefersToFormula());
                assertEquals("'Ref'!$A$2:$A$4", wb.getName("Range3").getRefersToFormula());
            }
        }

        @Test
        void namedRange_inExcelWorkbookMultiSheet_verifyAcrossSheets() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Categories")
                        .column("Cat", s -> s)
                        .afterData(ctx -> {
                            ctx.namedRange("CatList", 0, 1, 3);
                            return ctx.getCurrentRow();
                        })
                        .write(Stream.of("Cat1", "Cat2", "Cat3"));

                workbook.<String>sheet("Products")
                        .column("Product", s -> s)
                        .afterData(ctx -> {
                            ctx.namedRange("ProdList", 0, 1, 2);
                            return ctx.getCurrentRow();
                        })
                        .write(Stream.of("Prod1", "Prod2"));

                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var catName = wb.getName("CatList");
                assertNotNull(catName, "CatList named range should exist");
                assertEquals("'Categories'!$A$2:$A$4", catName.getRefersToFormula());

                var prodName = wb.getName("ProdList");
                assertNotNull(prodName, "ProdList named range should exist");
                assertEquals("'Products'!$A$2:$A$3", prodName.getRefersToFormula());
            }
        }
    }

    // ============================================================
    // Feature 6: List Validation from Range
    // ============================================================
    @Nested
    class ListValidationFromRangeTests {

        @Test
        void listFromRange_createsValidationOnSheet() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Category", s -> s, c -> c
                            .validation(ExcelValidation.listFromRange("Sheet2!$A$1:$A$5")))
                    .write(Stream.of("Test"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertFalse(validations.isEmpty(),
                        "Data validations should exist on the sheet");
            }
        }

        @Test
        void listFromRange_withErrorConfiguration_verifyErrorBox() throws Exception {
            var handler = new ExcelWriter<String>()
                    .addColumn("Status", s -> s, c -> c
                            .validation(ExcelValidation.listFromRange("Statuses!$A$1:$A$3")
                                    .errorTitle("Invalid Status")
                                    .errorMessage("Please select a valid status.")
                                    .showError(true)))
                    .write(Stream.of("Active"));

            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertFalse(validations.isEmpty());
                var dv = validations.get(0);
                assertTrue(dv.getShowErrorBox(),
                        "Error box should be shown");
            }
        }

        @Test
        void listFromRange_factoryMethodReturnsNonNull() {
            var v = ExcelValidation.listFromRange("Options!$B$1:$B$10");
            assertNotNull(v, "listFromRange should return a non-null ExcelValidation");
        }

        @Test
        void listFromRange_combinedWithExcelSheetWriter_validationApplied() throws Exception {
            var baos = new ByteArrayOutputStream();
            try (var workbook = new ExcelWorkbook()) {
                workbook.<String>sheet("Data")
                        .column("Type", s -> s, c -> c
                                .validation(ExcelValidation.listFromRange("Lookup!$A$1:$A$5")))
                        .write(Stream.of("Item1"));
                workbook.finish().consumeOutputStream(baos);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertFalse(validations.isEmpty(),
                        "Data validations should be applied via ExcelSheetWriter");
            }
        }
    }
}
