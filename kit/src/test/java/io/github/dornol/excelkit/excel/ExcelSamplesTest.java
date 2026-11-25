package io.github.dornol.excelkit.excel;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The workbooks of {@link ExcelSamples} are what a person (or an agent driving Excel) checks against {@code manifest.json}.
 * This holds the manifest to what POI can see of the same files, so that Excel is asked only about things that are true
 * of the file; the rest of each check (does Excel open it cleanly, does the filter reach the data) is for Excel to say.
 */
class ExcelSamplesTest {

    @TempDir
    static Path dir;
    static List<ExcelSamples.Sample> samples;

    @BeforeAll
    static void generate() throws Exception {
        samples = ExcelSamples.generate(dir);
    }

    private static XSSFWorkbook open(String file) throws Exception {
        try (InputStream in = Files.newInputStream(dir.resolve(file))) {
            return new XSSFWorkbook(in);
        }
    }

    private static String filter(XSSFSheet sheet) {
        return sheet.getCTWorksheet().isSetAutoFilter() ? sheet.getCTWorksheet().getAutoFilter().getRef() : null;
    }

    @Test
    void everySampleHasADescriptionInTheManifestAndTheChecklist() throws Exception {
        String manifest = Files.readString(dir.resolve("manifest.json"));
        String checklist = Files.readString(dir.resolve("CHECKLIST.md"));
        assertEquals(33, samples.size());
        for (var sample : samples) {
            assertTrue(Files.exists(dir.resolve(sample.file())), sample.file());
            assertTrue(manifest.contains("\"" + sample.file() + "\""), sample.file());
            assertTrue(checklist.contains(sample.file()), sample.file());
            assertFalse(sample.checks().isEmpty(), sample.file());
        }
    }

    @Test
    void everyUnencryptedSampleOpensInPoi() throws Exception {
        for (var sample : samples) {
            if (sample.file().startsWith("07-") || sample.file().startsWith("08-")) continue;
            try (XSSFWorkbook wb = open(sample.file())) {
                assertTrue(wb.getNumberOfSheets() >= 1, sample.file());
            }
        }
    }

    @Test
    void theFiltersAreWhereTheManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("01-autofilter-plain.xlsx")) {
            assertEquals("A1:B8", filter(wb.getSheet("Report")));
            assertEquals("Sum", wb.getSheet("Report").getRow(8).getCell(0).getStringCellValue());
            assertEquals("Average", wb.getSheet("Report").getRow(9).getCell(0).getStringCellValue());
        }
        try (XSSFWorkbook wb = open("02-autofilter-group-header.xlsx")) {
            assertEquals("A2:B5", filter(wb.getSheet("Report")));
            assertEquals("A1:B1", wb.getSheet("Report").getMergedRegion(0).formatAsString());
        }
        try (XSSFWorkbook wb = open("03-autofilter-rollover.xlsx")) {
            assertEquals(List.of("Report", "Report (2)", "Report (3)"), List.of(wb.getSheetName(0), wb.getSheetName(1), wb.getSheetName(2)));
            assertEquals("A1:B5", filter(wb.getSheet("Report")));
            assertEquals("A1:B5", filter(wb.getSheet("Report (2)")));
            assertEquals("A1:B3", filter(wb.getSheet("Report (3)")));
        }
    }

    @Test
    void theTablesCarryTheirFilterOrNot() throws Exception {
        try (XSSFWorkbook wb = open("04-table-with-autofilter.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("A1:B7", sheet.getTables().get(0).getCellReferences().formatAsString());
            assertNotNull(sheet.getTables().get(0).getCTTable().getAutoFilter());
            assertNull(filter(sheet), "no sheet-level filter over the table");
        }
        try (XSSFWorkbook wb = open("05-table-without-autofilter.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("A1:B7", sheet.getTables().get(0).getCellReferences().formatAsString());
            assertFalse(sheet.getTables().get(0).getCTTable().isSetAutoFilter());
            assertNull(filter(sheet));
        }
    }

    @Test
    void theChartsSitBelowTheirSummaryRows() throws Exception {
        try (XSSFWorkbook wb = open("06-chart-rollover-summary.xlsx")) {
            int[] summaryLastRow1Based = {7, 7, 5};
            for (int s = 0; s < 3; s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                assertEquals(1, sheet.getDrawingPatriarch().getCharts().size(), "sheet " + s);
                int topRow1Based = sheet.getDrawingPatriarch().getCTDrawing().getTwoCellAnchorArray(0).getFrom().getRow() + 1;
                assertTrue(topRow1Based > summaryLastRow1Based[s], "sheet " + s + ": chart starts at row " + topRow1Based);
                assertEquals("Sum", sheet.getRow(summaryLastRow1Based[s] - 2).getCell(0).getStringCellValue(), "sheet " + s);
            }
            // the last sheet holds the items 8 and 9: the average is 8.5 and reads so, the sum 17
            XSSFSheet last = wb.getSheetAt(2);
            assertEquals("#,##0", last.getRow(3).getCell(1).getCellStyle().getDataFormatString());
            assertEquals("#,##0.0", last.getRow(4).getCell(1).getCellStyle().getDataFormatString());
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertEquals(17, evaluator.evaluate(last.getRow(3).getCell(1)).getNumberValue());
            assertEquals(8.5, evaluator.evaluate(last.getRow(4).getCell(1)).getNumberValue());
        }
    }

    @Test
    void theEncryptedSamplesOpenWithTheirPasswordOnly() throws Exception {
        for (String[] sample : new String[][]{{"07-encrypted-password-secret.xlsx", "secret"}, {"08-encrypted-late-password.xlsx", "Late-Binding1"}}) {
            try (POIFSFileSystem fs = new POIFSFileSystem(Files.newInputStream(dir.resolve(sample[0])))) {
                Decryptor dec = Decryptor.getInstance(new EncryptionInfo(fs));
                assertFalse(dec.verifyPassword("wrong"), sample[0]);
                assertTrue(dec.verifyPassword(sample[1]), sample[0]);
                try (XSSFWorkbook wb = new XSSFWorkbook(dec.getDataStream(fs))) {
                    assertEquals("item 0", wb.getSheet("Report").getRow(1).getCell(0).getStringCellValue());
                }
            }
        }
    }

    @Test
    void theFormatsAreTheOnesTheManifestExpectsToSee() throws Exception {
        try (XSSFWorkbook wb = open("09-formats-and-dates.xlsx")) {
            var row = wb.getSheet("Report").getRow(1);
            assertEquals("yyyy-mm-dd", row.getCell(0).getCellStyle().getDataFormatString());
            assertEquals("yyyy-mm-dd hh:mm:ss", row.getCell(2).getCellStyle().getDataFormatString());
            assertEquals("yyyy-mm-dd hh:mm:ss.000", row.getCell(3).getCellStyle().getDataFormatString());
            assertEquals("#,##0", row.getCell(4).getCellStyle().getDataFormatString());
            assertEquals("#,##0.00", row.getCell(5).getCellStyle().getDataFormatString());
            assertEquals(CellType.STRING, row.getCell(6).getCellType());
            assertEquals("0012", row.getCell(6).getStringCellValue());
            assertEquals("@", row.getCell(6).getCellStyle().getDataFormatString(), "the Text format keeps 0012 when the cell is edited");
            assertEquals(CellType.FORMULA, row.getCell(7).getCellType());
            assertEquals("E2*2", row.getCell(7).getCellFormula());
            assertEquals("E2*2", row.getCell(8).getCellFormula());
            assertEquals(2469134.0, row.getCell(8).getNumericCellValue(), "the result is stored next to the formula");
            assertEquals("#,##0", row.getCell(8).getCellStyle().getDataFormatString());
            assertEquals(1234567.0, row.getCell(4).getNumericCellValue());
        }
    }

    @Test
    void theProtectionIsWhatTheManifestExpects() throws Exception {
        try (XSSFWorkbook wb = open("10-protection.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertTrue(sheet.getProtect());
            assertFalse(sheet.getRow(1).getCell(0).getCellStyle().getLocked());
            assertFalse(sheet.getRow(2).getCell(0).getCellStyle().getLocked());
            assertTrue(sheet.getRow(1).getCell(1).getCellStyle().getLocked());
            assertTrue(wb.isStructureLocked());
        }
    }

    @Test
    void theValidationsAndFormattingAreThere() throws Exception {
        try (XSSFWorkbook wb = open("11-conditional-and-validation.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(2, sheet.getDataValidations().size());
            assertTrue(sheet.getSheetConditionalFormatting().getNumConditionalFormattings() >= 1);
            int rules = 0;
            for (int i = 0; i < sheet.getSheetConditionalFormatting().getNumConditionalFormattings(); i++) {
                rules += sheet.getSheetConditionalFormatting().getConditionalFormattingAt(i).getNumberOfRules();
            }
            assertTrue(rules >= 2, "a comparison rule and a data bar: " + rules);
        }
    }

    @Test
    void thePicturesAndLinksAreThere() throws Exception {
        try (XSSFWorkbook wb = open("12-images-links-richtext.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(3, sheet.getDrawingPatriarch().getShapes().stream()
                    .filter(s -> s instanceof org.apache.poi.xssf.usermodel.XSSFPicture).count());
            assertEquals(3, sheet.getHyperlinkList().size());
        }
    }

    @Test
    void theWorkbookOfSeveralSheetsIsWhatTheManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("13-multi-sheet-workbook.xlsx")) {
            assertEquals(List.of("Users", "Orders", "Archive", "Archive (2)"),
                    List.of(wb.getSheetName(0), wb.getSheetName(1), wb.getSheetName(2), wb.getSheetName(3)));
            assertEquals("A1:B6", filter(wb.getSheet("Users")));
            assertEquals("Total", wb.getSheet("Orders").getRow(5).getCell(0).getStringCellValue());
            assertEquals(1, wb.getSheet("Archive").getTables().size());
            assertEquals("ArchiveTable_1", wb.getSheet("Archive").getTables().get(0).getName());
            assertEquals("ArchiveTable_2", wb.getSheet("Archive (2)").getTables().get(0).getName());
            assertNotNull(wb.getSheet("Archive").getTables().get(0).getCTTable().getAutoFilter());
            assertEquals(1, wb.getSheet("Users").getPaneInformation().getHorizontalSplitPosition());
        }
    }

    @Test
    void theNamesPanesAndPrintSetupAreThere() throws Exception {
        try (XSSFWorkbook wb = open("14-names-panes-print.xlsx")) {
            assertEquals("Report!$A$2:$A$6", wb.getName("Names").getRefersToFormula().replace("'", ""));
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(1, sheet.getPaneInformation().getHorizontalSplitPosition());
            assertEquals(1, sheet.getPaneInformation().getVerticalSplitPosition());
            assertEquals(PrintOrientation.LANDSCAPE, sheet.getPrintSetup().getOrientation());
        }
    }

    @Test
    void theFilledTemplateAndTheEverythingSampleAreAsDescribed() throws Exception {
        try (XSSFWorkbook wb = open("15-template-filled.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("A1:C1", sheet.getMergedRegion(0).formatAsString());
            assertEquals(1, wb.getAllPictures().size());
            assertEquals(1, sheet.getDrawingPatriarch().getCharts().size());
            assertEquals("SalesTable", sheet.getTables().get(0).getName());
            assertEquals("A4:B7", sheet.getTables().get(0).getCellReferences().formatAsString());
            assertEquals("A", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Filled by the writer", sheet.getRow(8).getCell(0).getStringCellValue());
            assertEquals("item 0", sheet.getRow(11).getCell(0).getStringCellValue());
            assertEquals("second sheet cell", wb.getSheet("Notes").getRow(1).getCell(0).getStringCellValue());
        }
        try (XSSFWorkbook wb = open("16-everything-at-once.xlsx")) {
            assertEquals(3, wb.getNumberOfSheets());
        }
    }

    @Test
    void theGapSampleKeepsItsSummaryBelowAnEmptyRow() throws Exception {
        try (XSSFWorkbook wb = open("17-autofilter-summary-after-gap.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("A1:B8", filter(sheet));
            assertNull(sheet.getRow(8), "row 9 is empty");
            assertEquals("Sum", sheet.getRow(9).getCell(0).getStringCellValue());
            assertEquals("Average", sheet.getRow(10).getCell(0).getStringCellValue());
        }
    }

    @Test
    void theSubtotalSampleCalculatesTheWayItsManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("27-group-subtotal.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertEquals(30, evaluator.evaluate(sheet.getRow(3).getCell(2)).getNumberValue());
            assertEquals(5, evaluator.evaluate(sheet.getRow(5).getCell(2)).getNumberValue());
            assertEquals("SUBTOTAL(9,C2:C6)", sheet.getRow(6).getCell(2).getCellFormula());
            assertEquals(35, evaluator.evaluate(sheet.getRow(6).getCell(2)).getNumberValue(), "not 70: the subtotals are skipped");
            assertEquals("A total", sheet.getRow(3).getCell(1).getStringCellValue());
            assertEquals("A1:C6", filter(sheet));
        }
    }

    @Test
    void theRolloverSubtotalSampleHasTheRowsItsManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("28-subtotal-rollover.xlsx")) {
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertEquals(3, wb.getNumberOfSheets());
            int[][] rows = {{5, 7, 8}, {4, 7, 8}, {4, 5}};
            double[][] values = {{6, 4, 10}, {11, 15, 26}, {19, 19}};
            for (int s = 0; s < 3; s++) {
                for (int i = 0; i < rows[s].length; i++) {
                    assertEquals(values[s][i],
                            evaluator.evaluate(wb.getSheetAt(s).getRow(rows[s][i] - 1).getCell(2)).getNumberValue(),
                            "sheet " + s + " row " + rows[s][i]);
                }
            }
            assertEquals("A1:C7", filter(wb.getSheetAt(0)));
            assertEquals("A1:C4", filter(wb.getSheetAt(2)));
            assertArrayEquals(new int[]{4}, wb.getSheetAt(0).getRowBreaks(), "a break below the subtotal row 5: the new page starts at row 6");
            assertArrayEquals(new int[]{3}, wb.getSheetAt(1).getRowBreaks(), "below the subtotal row 4: the new page starts at row 5");
            assertEquals(0, wb.getSheetAt(2).getRowBreaks().length);
        }
    }

    @Test
    void theProtectionSubtotalSampleLocksTheSubtotalRowsOnly() throws Exception {
        try (XSSFWorkbook wb = open("29-subtotal-protection.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertTrue(sheet.getProtect());
            assertFalse(sheet.getRow(1).getCell(1).getCellStyle().getLocked());
            assertFalse(sheet.getRow(1).getCell(2).getCellStyle().getLocked());
            assertTrue(sheet.getRow(1).getCell(0).getCellStyle().getLocked());
            for (int row : new int[]{3, 5}) {
                for (int col = 0; col < 3; col++) assertTrue(sheet.getRow(row).getCell(col).getCellStyle().getLocked(), row + "," + col);
            }
        }
    }

    @Test
    void theFilterAndRulesSubtotalSampleHoldsTheTotalsItsManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("30-subtotal-filter-and-rules.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertEquals(30, evaluator.evaluate(sheet.getRow(3).getCell(2)).getNumberValue());
            assertEquals(35, evaluator.evaluate(sheet.getRow(6).getCell(2)).getNumberValue());
            assertEquals(65, evaluator.evaluate(sheet.getRow(7).getCell(2)).getNumberValue());
            assertEquals("A", sheet.getRow(3).getCell(0).getStringCellValue(), "the subtotal row shows its key");
            assertEquals("A total", sheet.getRow(3).getCell(1).getStringCellValue());
            assertEquals("A1:C7", filter(sheet));
            var total = sheet.getRow(7).getCell(2);
            assertEquals("#,##0", total.getCellStyle().getDataFormatString(), "the total takes the format of its column");
            assertTrue(wb.getFontAt(total.getCellStyle().getFontIndex()).getBold(), "the total row is bold");
            assertTrue(wb.getFontAt(sheet.getRow(7).getCell(0).getCellStyle().getFontIndex()).getBold(), "across the whole row");
            assertEquals(2, sheet.getDataValidations().size());
            assertEquals(1, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
            var ranges = sheet.getSheetConditionalFormatting().getConditionalFormattingAt(0).getFormattingRanges();
            assertEquals(java.util.List.of("C2:C3", "C5:C6"), java.util.Arrays.stream(ranges)
                    .map(org.apache.poi.ss.util.CellRangeAddress::formatAsString).toList(), "the rule skips the subtotal rows");
        }
    }

    @Test
    void theComboChartSampleHasABarGroupAndTwoLineGroupsOnTwoPairsOfAxes() throws Exception {
        try (XSSFWorkbook wb = open("31-combo-chart.xlsx")) {
            var plot = wb.getSheet("Sales").getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea();
            assertEquals(1, plot.sizeOfBarChartArray());
            assertEquals(2, plot.sizeOfLineChartArray());
            assertEquals(2, plot.sizeOfCatAxArray());
            assertEquals(2, plot.sizeOfValAxArray());
        }
    }

    @Test
    void theEmptyMessageSampleSaysSoOnTheEmptySheetOnly() throws Exception {
        try (XSSFWorkbook wb = open("32-empty-message.xlsx")) {
            XSSFSheet empty = wb.getSheet("Empty");
            assertEquals("No data", empty.getRow(1).getCell(0).getStringCellValue());
            assertEquals("A2:C2", empty.getMergedRegion(0).formatAsString());
            assertEquals("Total", empty.getRow(2).getCell(0).getStringCellValue());
            XSSFSheet rows = wb.getSheet("Rows");
            assertEquals("item 0", rows.getRow(1).getCell(0).getStringCellValue());
            assertEquals(0, rows.getNumMergedRegions());
            assertEquals(3, rows.getLastRowNum());
        }
    }

    @Test
    void theLongValidationFormulaSampleHoldsOneFormulaUnderTheLimitAndOneOver() throws Exception {
        try (XSSFWorkbook wb = open("33-long-validation-formula.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            var validations = sheet.getDataValidations();
            assertEquals(2, validations.size());
            int[] lengths = validations.stream().mapToInt(v -> v.getValidationConstraint().getFormula1().length()).sorted().toArray();
            assertTrue(lengths[0] <= 255, "one formula is under Excel's limit: " + lengths[0]);
            assertTrue(lengths[1] > 255, "one formula is over it: " + lengths[1]);
        }
    }

    // ---- the upload-form samples (18-24) ------------------------------------------------------------------

    private static boolean arrowHidden(XSSFSheet sheet, int validation) {
        return sheet.getCTWorksheet().getDataValidations().getDataValidationArray(validation).getShowDropDown();
    }

    private static String fill(org.apache.poi.ss.usermodel.Cell cell) {
        var style = cell.getCellStyle();
        if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) return "none";
        byte[] rgb = ((org.apache.poi.xssf.usermodel.XSSFColor) style.getFillForegroundColorColor()).getRGB();
        return String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    @Test
    void theDropdownSampleKeepsTheArrowAndTheOptions() throws Exception {
        try (XSSFWorkbook wb = open("18-dropdown-validation.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Form");
            var validations = sheet.getDataValidations();
            assertEquals(3, validations.size());
            for (int i = 0; i < 3; i++) assertFalse(arrowHidden(sheet, i), "validation " + i);
            assertEquals("\"Active,Inactive,Pending\"", validations.get(0).getValidationConstraint().getFormula1());
            var priority = validations.get(1);
            assertFalse(priority.getEmptyCellAllowed());
            assertTrue(priority.getShowPromptBox());
            assertEquals("Priority", priority.getPromptBoxTitle());
            assertEquals("Pick one of the two", priority.getPromptBoxText());
            assertTrue(validations.get(2).getEmptyCellAllowed());
            assertEquals("Codes!$A$2:$A$4", validations.get(2).getValidationConstraint().getFormula1());
            assertEquals(2, wb.getNumberOfSheets());
        }
    }

    @Test
    void theDependentSampleOffersTheCountsTheManifestExpects() throws Exception {
        try (XSSFWorkbook wb = open("19-dependent-dropdown.xlsx")) {
            assertEquals(List.of("Form", "_excelkit_dep_1", "_excelkit_dep_2"),
                    List.of(wb.getSheetName(0), wb.getSheetName(1), wb.getSheetName(2)));
            assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(1));
            assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(2));
            XSSFSheet sheet = wb.getSheet("Form");
            var validations = sheet.getDataValidations();
            assertEquals(3, validations.size(), "Category (dropdown), Item, Sub");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            var scratch = sheet.getRow(1).createCell(25);

            String itemFormula = validations.get(1).getValidationConstraint().getFormula1();
            for (var expected : java.util.Map.of("가", 3, "나", 2, "다", 1).entrySet()) {
                sheet.getRow(1).createCell(0).setCellValue(expected.getKey()); // a new cell: POI 5 ignores setCellValue on a cell it has read from a file
                scratch.setCellFormula("COUNTA(" + itemFormula + ")");
                evaluator = wb.getCreationHelper().createFormulaEvaluator(); // a new one: POI 5 keeps the value of a changed cell it has already read
                assertEquals(expected.getValue(), (int) evaluator.evaluate(scratch).getNumberValue(), expected.getKey());
            }

            String subFormula = validations.get(2).getValidationConstraint().getFormula1();
            assertEquals(org.apache.poi.ss.usermodel.CellType.NUMERIC, sheet.getRow(1).getCell(2).getCellType(), "a number, not text");
            for (var expected : java.util.Map.of(1, 3, 2, 1).entrySet()) {
                sheet.getRow(1).createCell(2).setCellValue((double) expected.getKey()); // a new cell: POI 5 ignores setCellValue on a cell it has read from a file
                scratch.setCellFormula("COUNTA(" + subFormula + ")");
                evaluator = wb.getCreationHelper().createFormulaEvaluator(); // a new one: POI 5 keeps the value of a changed cell it has already read
                assertEquals(expected.getValue(), (int) evaluator.evaluate(scratch).getNumberValue(), "code " + expected.getKey());
            }
        }
    }

    @Test
    void theHiddenSheetSampleOpensOnTheVisibleSheetAtTheZoomAsked() throws Exception {
        try (XSSFWorkbook wb = open("20-hidden-sheets-and-zoom.xlsx")) {
            assertEquals(List.of("Codes", "Secret", "Main"), List.of(wb.getSheetName(0), wb.getSheetName(1), wb.getSheetName(2)));
            assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(0));
            assertEquals(SheetVisibility.VERY_HIDDEN, wb.getSheetVisibility(1));
            assertEquals(SheetVisibility.VISIBLE, wb.getSheetVisibility(2));
            assertEquals(2, wb.getActiveSheetIndex());
            assertFalse(wb.getSheetAt(0).isSelected(), "a hidden sheet is not selected");
            assertEquals(90, wb.getSheet("Main").getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale());
        }
    }

    @Test
    void thePageBreakSampleBreaksWhereTheManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("21-page-breaks.xlsx")) {
            // POI reports the 0-based index of the last row ABOVE a break, so the first row of the next page is
            // Excel row (value + 2): {3, 6} are the pages that start on rows 5 and 8, as Excel's HPageBreaks says
            assertArrayEquals(new int[]{3, 6}, wb.getSheet("EveryThree").getRowBreaks());
            assertArrayEquals(new int[]{2, 5}, wb.getSheet("ByDept").getRowBreaks());
        }
    }

    @Test
    void theMergeSampleMergesWhereTheManifestSays() throws Exception {
        try (XSSFWorkbook wb = open("22-merge-same-values.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(List.of("A2:A4", "A5:A6", "B2:B3", "B5:B6"), sheet.getMergedRegions().stream()
                    .map(org.apache.poi.ss.util.CellRangeAddress::formatAsString).sorted().toList());
            assertEquals("Sales", sheet.getRow(2).getCell(0).getStringCellValue(), "the covered cell keeps its value");
        }
    }

    @Test
    void theStripeAndRequiredMarkSampleIsAsDescribed() throws Exception {
        try (XSSFWorkbook wb = open("23-striped-rows-required-mark.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("@\" *\"", sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString());
            assertNotEquals("@\" *\"", sheet.getRow(0).getCell(1).getCellStyle().getDataFormatString());
            for (int r : new int[]{2, 4, 6}) assertEquals("none", fill(sheet.getRow(r - 1).getCell(0)), "row " + r);
            for (int r : new int[]{3, 5, 7}) assertEquals("F2F2F2", fill(sheet.getRow(r - 1).getCell(0)), "row " + r);
            assertEquals("F2F2F2", fill(sheet.getRow(2).getCell(1)));
        }
    }

    @Test
    void theErrorReportSampleShadesAndCommentsTheFailedCells() throws Exception {
        try (XSSFWorkbook wb = open("24-error-report.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Errors");
            assertEquals(3, (long) sheet.getRow(1).getCell(0).getNumericCellValue());
            assertEquals(5, (long) sheet.getRow(2).getCell(0).getNumericCellValue());
            assertEquals("age is not a number", sheet.getRow(1).getCell(3).getStringCellValue());
            assertEquals("not a number", sheet.getRow(1).getCell(2).getCellComment().getString().getString());
            assertEquals("duplicate of row 3", sheet.getRow(2).getCell(1).getCellComment().getString().getString());
            assertNull(sheet.getRow(1).getCell(1).getCellComment());
            for (var cell : new String[]{"C2", "B3", "D2", "D3"}) {
                var ref = new org.apache.poi.ss.util.CellReference(cell);
                assertEquals("FFC7CE", fill(sheet.getRow(ref.getRow()).getCell(ref.getCol())), cell);
            }
            assertEquals("none", fill(sheet.getRow(1).getCell(1)));
            assertEquals("none", fill(sheet.getRow(2).getCell(2)));
        }
    }

    @Test
    void theTableSampleKeepsThePlainHeaderNameUnderTheMark() throws Exception {
        try (XSSFWorkbook wb = open("25-required-mark-in-table.xlsx")) {
            XSSFSheet sheet = wb.getSheet("Report");
            var table = sheet.getTables().get(0);
            assertEquals("Marked", table.getName());
            assertEquals("A1:B7", table.getCellReferences().formatAsString());
            assertEquals("Name", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("@\" *\"", sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString());
            for (int r : new int[]{3, 5, 7}) assertEquals("F2F2F2", fill(sheet.getRow(r - 1).getCell(0)), "row " + r);
        }
    }

    @Test
    void theRolloverSampleHasEachFeatureOnEachSheet() throws Exception {
        try (XSSFWorkbook wb = open("26-rollover-combined.xlsx")) {
            assertEquals(List.of("Report", "Report (2)", "Report (3)", "_excelkit_dep_1"),
                    List.of(wb.getSheetName(0), wb.getSheetName(1), wb.getSheetName(2), wb.getSheetName(3)));
            assertEquals(SheetVisibility.HIDDEN, wb.getSheetVisibility(3));
            var merged = new java.util.ArrayList<List<String>>();
            for (int i = 0; i < 3; i++) merged.add(wb.getSheetAt(i).getMergedRegions().stream()
                    .map(org.apache.poi.ss.util.CellRangeAddress::formatAsString).sorted().toList());
            assertEquals(List.of(List.of("A2:A4"), List.of("A2:A3", "A4:A5"), List.of("A2:A3")), merged);
            // the stored value is the last row above the break: pages that start on Excel rows 5 and 4
            assertArrayEquals(new int[]{3}, wb.getSheet("Report").getRowBreaks());
            assertArrayEquals(new int[]{2}, wb.getSheet("Report (2)").getRowBreaks());
            assertArrayEquals(new int[0], wb.getSheet("Report (3)").getRowBreaks());
            assertEquals("A1:C5", filter(wb.getSheet("Report")));
            assertEquals("A1:C5", filter(wb.getSheet("Report (2)")));
            assertEquals("A1:C3", filter(wb.getSheet("Report (3)")));
            for (int i = 0; i < 3; i++) {
                var sheet = wb.getSheetAt(i);
                assertEquals(80, sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale(), "sheet " + i);
                assertEquals(1, sheet.getDataValidations().size(), "sheet " + i);
                assertEquals("F2F2F2", fill(sheet.getRow(2).getCell(2)), "sheet " + i + " C3");
                assertEquals("none", fill(sheet.getRow(1).getCell(2)), "sheet " + i + " C2");
            }
        }
    }
}
