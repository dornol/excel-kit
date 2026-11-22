package io.github.dornol.excelkit.excel;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
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
        assertEquals(16, samples.size());
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
            assertEquals(CellType.FORMULA, row.getCell(7).getCellType());
            assertEquals("E2*2", row.getCell(7).getCellFormula());
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
}
