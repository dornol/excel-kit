package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import static org.junit.jupiter.api.Assertions.*;

/** Also runs against all POI runtimes with the same compiled library JAR. */
class PublicSheetFeatureCompatibilityTest {
    record Entry(String name, int value) { }

    @Test
    void streamingWorkbookPersistsChartsLabelsTabColorConditionalRulesAndTable() throws Exception {
        var output = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create(options -> options.rowAccessWindowSize(1))) {
            workbook.<String>sheet("Cover").column("Info", value -> value).write(Stream.of("unchanged"));
            workbook.<Entry>sheet("Data")
                .column("Name", Entry::name)
                .column("Value", Entry::value, column -> column.type(ExcelDataType.INTEGER))
                .tabColor(12, 34, 56)
                .table("DataTable")
                .conditionalFormatting(rule -> rule.columns(1).dataBar(ExcelColor.BLUE))
                .conditionalFormatting(rule -> rule.columns(1).iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
                .chart(chart -> chart.categoryColumn(0).valueColumn(1, "Value").showDataLabels(true))
                .write(IntStream.range(0, 20).mapToObj(i -> new Entry("Item " + i, i)));
            workbook.finish().writeTo(output);
        }
        try (var workbook = reopen(output)) {
            assertEquals("unchanged", workbook.getSheet("Cover").getRow(1).getCell(0).getStringCellValue());
            assertNull(workbook.getSheet("Cover").getTabColor());
            assertTrue(workbook.getSheet("Cover").getTables().isEmpty());
            var sheet = workbook.getSheet("Data");
            assertEquals(20, sheet.getLastRowNum());
            assertArrayEquals(new byte[]{12, 34, 56}, sheet.getTabColor().getRGB());
            assertConditionalRules(sheet);
            assertEquals(1, sheet.getTables().size());
            assertEquals("DataTable", sheet.getTables().get(0).getName());
            assertEquals("A1:B21", sheet.getTables().get(0).getCTTable().getRef());
            assertEquals("Name", sheet.getTables().get(0).getCTTable().getTableColumns().getTableColumnArray(0).getName());
            var charts = sheet.getDrawingPatriarch().getCharts();
            assertEquals(1, charts.size());
            var bar = charts.get(0).getCTChart().getPlotArea().getBarChartArray(0);
            assertTrue(bar.getDLbls().getShowVal().getVal());
            assertTrue(bar.getSerArray(0).getVal().getNumRef().getF().contains("Data"));
        }
    }

    @Test
    void rolloverSheetsKeepTheirOwnTablesAndConditionalFormatting() throws Exception {
        var output = new ByteArrayOutputStream();
        ExcelWriter.<Entry>create(options -> options.rowAccessWindowSize(1))
            .column("Name", Entry::name)
            .column("Value", Entry::value, column -> column.type(ExcelDataType.INTEGER))
            .maxRows(3).table("RowsTable").tabColor(12, 34, 56)
            .conditionalFormatting(rule -> rule.columns(1).dataBar(ExcelColor.BLUE))
            .conditionalFormatting(rule -> rule.columns(1).iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
            .write(IntStream.range(0, 8).mapToObj(i -> new Entry("Item " + i, i))).writeTo(output);
        try (var workbook = reopen(output)) {
            assertTrue(workbook.getNumberOfSheets() > 1);
            int dataRows = 0;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                var sheet = workbook.getSheetAt(i);
                dataRows += sheet.getLastRowNum();
                assertArrayEquals(new byte[]{12, 34, 56}, sheet.getTabColor().getRGB());
                assertEquals("RowsTable_" + (i + 1), sheet.getTables().get(0).getName());
                assertEquals("Name", sheet.getTables().get(0).getCTTable().getTableColumns().getTableColumnArray(0).getName());
                assertEquals("A1:B" + (sheet.getLastRowNum() + 1), sheet.getTables().get(0).getCTTable().getRef());
                assertConditionalRules(sheet);
            }
            assertEquals(8, dataRows);
        }
    }

    @Test
    void templateTableUsesExistingHeaderAndPreservesOtherSheets() throws Exception {
        var template = new ByteArrayOutputStream();
        try (var workbook = new XSSFWorkbook()) {
            workbook.createSheet("Cover").createRow(0).createCell(0).setCellValue("preserved");
            var row = workbook.createSheet("Data").createRow(0);
            row.createCell(0).setCellValue("Name");
            row.createCell(1).setCellValue("Value");
            workbook.write(template);
        }
        var output = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template.toByteArray()), 1)) {
            writer.<Entry>list(1, 1).column("Mapped Name", Entry::name).column("Mapped Value", Entry::value)
                .table("TemplateData").write(Stream.of(new Entry("Alice", 10), new Entry("Bob", 20)));
            writer.finish().writeTo(output);
        }
        try (var workbook = reopen(output)) {
            assertEquals("preserved", workbook.getSheet("Cover").getRow(0).getCell(0).getStringCellValue());
            var sheet = workbook.getSheet("Data");
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("TemplateData", sheet.getTables().get(0).getName());
            assertEquals("Name", sheet.getTables().get(0).getCTTable().getTableColumns().getTableColumnArray(0).getName());
            assertEquals("A1:B3", sheet.getTables().get(0).getCTTable().getRef());
        }
    }

    @Test
    void templateWithGeneratedHeadersKeepsTableColumnNamesAfterFlush() throws Exception {
        var template = new ByteArrayOutputStream();
        try (var workbook = new XSSFWorkbook()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Title");
            workbook.write(template);
        }
        var output = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template.toByteArray()), 1)) {
            writer.<Entry>list(1).column("Name", Entry::name).column("Value", Entry::value)
                .table("GeneratedHeaders").writeWithHeaders(Stream.of(new Entry("Alice", 10), new Entry("Bob", 20)));
            writer.finish().writeTo(output);
        }
        try (var workbook = reopen(output)) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("Title", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Name", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("A2:B4", sheet.getTables().get(0).getCTTable().getRef());
            assertEquals("Name", sheet.getTables().get(0).getCTTable().getTableColumns().getTableColumnArray(0).getName());
        }
    }

    private static void assertConditionalRules(XSSFSheet sheet) {
        boolean dataBar = false;
        boolean iconSet = false;
        for (var formatting : sheet.getCTWorksheet().getConditionalFormattingList()) {
            for (var rule : formatting.getCfRuleList()) {
                dataBar |= rule.isSetDataBar();
                iconSet |= rule.isSetIconSet();
            }
        }
        assertTrue(dataBar, "Data bar must be serialized");
        assertTrue(iconSet, "Icon set must be serialized");
    }

    private static XSSFWorkbook reopen(ByteArrayOutputStream output) throws Exception {
        return new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
    }
}
