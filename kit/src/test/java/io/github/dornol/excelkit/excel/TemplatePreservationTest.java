package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Base64;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A template is somebody's layout: whatever it holds beyond cell values (merged cells, pictures, charts, a table,
 * conditional formatting, validations, names, widths, panes, headers, ...) must come out of the writer as it went in, on
 * every sheet, while the writer fills in its cells and a list below them.
 */
class TemplatePreservationTest {

    /** A 1x1 PNG. */
    private static final byte[] PNG = Base64.getDecoder().decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private static byte[] template;
    private static byte[] output;

    private record Item(String name, int qty) {
    }

    @BeforeAll
    static void build() throws Exception {
        template = buildTemplate();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B2", "A");                           // the cell that has a validation
            writer.cell("A9", "Filled by the writer");        // below the table, above the list
            writer.<Item>list(11).column("name", Item::name).column("qty", Item::qty)
                    .write(List.of(new Item("x", 1), new Item("y", 2)));
            writer.sheet("Notes");
            writer.cell("A2", "second sheet cell");
            writer.finish().writeTo(out);
        }
        output = out.toByteArray();
    }

    private static byte[] buildTemplate() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.createSheet("Report");

            XSSFCellStyle titleStyle = wb.createCellStyle();
            XSSFFont bold = wb.createFont();
            bold.setBold(true);
            titleStyle.setFont(bold);
            titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            var title = sheet.createRow(0).createCell(0);
            title.setCellValue("Quarterly report");
            title.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));

            // a table with its data, rows 4-7 (0-based 3..6)
            var header = sheet.createRow(3);
            header.createCell(0).setCellValue("Region");
            header.createCell(1).setCellValue("Sales");
            for (int i = 0; i < 3; i++) {
                var row = sheet.createRow(4 + i);
                row.createCell(0).setCellValue("R" + i);
                row.createCell(1).setCellValue(10 * (i + 1));
            }
            var table = sheet.createTable(new AreaReference("A4:B7", SpreadsheetVersion.EXCEL2007));
            table.setName("SalesTable");
            table.setDisplayName("SalesTable");
            table.updateHeaders();

            // a picture and a chart
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 4, 0, 6, 3), wb.addPicture(PNG, XSSFWorkbook.PICTURE_TYPE_PNG));
            XSSFChart chart = drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 8, 0, 14, 12));
            XDDFCategoryAxis bottom = chart.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis left = chart.createValueAxis(AxisPosition.LEFT);
            XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(4, 6, 0, 0));
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(4, 6, 1, 1));
            XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottom, left);
            // a series without a title is written as an empty <c:tx/>, which Excel refuses
            data.addSeries(categories, values).setTitle("Sales", new org.apache.poi.ss.util.CellReference("Report", 3, 1, true, true));
            chart.plot(data);

            // conditional formatting and a validation
            SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();
            ConditionalFormattingRule rule = formatting.createConditionalFormattingRule(ComparisonOperator.GT, "15");
            formatting.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("B5:B7")}, rule);
            DataValidationHelper helper = sheet.getDataValidationHelper();
            DataValidation validation = helper.createValidation(helper.createExplicitListConstraint(new String[]{"A", "B"}),
                    new CellRangeAddressList(1, 1, 1, 1));
            sheet.addValidationData(validation);

            // names, widths, panes, headers, colours, print setup
            var name = wb.createName();
            name.setNameName("ReportTitle");
            name.setRefersToFormula("Report!$A$1");
            sheet.setColumnWidth(0, 30 * 256);
            sheet.setColumnHidden(5, true);
            sheet.createFreezePane(0, 1);
            sheet.getHeader().setCenter("Template header");
            sheet.getFooter().setRight("Page &P");
            sheet.setTabColor(new XSSFColor(new byte[]{(byte) 255, 0, 0}, null));
            sheet.getPrintSetup().setOrientation(PrintOrientation.LANDSCAPE);
            sheet.getPrintSetup().setScale((short) 80);

            // another sheet that nobody writes to
            var notes = wb.createSheet("Notes");
            notes.createRow(0).createCell(0).setCellValue("kept note");
            notes.setColumnWidth(0, 20 * 256);

            wb.getProperties().getCoreProperties().setTitle("Template title");
            wb.getProperties().getCoreProperties().setCreator("Template author");
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static XSSFWorkbook open() throws Exception {
        return new XSSFWorkbook(new ByteArrayInputStream(output));
    }

    @Test
    void theMergedTitleAndItsStyleSurvive() throws Exception {
        try (XSSFWorkbook wb = open()) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(1, sheet.getNumMergedRegions());
            assertEquals("A1:C1", sheet.getMergedRegion(0).formatAsString());
            var title = sheet.getRow(0).getCell(0);
            assertEquals("Quarterly report", title.getStringCellValue());
            assertTrue(title.getCellStyle().getFont().getBold(), "the title is still bold");
            assertEquals(FillPatternType.SOLID_FOREGROUND, title.getCellStyle().getFillPattern());
        }
    }

    @Test
    void thePictureSurvivesWithItsBytesAndPlace() throws Exception {
        try (XSSFWorkbook wb = open()) {
            assertEquals(1, wb.getAllPictures().size(), "pictures in the workbook");
            assertArrayEquals(PNG, wb.getAllPictures().get(0).getData());
            var shapes = wb.getSheet("Report").getDrawingPatriarch().getShapes();
            var picture = shapes.stream().filter(s -> s instanceof org.apache.poi.xssf.usermodel.XSSFPicture)
                    .map(s -> (org.apache.poi.xssf.usermodel.XSSFPicture) s).findFirst().orElse(null);
            assertNotNull(picture, "a picture shape on the Report sheet");
            assertEquals(4, picture.getClientAnchor().getCol1());
            assertEquals(6, picture.getClientAnchor().getCol2());
        }
    }

    @Test
    void theChartSurvivesWithItsData() throws Exception {
        try (XSSFWorkbook wb = open()) {
            var charts = wb.getSheet("Report").getDrawingPatriarch().getCharts();
            assertEquals(1, charts.size());
            var series = charts.get(0).getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0);
            assertEquals("Report!$B$5:$B$7", series.getVal().getNumRef().getF());
            assertEquals("Report!$A$5:$A$7", series.getCat().getStrRef().getF());
        }
    }

    @Test
    void theTableSurvivesWithItsNameAndRange() throws Exception {
        try (XSSFWorkbook wb = open()) {
            var tables = wb.getSheet("Report").getTables();
            assertEquals(1, tables.size());
            assertEquals("SalesTable", tables.get(0).getName());
            assertEquals("A4:B7", tables.get(0).getCellReferences().formatAsString());
        }
    }

    @Test
    void conditionalFormattingAndValidationSurvive() throws Exception {
        try (XSSFWorkbook wb = open()) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals(1, sheet.getSheetConditionalFormatting().getNumConditionalFormattings());
            var formatting = sheet.getSheetConditionalFormatting().getConditionalFormattingAt(0);
            assertEquals("B5:B7", formatting.getFormattingRanges()[0].formatAsString());
            assertEquals(ComparisonOperator.GT, formatting.getRule(0).getComparisonOperation());

            assertEquals(1, sheet.getDataValidations().size());
            var validation = sheet.getDataValidations().get(0);
            assertEquals("B2", validation.getRegions().getCellRangeAddress(0).formatAsString());
            assertArrayEquals(new String[]{"A", "B"}, validation.getValidationConstraint().getExplicitListValues());
        }
    }

    @Test
    void namesWidthsPanesHeadersAndColoursSurvive() throws Exception {
        try (XSSFWorkbook wb = open()) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertNotNull(wb.getName("ReportTitle"));
            assertEquals("Report!$A$1", wb.getName("ReportTitle").getRefersToFormula());
            assertEquals(30 * 256, sheet.getColumnWidth(0));
            assertTrue(sheet.isColumnHidden(5), "the hidden column");
            assertNotNull(sheet.getPaneInformation());
            assertTrue(sheet.getPaneInformation().isFreezePane());
            assertEquals(1, sheet.getPaneInformation().getHorizontalSplitPosition());
            assertEquals("Template header", sheet.getHeader().getCenter());
            assertEquals("Page &P", sheet.getFooter().getRight());
            assertNotNull(sheet.getTabColor());
            assertEquals(PrintOrientation.LANDSCAPE, sheet.getPrintSetup().getOrientation());
            assertEquals(80, sheet.getPrintSetup().getScale());
        }
    }

    @Test
    void aSheetNobodyWroteToKeepsItsContentAndWidth() throws Exception {
        try (XSSFWorkbook wb = open()) {
            XSSFSheet notes = wb.getSheet("Notes");
            assertEquals("kept note", notes.getRow(0).getCell(0).getStringCellValue());
            assertEquals(20 * 256, notes.getColumnWidth(0));
            assertEquals("second sheet cell", notes.getRow(1).getCell(0).getStringCellValue(), "and what was written to it");
        }
    }

    @Test
    void theDocumentPropertiesSurvive() throws Exception {
        try (XSSFWorkbook wb = open()) {
            assertEquals("Template title", wb.getProperties().getCoreProperties().getTitle());
            assertEquals("Template author", wb.getProperties().getCoreProperties().getCreator());
        }
    }

    @Test
    void theTemplatesOwnCellsKeepTheirValuesWhileTheWriterFillsItsOwn() throws Exception {
        try (XSSFWorkbook wb = open()) {
            XSSFSheet sheet = wb.getSheet("Report");
            assertEquals("Region", sheet.getRow(3).getCell(0).getStringCellValue());
            assertEquals(30.0, sheet.getRow(6).getCell(1).getNumericCellValue());
            assertEquals("A", sheet.getRow(1).getCell(1).getStringCellValue(), "the cell with the validation");
            assertEquals("Filled by the writer", sheet.getRow(8).getCell(0).getStringCellValue());
            assertEquals("x", sheet.getRow(11).getCell(0).getStringCellValue(), "the list wrote its first row at index 11");
        }
    }
}
