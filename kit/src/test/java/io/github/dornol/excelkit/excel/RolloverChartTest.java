package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * With {@code maxRows()} rollover a chart is drawn on every sheet over that sheet's own rows, as
 * {@link ExcelSheetWriter} does. A chart on the last sheet alone would plot only the remainder and leave the
 * other sheets without one.
 */
class RolloverChartTest {

    private record Item(String name, int qty) {
    }

    private static XSSFWorkbook write(boolean summary) throws Exception {
        ExcelWriter<Item> writer = ExcelWriter.<Item>create().maxRows(4)
                .column("Name", Item::name)
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"));
        if (summary) writer.summary(s -> s.label("Total").sum("Qty"));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(IntStream.range(0, 10).mapToObj(i -> new Item("item " + i, i)).toList()).writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    @Test
    void everySheetGetsAChartOverItsOwnRows() throws Exception {
        try (XSSFWorkbook wb = write(false)) {
            assertEquals(3, wb.getNumberOfSheets());
            // 10 rows with maxRows(4): data rows 2-5, 2-5 and 2-3
            List<String> expected = List.of("$B$2:$B$5", "$B$2:$B$5", "$B$2:$B$3");
            for (int s = 0; s < 3; s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                assertNotNull(sheet.getDrawingPatriarch(), "sheet " + s + " has no drawing");
                var charts = sheet.getDrawingPatriarch().getCharts();
                assertEquals(1, charts.size(), "charts on sheet " + s);
                String values = charts.get(0).getCTChart().getPlotArea().getBarChartArray(0)
                        .getSerArray(0).getVal().getNumRef().getF();
                assertTrue(values.endsWith(expected.get(s)), "sheet " + s + " plots " + values);
            }
        }
    }

    @Test
    void aChartOnARolledOverSheetSitsBelowItsSummaryRow() throws Exception {
        try (XSSFWorkbook wb = write(true)) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                int summaryRow = -1;
                for (var row : sheet) {
                    var first = row.getCell(0);
                    if (first != null && "Total".equals(first.getStringCellValue())) summaryRow = row.getRowNum();
                }
                assertTrue(summaryRow > 0, "sheet " + s + " has no summary row");
                int chartTop = sheet.getDrawingPatriarch().getCTDrawing().getTwoCellAnchorArray(0).getFrom().getRow();
                assertTrue(chartTop > summaryRow, "sheet " + s + ": the chart starts at row " + chartTop
                        + " and covers the summary at row " + summaryRow);
            }
        }
    }

    @Test
    void theMultiSheetWriterKeepsEveryRolledOverChartBelowItsSummaryRowToo() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Data").maxRows(4)
                    .column("Name", Item::name)
                    .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Qty"))
                    .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"))
                    .write(IntStream.range(0, 10).mapToObj(i -> new Item("item " + i, i)));
            book.finish().writeTo(out);
        }
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getNumberOfSheets());
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                int summaryRow = -1;
                for (var row : sheet) {
                    var first = row.getCell(0);
                    if (first != null && "Total".equals(first.getStringCellValue())) summaryRow = row.getRowNum();
                }
                assertTrue(summaryRow > 0, "sheet " + s + " has no summary row");
                int chartTop = sheet.getDrawingPatriarch().getCTDrawing().getTwoCellAnchorArray(0).getFrom().getRow();
                assertTrue(chartTop > summaryRow, "sheet " + s + ": the chart starts at row " + chartTop
                        + " and covers the summary at row " + summaryRow);
            }
        }
    }
}
