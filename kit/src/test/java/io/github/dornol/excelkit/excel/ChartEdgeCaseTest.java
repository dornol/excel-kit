package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelChartConfig} to cover uncovered branches:
 * - All chart types (SCATTER, AREA, DOUGHNUT) with actual write
 * - Chart without title
 * - All legend positions
 * - Bar grouping variants (STANDARD, PERCENT_STACKED)
 * - showDataLabels with each chart type
 * - Chart without axis titles
 */
class ChartEdgeCaseTest {

    record Item(String name, int value, double score) {}

    private static Stream<Item> testData() {
        return Stream.of(new Item("A", 100, 1.5), new Item("B", 200, 3.0), new Item("C", 150, 2.5));
    }

    // ============================================================
    // SCATTER chart with actual write
    // ============================================================
    @Test
    void scatterChart_withDataLabels_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("X", i -> i.value, c -> c.type(ExcelDataType.DOUBLE))
                .addColumn("Y", i -> i.score, c -> c.type(ExcelDataType.DOUBLE))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.SCATTER)
                        .categoryColumn(0)
                        .valueColumn(1, "Score")
                        .categoryAxisTitle("Value")
                        .valueAxisTitle("Score")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing);
            assertFalse(drawing.getCharts().isEmpty());
            var plotArea = drawing.getCharts().get(0).getCTChart().getPlotArea();
            assertFalse(plotArea.getScatterChartList().isEmpty(), "Should have scatter chart");
        }
    }

    @Test
    void scatterChart_withoutAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("X", i -> i.value, c -> c.type(ExcelDataType.DOUBLE))
                .addColumn("Y", i -> i.score, c -> c.type(ExcelDataType.DOUBLE))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.SCATTER)
                        .categoryColumn(0)
                        .valueColumn(1, "Score"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    // ============================================================
    // AREA chart with actual write
    // ============================================================
    @Test
    void areaChart_withDataLabels_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.AREA)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .categoryAxisTitle("Category")
                        .valueAxisTitle("Amount")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var plotArea = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea();
            assertFalse(plotArea.getAreaChartList().isEmpty(), "Should have area chart");
        }
    }

    @Test
    void areaChart_withoutAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.AREA)
                        .categoryColumn(0)
                        .valueColumn(1, "Value"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    // ============================================================
    // DOUGHNUT chart with actual write
    // ============================================================
    @Test
    void doughnutChart_withDataLabels_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.DOUGHNUT)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var plotArea = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea();
            assertFalse(plotArea.getDoughnutChartList().isEmpty(), "Should have doughnut chart");
        }
    }

    // ============================================================
    // Chart without title (title == null branch)
    // ============================================================
    @Test
    void chart_withoutTitle_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Value"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            assertNotNull(chart);
        }
    }

    // ============================================================
    // All legend positions
    // ============================================================
    @ParameterizedTest
    @EnumSource(ExcelChartConfig.LegendPosition.class)
    void chart_allLegendPositions(ExcelChartConfig.LegendPosition position) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .legendPosition(position))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            assertNotNull(chart.getCTChart().getLegend());
        }
    }

    // ============================================================
    // Chart without legend (legendPosition == null, default)
    // ============================================================
    @Test
    void chart_withoutLegend_shouldNotHaveLegend() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("No Legend")
                        .categoryColumn(0)
                        .valueColumn(1, "Value"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            assertFalse(chart.getCTChart().isSetLegend());
        }
    }

    // ============================================================
    // Bar grouping variants
    // ============================================================
    @ParameterizedTest
    @EnumSource(ExcelChartConfig.BarGrouping.class)
    void barChart_allGroupings(ExcelChartConfig.BarGrouping grouping) throws IOException {
        record Data(String name, int a, int b) {}

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Data>()
                .addColumn("Name", Data::name)
                .addColumn("A", d -> d.a, c -> c.type(ExcelDataType.INTEGER))
                .addColumn("B", d -> d.b, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "A")
                        .valueColumn(2, "B")
                        .barGrouping(grouping))
                .write(Stream.of(new Data("X", 10, 20), new Data("Y", 30, 40)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    // ============================================================
    // showDataLabels with BAR, LINE, PIE
    // ============================================================
    @Test
    void barChart_showDataLabels() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var barChart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0)
                    .getCTChart().getPlotArea().getBarChartList().get(0);
            assertTrue(barChart.isSetDLbls(), "Data labels should be set");
        }
    }

    @Test
    void lineChart_showDataLabels() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.LINE)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var lineChart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0)
                    .getCTChart().getPlotArea().getLineChartList().get(0);
            assertTrue(lineChart.isSetDLbls());
        }
    }

    @Test
    void pieChart_showDataLabels() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.PIE)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .showDataLabels(true))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var pieChart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0)
                    .getCTChart().getPlotArea().getPieChartList().get(0);
            assertTrue(pieChart.isSetDLbls());
        }
    }

    // ============================================================
    // Chart without axis titles (null branch in applyAxisTitles)
    // ============================================================
    @Test
    void barChart_withoutAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("No Axis Titles")
                        .categoryColumn(0)
                        .valueColumn(1, "Value"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            var catAxis = chart.getCTChart().getPlotArea().getCatAxList();
            assertFalse(catAxis.get(0).isSetTitle(), "Category axis should not have a title");
        }
    }

    @Test
    void lineChart_withoutAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.LINE)
                        .categoryColumn(0)
                        .valueColumn(1, "Value"))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    // ============================================================
    // showDataLabels=false (default, but explicit)
    // ============================================================
    @Test
    void chart_showDataLabelsFalse_noDataLabels() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Value")
                        .showDataLabels(false))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var barChart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0)
                    .getCTChart().getPlotArea().getBarChartList().get(0);
            assertFalse(barChart.isSetDLbls(), "Data labels should not be set");
        }
    }
}
