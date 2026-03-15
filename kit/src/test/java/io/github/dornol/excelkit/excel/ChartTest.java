package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ChartTest {

    record Product(String name, int sales) {}

    @Test
    void chart_barChart_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Sales Chart")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales"))
                .write(Stream.of(new Product("A", 100), new Product("B", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing, "Sheet should have a drawing");
            assertFalse(drawing.getCharts().isEmpty(), "Drawing should contain charts");
        }
    }

    @Test
    void chart_lineChart_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.LINE)
                        .title("Sales Trend")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales"))
                .write(Stream.of(new Product("Q1", 100), new Product("Q2", 150)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_pieChart_shouldBeCreated() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.PIE)
                        .title("Market Share")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales"))
                .write(Stream.of(new Product("A", 60), new Product("B", 40)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_withCustomPosition() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .position(3, 5, 12, 20))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_multipleSeries() throws IOException {
        record Data(String name, int sales, int profit) {}

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Data>()
                .addColumn("Name", Data::name)
                .addColumn("Sales", d -> d.sales, c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Profit", d -> d.profit, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Sales vs Profit")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .valueColumn(2, "Profit"))
                .write(Stream.of(new Data("A", 100, 30), new Data("B", 200, 50)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Product>sheet("Products")
                    .column("Name", Product::name)
                    .column("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.BAR)
                            .categoryColumn(0)
                            .valueColumn(1, "Sales"))
                    .write(Stream.of(new Product("A", 100)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_barChart_withAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Sales Chart")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .categoryAxisTitle("Product")
                        .valueAxisTitle("Amount"))
                .write(Stream.of(new Product("A", 100), new Product("B", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing, "Sheet should have a drawing");
            List<XSSFChart> charts = drawing.getCharts();
            assertFalse(charts.isEmpty(), "Drawing should contain charts");

            XSSFChart chart = charts.get(0);
            // Verify category axis title "Product"
            var catAxis = chart.getCTChart().getPlotArea().getCatAxList();
            assertFalse(catAxis.isEmpty(), "Chart should have a category axis");
            assertTrue(catAxis.get(0).isSetTitle(), "Category axis should have a title");

            // Verify value axis title "Amount"
            var valAxis = chart.getCTChart().getPlotArea().getValAxList();
            assertFalse(valAxis.isEmpty(), "Chart should have a value axis");
            assertTrue(valAxis.get(0).isSetTitle(), "Value axis should have a title");
        }
    }

    @Test
    void chart_lineChart_withAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.LINE)
                        .title("Sales Trend")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .categoryAxisTitle("Quarter")
                        .valueAxisTitle("Revenue"))
                .write(Stream.of(new Product("Q1", 100), new Product("Q2", 150)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);

            // Verify category axis title "Quarter"
            var catAxis = chart.getCTChart().getPlotArea().getCatAxList();
            assertFalse(catAxis.isEmpty(), "Chart should have a category axis");
            assertTrue(catAxis.get(0).isSetTitle(), "Category axis should have a title");

            // Verify value axis title "Revenue"
            var valAxis = chart.getCTChart().getPlotArea().getValAxList();
            assertFalse(valAxis.isEmpty(), "Chart should have a value axis");
            assertTrue(valAxis.get(0).isSetTitle(), "Value axis should have a title");
        }
    }

    @Test
    void chart_withLegendPosition() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Sales Chart")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM))
                .write(Stream.of(new Product("A", 100), new Product("B", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            assertNotNull(chart.getOrAddLegend(), "Chart should have a legend");
            // Verify legend position is BOTTOM
            var ctLegend = chart.getCTChart().getLegend();
            assertNotNull(ctLegend, "CT legend should exist");
            assertEquals(
                    org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos.B,
                    ctLegend.getLegendPos().getVal(),
                    "Legend position should be BOTTOM"
            );
        }
    }

    @Test
    void chart_barChart_horizontalDirection() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Sales", p -> p.sales, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Horizontal Bar")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .barDirection(ExcelChartConfig.BarDirection.HORIZONTAL))
                .write(Stream.of(new Product("A", 100), new Product("B", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            var barChartList = chart.getCTChart().getPlotArea().getBarChartList();
            assertFalse(barChartList.isEmpty(), "Chart should have a bar chart");
            assertEquals(
                    org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir.BAR,
                    barChartList.get(0).getBarDir().getVal(),
                    "Bar direction should be BAR (horizontal)"
            );
        }
    }

    @Test
    void chart_barChart_stackedGrouping() throws IOException {
        record Data(String name, int sales, int profit) {}

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Data>()
                .addColumn("Name", Data::name)
                .addColumn("Sales", d -> d.sales, c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Profit", d -> d.profit, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .title("Stacked Chart")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .valueColumn(2, "Profit")
                        .barGrouping(ExcelChartConfig.BarGrouping.STACKED))
                .write(Stream.of(new Data("A", 100, 30), new Data("B", 200, 50)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            var barChartList = chart.getCTChart().getPlotArea().getBarChartList();
            assertFalse(barChartList.isEmpty(), "Chart should have a bar chart");
            assertEquals(
                    org.openxmlformats.schemas.drawingml.x2006.chart.STBarGrouping.STACKED,
                    barChartList.get(0).getGrouping().getVal(),
                    "Bar grouping should be STACKED"
            );
        }
    }

    @Test
    void chart_noValueSeries_shouldNotCreateChart() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNull(wb.getSheetAt(0).getDrawingPatriarch());
        }
    }
}
