package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class AdditionalChartTest {

    record Product(String name, int sales, double price) {}

    @Test
    void scatterChart_createsSuccessfully() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("A", 100, 10.5),
                new Product("B", 200, 20.3),
                new Product("C", 150, 15.7)
        );

        ExcelWriter.<Product>create()
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.SCATTER)
                        .title("Sales vs Price")
                        .categoryColumn(0)
                        .valueColumn(1, "Price")
                        .showDataLabels(true))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing, "Drawing should exist for scatter chart");
            assertFalse(drawing.getCharts().isEmpty(), "Should have at least one chart");
            // Verify it is a scatter chart
            var scatterList = drawing.getCharts().get(0).getCTChart()
                    .getPlotArea().getScatterChartList();
            assertFalse(scatterList.isEmpty(), "Should have a scatter chart");
        }
    }

    @Test
    void areaChart_createsSuccessfully() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("Q1", 100, 10.5),
                new Product("Q2", 200, 20.3),
                new Product("Q3", 150, 15.7)
        );

        ExcelWriter.<Product>create()
                .column("Quarter", Product::name)
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.AREA)
                        .title("Sales Trend")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing, "Drawing should exist for area chart");
            assertFalse(drawing.getCharts().isEmpty());
            var areaList = drawing.getCharts().get(0).getCTChart()
                    .getPlotArea().getAreaChartList();
            assertFalse(areaList.isEmpty(), "Should have an area chart");
        }
    }

    @Test
    void doughnutChart_createsSuccessfully() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("Phone", 300, 0),
                new Product("Laptop", 200, 0),
                new Product("Tablet", 100, 0)
        );

        ExcelWriter.<Product>create()
                .column("Category", Product::name)
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.DOUGHNUT)
                        .title("Sales Distribution")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .showDataLabels(true))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
            assertNotNull(drawing, "Drawing should exist for doughnut chart");
            assertFalse(drawing.getCharts().isEmpty());
            var doughnutList = drawing.getCharts().get(0).getCTChart()
                    .getPlotArea().getDoughnutChartList();
            assertFalse(doughnutList.isEmpty(), "Should have a doughnut chart");
        }
    }

    @Test
    void areaChart_withAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("Q1", 100, 10.5),
                new Product("Q2", 200, 20.3)
        );

        ExcelWriter.<Product>create()
                .column("Quarter", Product::name)
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.AREA)
                        .title("Sales Trend")
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .categoryAxisTitle("Quarter")
                        .valueAxisTitle("Sales Amount"))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            var catAxis = chart.getCTChart().getPlotArea().getCatAxList();
            assertFalse(catAxis.isEmpty(), "Chart should have a category axis");
            assertTrue(catAxis.get(0).isSetTitle(), "Category axis should have a title");
        }
    }

    @Test
    void scatterChart_withAxisTitles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("A", 100, 10.5),
                new Product("B", 200, 20.3)
        );

        ExcelWriter.<Product>create()
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.SCATTER)
                        .categoryColumn(0)
                        .valueColumn(1, "Price")
                        .categoryAxisTitle("Sales")
                        .valueAxisTitle("Price"))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            // Scatter uses value axes (not category axes)
            var valAxes = chart.getCTChart().getPlotArea().getValAxList();
            assertTrue(valAxes.size() >= 2, "Scatter chart should have 2 value axes");
        }
    }

    @Test
    void scatterChart_multiSeries() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("A", 100, 10.5),
                new Product("B", 200, 20.3)
        );

        ExcelWriter.<Product>create()
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                .column("Volume", p -> p.sales() * 2, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.SCATTER)
                        .title("Multi-Series Scatter")
                        .categoryColumn(0)
                        .valueColumn(1, "Price")
                        .valueColumn(2, "Volume"))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            var scatterList = chart.getCTChart().getPlotArea().getScatterChartList();
            assertFalse(scatterList.isEmpty());
            // Should have 2 series
            assertEquals(2, scatterList.get(0).getSerList().size(), "Should have 2 scatter series");
        }
    }

    @Test
    void areaChart_multipleSeries() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        List<Product> data = List.of(
                new Product("Q1", 100, 10.5),
                new Product("Q2", 200, 20.3)
        );

        ExcelWriter.<Product>create()
                .column("Quarter", Product::name)
                .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                .chart(chart -> chart
                        .type(ExcelChartConfig.ChartType.AREA)
                        .categoryColumn(0)
                        .valueColumn(1, "Sales")
                        .valueColumn(2, "Price"))
                .write(data.stream())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var areaList = wb.getSheetAt(0).getDrawingPatriarch().getCharts()
                    .get(0).getCTChart().getPlotArea().getAreaChartList();
            assertEquals(2, areaList.get(0).getSerList().size(), "Should have 2 area series");
        }
    }

    @Test
    void newChartTypes_inExcelWorkbook() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Product>sheet("Scatter")
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .categoryColumn(0)
                            .valueColumn(1, "Price"))
                    .write(Stream.of(new Product("A", 100, 10.5)));

            wb.<Product>sheet("Area")
                    .column("Name", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .categoryColumn(0)
                            .valueColumn(1, "Sales"))
                    .write(Stream.of(new Product("Q1", 100, 0)));

            wb.<Product>sheet("Doughnut")
                    .column("Name", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.DOUGHNUT)
                            .categoryColumn(0)
                            .valueColumn(1, "Sales"))
                    .write(Stream.of(new Product("A", 60, 0)));

            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, xwb.getNumberOfSheets());
            for (int i = 0; i < 3; i++) {
                assertNotNull(xwb.getSheetAt(i).getDrawingPatriarch(),
                        "Sheet " + i + " should have a chart");
            }
        }
    }

    @Test
    void chartType_enum_coverage() {
        assertEquals(6, ExcelChartConfig.ChartType.values().length);
        assertNotNull(ExcelChartConfig.ChartType.valueOf("BAR"));
        assertNotNull(ExcelChartConfig.ChartType.valueOf("LINE"));
        assertNotNull(ExcelChartConfig.ChartType.valueOf("PIE"));
        assertNotNull(ExcelChartConfig.ChartType.valueOf("SCATTER"));
        assertNotNull(ExcelChartConfig.ChartType.valueOf("AREA"));
        assertNotNull(ExcelChartConfig.ChartType.valueOf("DOUGHNUT"));
    }
}
