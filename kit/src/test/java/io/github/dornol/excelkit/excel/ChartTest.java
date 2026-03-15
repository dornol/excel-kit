package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
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
