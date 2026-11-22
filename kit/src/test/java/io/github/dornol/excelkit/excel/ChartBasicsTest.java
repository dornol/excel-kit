package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ChartBasicsTest {
    record Data(String name, int value, int extra) {}

    @Test
    void chart_withNoTitle() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Data>create()
                .column("Name", Data::name)
                .column("Value", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> ch
                        .type(ExcelChartConfig.ChartType.BAR)
                        .categoryColumn(0)
                        .valueColumn(1, "Values"))
                .write(Stream.of(new Data("A", 10, 0)))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_defaultPosition() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Data>create()
                .column("Name", Data::name)
                .column("Value", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> ch
                        .type(ExcelChartConfig.ChartType.LINE)
                        .title("Auto Position")
                        .categoryColumn(0)
                        .valueColumn(1, "Values"))
                .write(Stream.of(new Data("A", 10, 0), new Data("B", 20, 0)))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
        }
    }

    @Test
    void chart_nonZeroCategoryColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Data>create()
                .column("ID", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                .column("Name", Data::name)
                .column("Value", d -> d.extra, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> ch
                        .type(ExcelChartConfig.ChartType.PIE)
                        .categoryColumn(1)  // use Name as category
                        .valueColumn(2, "Values"))
                .write(Stream.of(new Data("A", 1, 60), new Data("B", 2, 40)))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
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
