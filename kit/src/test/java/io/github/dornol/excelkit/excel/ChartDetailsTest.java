package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.ExcelChartConfig.ChartType;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** What a chart shows, read back from the chart XML: its series, its data labels, its titles. */
class ChartDetailsTest {

    record Row(String name, int value) {
    }

    private static String chartXml(Consumer<ExcelChartConfig> configure) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().column("Name", Row::name).column("Value", Row::value)
                .chart(configure)
                .write(List.of(new Row("a", 1), new Row("b", 2), new Row("c", 3))).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            List<XSSFChart> charts = workbook.getSheetAt(0).getDrawingPatriarch().getCharts();
            assertEquals(1, charts.size());
            return charts.get(0).getCTChart().xmlText();
        }
    }

    private static Consumer<ExcelChartConfig> chart(ChartType type) {
        return c -> c.type(type).categoryColumn(0).valueColumn(1, "Series");
    }

    /** The value of an on/off element such as {@code <c:showVal val="1"/>}, whether XmlBeans prints 1 or true. */
    private static boolean flag(String xml, String element) {
        Matcher matcher = Pattern.compile("<c:" + element + " val=\"(\\w+)\"/>").matcher(xml);
        assertTrue(matcher.find(), element + " is set in " + xml);
        return matcher.group(1).equals("1") || matcher.group(1).equals("true");
    }

    @Test
    void everyChartTypeHasItsSeries() throws Exception {
        for (ChartType type : ChartType.values()) {
            String xml = chartXml(chart(type));
            assertEquals(1, count(xml, "<c:ser>"), type + " has one series: " + xml);
            assertTrue(xml.contains("Series"), type + " names its series");
        }
    }

    @Test
    void dataLabelsShowTheValueOnlyWhenAsked() throws Exception {
        for (ChartType type : ChartType.values()) {
            String labelled = chartXml(chart(type).andThen(c -> c.showDataLabels(true)));
            assertTrue(flag(labelled, "showVal"), type + " shows the value");
            assertFalse(flag(labelled, "showCatName"), type + " does not show the category");
            assertFalse(flag(labelled, "showSerName"), type + " does not show the series name");

            String plain = chartXml(chart(type));
            assertFalse(plain.contains("<c:dLbls>") && flag(plain, "showVal"), type + " has no value labels by default");
        }
    }

    @Test
    void axisTitlesAndChartTitleAreWritten() throws Exception {
        for (ChartType type : List.of(ChartType.BAR, ChartType.LINE, ChartType.AREA, ChartType.SCATTER)) {
            String xml = chartXml(chart(type).andThen(c -> c.title("Sales").categoryAxisTitle("Month").valueAxisTitle("Amount")));
            assertTrue(xml.contains("Sales"), type + " title");
            assertTrue(xml.contains("Month"), type + " category axis title");
            assertTrue(xml.contains("Amount"), type + " value axis title");
        }
    }

    @Test
    void theLegendIsPlacedWhereAsked() throws Exception {
        for (ExcelChartConfig.LegendPosition position : ExcelChartConfig.LegendPosition.values()) {
            String xml = chartXml(chart(ChartType.BAR).andThen(c -> c.legendPosition(position)));
            String code = switch (position) {
                case BOTTOM -> "b";
                case LEFT -> "l";
                case RIGHT -> "r";
                case TOP -> "t";
                default -> null;
            };
            if (code != null) assertTrue(xml.contains("<c:legendPos val=\"" + code + "\"/>"), position + ": " + xml);
        }
    }

    private static int count(String text, String part) {
        int count = 0;
        for (int at = text.indexOf(part); at >= 0; at = text.indexOf(part, at + part.length())) count++;
        return count;
    }
}
