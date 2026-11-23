package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.ExcelChartConfig.BarDirection;
import io.github.dornol.excelkit.excel.ExcelChartConfig.ChartType;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STCrosses;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** A line over bars, on the same value axis or on a second one on the right: the chart XML Excel reads. */
class ChartComboTest {
    record Month(String name, int revenue, int target, double growth) {}

    private static final List<Month> MONTHS = List.of(
            new Month("Jan", 10, 12, 0.1), new Month("Feb", 20, 18, 0.5), new Month("Mar", 30, 33, 0.25));

    private static ExcelWriter<Month> writer(Consumer<ExcelChartConfig> chart) {
        return ExcelWriter.<Month>create().sheetName("Sales")
                .column("Month", Month::name)
                .column("Revenue", Month::revenue, c -> c.type(ExcelDataType.INTEGER))
                .column("Target", Month::target, c -> c.type(ExcelDataType.INTEGER))
                .column("Growth", Month::growth, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .chart(chart);
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static CTPlotArea plot(XSSFWorkbook wb, int sheet) {
        List<XSSFChart> charts = wb.getSheetAt(sheet).getDrawingPatriarch().getCharts();
        assertEquals(1, charts.size(), "one chart on sheet " + sheet);
        return charts.get(0).getCTChart().getPlotArea();
    }

    private static List<Long> lineAxisIds(CTLineChart chart) {
        List<Long> ids = new ArrayList<>();
        for (var ax : chart.getAxIdList()) ids.add(ax.getVal());
        return ids;
    }

    private static List<Long> axisIds(org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart chart) {
        List<Long> ids = new ArrayList<>();
        for (var ax : chart.getAxIdList()) ids.add(ax.getVal());
        return ids;
    }

    private static String values(CTLineChart chart, int series) {
        return chart.getSerArray(series).getVal().getNumRef().getF();
    }

    @Nested
    class OnThePrimaryAxis {
        @Test
        void aLineOverBarsSharesTheAxesOfTheBars() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").lineColumn(2, "Target"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertEquals(1, plot.sizeOfBarChartArray());
                assertEquals(1, plot.sizeOfLineChartArray());
                assertEquals(axisIds(plot.getBarChartArray(0)), lineAxisIds(plot.getLineChartArray(0)), "the same two axes");
                assertEquals(2, plot.sizeOfCatAxArray() + plot.sizeOfValAxArray(), "no extra axes");
                assertEquals("Sales!$C$2:$C$4", values(plot.getLineChartArray(0), 0));
                assertEquals("Target", plot.getLineChartArray(0).getSerArray(0).getTx().getV(), "the series title");
                assertEquals("Revenue", plot.getBarChartArray(0).getSerArray(0).getTx().getV());
            }
        }

        @Test
        void everySeriesHasItsOwnIndexAndOrder() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                List<Long> idx = new ArrayList<>();
                List<Long> order = new ArrayList<>();
                for (var bar : plot.getBarChartArray()) for (var s : bar.getSerList()) { idx.add(s.getIdx().getVal()); order.add(s.getOrder().getVal()); }
                for (var line : plot.getLineChartArray()) for (var s : line.getSerList()) { idx.add(s.getIdx().getVal()); order.add(s.getOrder().getVal()); }
                assertEquals(3, idx.size());
                assertEquals(3, idx.stream().distinct().count(), "unique idx: " + idx);
                assertEquals(3, order.stream().distinct().count(), "unique order: " + order);
            }
        }

        @Test
        void aLineChartTakesAnotherLineAndAnAreaChartToo() throws IOException {
            for (var type : List.of(ChartType.LINE, ChartType.AREA)) {
                var w = writer(c -> c.type(type).categoryColumn(0).valueColumn(1, "Revenue").lineColumn(2, "Target"));
                try (var wb = open(w.write(MONTHS))) {
                    var plot = plot(wb, 0);
                    assertEquals(type == ChartType.LINE ? 2 : 1, plot.sizeOfLineChartArray(), type.name());
                    assertEquals(type == ChartType.AREA ? 1 : 0, plot.sizeOfAreaChartArray(), type.name());
                    assertEquals("Sales!$C$2:$C$4", values(plot.getLineChartArray(plot.sizeOfLineChartArray() - 1), 0), type.name());
                }
            }
        }

        @Test
        void severalLinesShareOneLineGroup() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").lineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertEquals(1, plot.sizeOfLineChartArray());
                assertEquals(2, plot.getLineChartArray(0).sizeOfSerArray());
            }
        }
    }

    @Nested
    class OnTheSecondaryAxis {
        @Test
        void aSecondaryLineGetsAValueAxisOnTheRightAndAHiddenCategoryAxis() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .secondaryLineColumn(3, "Growth").secondaryValueAxisTitle("Growth %").valueAxisTitle("Revenue"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertEquals(1, plot.sizeOfBarChartArray());
                assertEquals(1, plot.sizeOfLineChartArray());
                assertEquals(2, plot.sizeOfCatAxArray());
                assertEquals(2, plot.sizeOfValAxArray());
                var line = plot.getLineChartArray(0);
                var barAxes = axisIds(plot.getBarChartArray(0));
                var lineAxes = lineAxisIds(line);
                assertTrue(lineAxes.stream().noneMatch(barAxes::contains), "the line has axes of its own: " + lineAxes + " vs " + barAxes);

                var second = valueAxis(plot, lineAxes.get(1));
                assertEquals(STAxPos.R, second.getAxPos().getVal(), "on the right");
                assertEquals(STCrosses.MAX, second.getCrosses().getVal(), "crossing at the maximum, so it sits on the right");
                assertEquals(lineAxes.get(0), second.getCrossAx().getVal());
                assertNotNull(second.getTitle(), "the secondary axis title");
                var primary = valueAxis(plot, barAxes.get(1));
                assertEquals(STAxPos.L, primary.getAxPos().getVal());
                assertNotNull(primary.getTitle());

                var secondCategory = categoryAxis(plot, lineAxes.get(0));
                assertTrue(secondCategory.getDelete().getVal(), "the second category axis is not shown");
                assertEquals(lineAxes.get(1), secondCategory.getCrossAx().getVal());
                assertFalse(categoryAxis(plot, barAxes.get(0)).getDelete().getVal(), "the main one is");
                assertEquals("Sales!$D$2:$D$4", values(line, 0));
            }
        }

        @Test
        void aPrimaryAndASecondaryLineAreTwoLineGroups() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertEquals(2, plot.sizeOfLineChartArray());
                var barAxes = axisIds(plot.getBarChartArray(0));
                assertEquals(barAxes, lineAxisIds(plot.getLineChartArray(0)));
                assertTrue(lineAxisIds(plot.getLineChartArray(1)).stream().noneMatch(barAxes::contains));
                assertEquals(4, plot.sizeOfCatAxArray() + plot.sizeOfValAxArray());
            }
        }

        @Test
        void theSecondaryAxisTitleIsOptional() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                var axes = lineAxisIds(plot.getLineChartArray(0));
                assertFalse(valueAxis(plot, axes.get(1)).isSetTitle());
            }
        }

        @Test
        void aLineChartAndAnAreaChartTakeASecondaryLineToo() throws IOException {
            for (var type : List.of(ChartType.LINE, ChartType.AREA)) {
                var w = writer(c -> c.type(type).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
                try (var wb = open(w.write(MONTHS))) {
                    var plot = plot(wb, 0);
                    assertEquals(4, plot.sizeOfCatAxArray() + plot.sizeOfValAxArray(), type.name());
                }
            }
        }
    }

    @Nested
    class WithTheRestOfTheChart {
        @Test
        void aChartWithoutLinesIsAsBefore() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertEquals(1, plot.sizeOfBarChartArray());
                assertEquals(0, plot.sizeOfLineChartArray());
                assertEquals(2, plot.sizeOfCatAxArray() + plot.sizeOfValAxArray());
            }
        }

        @Test
        void dataLabelsAndALegendCoverTheLinesToo() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").secondaryLineColumn(3, "Growth")
                    .showDataLabels(true).legendPosition(ExcelChartConfig.LegendPosition.BOTTOM).title("Sales"));
            try (var wb = open(w.write(MONTHS))) {
                var plot = plot(wb, 0);
                assertTrue(plot.getBarChartArray(0).getDLbls().getShowVal().getVal());
                for (var line : plot.getLineChartArray()) assertTrue(line.getDLbls().getShowVal().getVal());
                var chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
                assertNotNull(chart.getCTChart().getLegend());
            }
        }

        @Test
        void aSummaryAndTheDefaultPositionKeepTheChartBelowTheRows() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"))
                    .summary(s -> s.label("Month", "Total").sum("Revenue"));
            try (var wb = open(w.write(MONTHS))) {
                var anchor = wb.getSheetAt(0).getDrawingPatriarch().getShapes().get(0).getAnchor();
                assertTrue(((org.apache.poi.xssf.usermodel.XSSFClientAnchor) anchor).getRow1() > 5, "below the summary row");
            }
        }

        @Test
        void everyRolledOverSheetGetsAComboChartOverItsOwnRows() throws IOException {
            var rows = IntStream.range(0, 7).mapToObj(i -> new Month("m" + i, i, i, i / 10.0)).toList();
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").secondaryLineColumn(3, "Growth")).maxRows(3);
            try (var wb = open(w.write(rows))) {
                assertEquals(3, wb.getNumberOfSheets());
                for (int sheet = 0; sheet < 3; sheet++) {
                    var plot = plot(wb, sheet);
                    assertEquals(2, plot.sizeOfLineChartArray(), "sheet " + sheet);
                    int last = sheet < 2 ? 4 : 2;
                    String name = wb.getSheetName(sheet);
                    String quoted = name.contains(" ") ? "'" + name + "'" : name;
                    // POI writes a range of one row as a single cell
                    String tail = last == 2 ? "" : ":$%s$" + last;
                    assertEquals(quoted + "!$C$2" + tail.replace("%s", "C"), values(plot.getLineChartArray(0), 0), "primary line of sheet " + sheet);
                    assertEquals(quoted + "!$D$2" + tail.replace("%s", "D"), values(plot.getLineChartArray(1), 0), "secondary line of sheet " + sheet);
                }
            }
        }

        @Test
        void aWorkbookSheetWritesAComboChartToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Month>sheet("Sales")
                        .column("Month", Month::name)
                        .column("Revenue", Month::revenue, c -> c.type(ExcelDataType.INTEGER))
                        .column("Growth", Month::growth, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                        .chart(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(2, "Growth"))
                        .write(MONTHS);
                try (var wb = open(workbook.finish())) {
                    assertEquals(1, plot(wb, 0).sizeOfLineChartArray());
                }
            }
        }

        @Test
        void aWholeWorkbookWithoutTheDataRowsHasNoChartAndNoError() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(Stream.<Month>empty()))) {
                assertNull(wb.getSheetAt(0).getDrawingPatriarch(), "nothing to plot");
            }
        }
    }

    @Nested
    class LookOfTheChart {
        @Test
        void columnsSitBetweenTheTickMarksOnEveryValueAxis() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                for (var axis : plot(wb, 0).getValAxList()) {
                    assertEquals(org.openxmlformats.schemas.drawingml.x2006.chart.STCrossBetween.BETWEEN, axis.getCrossBetween().getVal(),
                            "the first and the last column are not cut in half");
                }
            }
        }

        @Test
        void aPlainBarChartHasItsColumnsBetweenTheTickMarksToo() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue"));
            try (var wb = open(w.write(MONTHS))) {
                assertEquals(org.openxmlformats.schemas.drawingml.x2006.chart.STCrossBetween.BETWEEN,
                        plot(wb, 0).getValAxArray(0).getCrossBetween().getVal());
            }
        }

        @Test
        void aLineChartKeepsItsPointsOnTheTickMarksAndASecondaryAxisMatchesIt() throws IOException {
            var w = writer(c -> c.type(ChartType.LINE).categoryColumn(0).valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
            try (var wb = open(w.write(MONTHS))) {
                for (var axis : plot(wb, 0).getValAxList()) {
                    assertEquals(org.openxmlformats.schemas.drawingml.x2006.chart.STCrossBetween.MID_CAT, axis.getCrossBetween().getVal());
                }
            }
        }

        @Test
        void linesAreNotSmoothed() throws IOException {
            var combo = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue")
                    .lineColumn(2, "Target").secondaryLineColumn(3, "Growth"));
            try (var wb = open(combo.write(MONTHS))) {
                int lines = 0;
                for (var group : plot(wb, 0).getLineChartArray()) {
                    for (var ser : group.getSerList()) {
                        lines++;
                        assertFalse(ser.getSmooth().getVal(), "a smoothed line shows values that are not in the data");
                    }
                }
                assertEquals(2, lines);
            }
            var plain = writer(c -> c.type(ChartType.LINE).categoryColumn(0).valueColumn(1, "Revenue").valueColumn(2, "Target"));
            try (var wb = open(plain.write(MONTHS))) {
                for (var ser : plot(wb, 0).getLineChartArray(0).getSerList()) assertFalse(ser.getSmooth().getVal());
            }
        }
    }

    @Nested
    class Refusals {
        @Test
        void linesNeedABarLineOrAreaChart() {
            for (var type : List.of(ChartType.PIE, ChartType.DOUGHNUT, ChartType.SCATTER)) {
                var w = writer(c -> c.type(type).categoryColumn(0).valueColumn(1, "Revenue").lineColumn(2, "Target"));
                var e = assertThrows(ExcelWriteException.class, () -> w.write(MONTHS), type.name());
                assertTrue(e.getMessage().contains("BAR, LINE or AREA"), e.getMessage());
            }
        }

        @Test
        void linesNeedVerticalBars() {
            var w = writer(c -> c.type(ChartType.BAR).barDirection(BarDirection.HORIZONTAL).categoryColumn(0)
                    .valueColumn(1, "Revenue").secondaryLineColumn(3, "Growth"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(MONTHS));
            assertTrue(e.getMessage().contains("vertical"), e.getMessage());
        }

        @Test
        void linesNeedAMainSeries() {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).lineColumn(2, "Target"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(MONTHS));
            assertTrue(e.getMessage().contains("valueColumn"), e.getMessage());
        }

        @Test
        void aLineColumnThatDoesNotExistIsNamed() {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "Revenue").lineColumn(9, "Ghost"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(MONTHS));
            assertTrue(e.getMessage().contains("Ghost") && e.getMessage().contains("9"), e.getMessage());
        }

        @Test
        void argumentsAreChecked() {
            var config = new ExcelChartConfig();
            assertThrows(NullPointerException.class, () -> config.lineColumn(1, null));
            assertThrows(NullPointerException.class, () -> config.secondaryLineColumn(1, null));
            assertThrows(IllegalArgumentException.class, () -> config.lineColumn(-1, "x"));
            assertThrows(IllegalArgumentException.class, () -> config.secondaryLineColumn(-1, "x"));
            assertDoesNotThrow(() -> config.secondaryValueAxisTitle(null));
        }
    }

    private static org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx valueAxis(CTPlotArea plot, long id) {
        for (var ax : plot.getValAxList()) if (ax.getAxId().getVal() == id) return ax;
        throw new AssertionError("no value axis " + id);
    }

    private static org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx categoryAxis(CTPlotArea plot, long id) {
        for (var ax : plot.getCatAxList()) if (ax.getAxId().getVal() == id) return ax;
        throw new AssertionError("no category axis " + id);
    }
}
