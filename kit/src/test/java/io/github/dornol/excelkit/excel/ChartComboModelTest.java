package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.ExcelChartConfig.ChartType;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Random;
import java.util.Set;
import java.util.function.Consumer;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Random combo charts compared with what was asked for: the series, their ranges and titles, the axes they sit on,
 * across every chart type that takes lines, with rollover, a summary and a title.
 */
class ChartComboModelTest {
    record R(String name, int a, int b, int c, int d) {}

    private static List<R> rows(int n) {
        return IntStream.range(0, n).mapToObj(i -> new R("n" + i, i, i * 2, i * 3, i * 4)).toList();
    }

    private static ExcelWriter<R> writer(Consumer<ExcelChartConfig> chart) {
        return ExcelWriter.<R>create().sheetName("Data")
                .column("Name", R::name)
                .column("A", R::a, c -> c.type(ExcelDataType.INTEGER))
                .column("B", R::b, c -> c.type(ExcelDataType.INTEGER))
                .column("C", R::c, c -> c.type(ExcelDataType.INTEGER))
                .column("D", R::d, c -> c.type(ExcelDataType.INTEGER))
                .chart(chart);
    }

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    /** A range as POI writes it: one row is a single cell. */
    private static String range(int col, int rowCount) {
        String c = column(col);
        return rowCount == 1 ? "Data!$" + c + "$2" : "Data!$" + c + "$2:$" + c + "$" + (1 + rowCount);
    }

    private static String column(int index) {
        return String.valueOf((char) ('A' + index));
    }

    /** A series as the chart XML states it. */
    record Seen(String title, String range, long idx, long order, List<Long> axes) {}

    private static List<Seen> seen(CTPlotArea plot) {
        List<Seen> out = new ArrayList<>();
        for (var bar : plot.getBarChartArray()) {
            List<Long> axes = bar.getAxIdList().stream().map(a -> a.getVal()).toList();
            for (var s : bar.getSerList()) out.add(new Seen(s.getTx().getV(), s.getVal().getNumRef().getF(), s.getIdx().getVal(), s.getOrder().getVal(), axes));
        }
        for (var line : plot.getLineChartArray()) {
            List<Long> axes = line.getAxIdList().stream().map(a -> a.getVal()).toList();
            for (var s : line.getSerList()) out.add(new Seen(s.getTx().getV(), s.getVal().getNumRef().getF(), s.getIdx().getVal(), s.getOrder().getVal(), axes));
        }
        for (var area : plot.getAreaChartArray()) {
            List<Long> axes = area.getAxIdList().stream().map(a -> a.getVal()).toList();
            for (var s : area.getSerList()) out.add(new Seen(s.getTx().getV(), s.getVal().getNumRef().getF(), s.getIdx().getVal(), s.getOrder().getVal(), axes));
        }
        return out;
    }

    @Nested
    class RandomCharts {
        @Test
        void everySeriesAskedForIsThereOnTheRightAxes() throws IOException {
            Random random = new Random(20260601);
            for (int round = 0; round < 250; round++) {
                ChartType type = new ChartType[]{ChartType.BAR, ChartType.LINE, ChartType.AREA}[random.nextInt(3)];
                int main = 1 + random.nextInt(2);
                int primaryLines = random.nextInt(3);
                int secondaryLines = random.nextInt(3);
                int rowCount = 1 + random.nextInt(12);
                boolean summary = random.nextBoolean();
                String desc = "round " + round + " " + type + " main=" + main + " p=" + primaryLines + " s=" + secondaryLines
                        + " rows=" + rowCount + " summary=" + summary;

                List<String> expectedTitles = new ArrayList<>();
                List<String> expectedRanges = new ArrayList<>();
                var w = writer(c -> {
                    c.type(type).categoryColumn(0);
                    for (int i = 0; i < main; i++) {
                        c.valueColumn(1 + i, "m" + i);
                        expectedTitles.add("m" + i);
                        expectedRanges.add(range(1 + i, rowCount));
                    }
                    int col = 1;
                    for (int i = 0; i < primaryLines; i++, col = col % 4 + 1) {
                        c.lineColumn(col, "p" + i);
                        expectedTitles.add("p" + i);
                        expectedRanges.add(range(col, rowCount));
                    }
                    for (int i = 0; i < secondaryLines; i++, col = col % 4 + 1) {
                        c.secondaryLineColumn(col, "s" + i);
                        expectedTitles.add("s" + i);
                        expectedRanges.add(range(col, rowCount));
                    }
                });
                if (summary) w.summary(s -> s.sum("A"));
                try (var wb = open(w.write(rows(rowCount)))) {
                    List<XSSFChart> charts = wb.getSheetAt(0).getDrawingPatriarch().getCharts();
                    assertEquals(1, charts.size(), desc);
                    var plot = charts.get(0).getCTChart().getPlotArea();
                    List<Seen> series = seen(plot);

                    assertEquals(expectedTitles.size(), series.size(), desc);
                    assertEquals(new HashSet<>(expectedTitles), new HashSet<>(series.stream().map(Seen::title).toList()), desc);
                    for (Seen s : series) {
                        int at = expectedTitles.indexOf(s.title());
                        assertEquals(expectedRanges.get(at), s.range(), desc + " " + s.title());
                    }
                    assertEquals(series.size(), series.stream().map(Seen::idx).distinct().count(), desc + " idx");
                    assertEquals(series.size(), series.stream().map(Seen::order).distinct().count(), desc + " order");

                    // axes: mains + primary lines share one pair, secondary lines have another
                    Set<List<Long>> pairs = new HashSet<>();
                    for (Seen s : series) pairs.add(s.axes());
                    assertEquals(secondaryLines > 0 ? 2 : 1, pairs.size(), desc + " axis pairs " + pairs);
                    int axes = plot.sizeOfCatAxArray() + plot.sizeOfValAxArray();
                    assertEquals(secondaryLines > 0 ? 4 : 2, axes, desc);
                    for (Seen s : series) {
                        if (s.title().startsWith("s")) {
                            assertTrue(series.stream().filter(o -> !o.title().startsWith("s"))
                                    .noneMatch(o -> o.axes().stream().anyMatch(s.axes()::contains)), desc + " " + s.title());
                        }
                    }
                }
            }
        }
    }

    @Nested
    class Rollover {
        @Test
        void everySheetHasItsChartOverItsOwnRowsWhateverTheSeries() throws IOException {
            Random random = new Random(7);
            for (int round = 0; round < 40; round++) {
                int maxRows = 2 + random.nextInt(6);
                int rowCount = 1 + random.nextInt(25);
                boolean secondary = random.nextBoolean();
                var w = writer(c -> {
                    c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "A").lineColumn(2, "B");
                    if (secondary) c.secondaryLineColumn(3, "C");
                }).maxRows(maxRows);
                try (var wb = open(w.write(rows(rowCount)))) {
                    int sheets = (rowCount + maxRows - 1) / maxRows;
                    assertEquals(sheets, wb.getNumberOfSheets(), "round " + round);
                    for (int s = 0; s < sheets; s++) {
                        int rowsHere = Math.min(maxRows, rowCount - s * maxRows);
                        String name = wb.getSheetName(s);
                        var plot = wb.getSheetAt(s).getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea();
                        for (Seen seen : seen(plot)) {
                            assertTrue(rowsHere == 1 ? seen.range().endsWith("$2") : seen.range().endsWith("$" + (1 + rowsHere)),
                                    "round " + round + " sheet " + s + " ends at the last row of that sheet: " + seen.range());
                            assertTrue(seen.range().contains(name.contains(" ") ? "'" + name + "'" : name),
                                    "the sheet's own name: " + seen.range() + " vs " + name);
                        }
                    }
                }
            }
        }
    }

    @Nested
    class Columns {
        @Test
        void aCategoryColumnThatDoesNotExistIsNamed() {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(9).valueColumn(1, "A"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(rows(3)));
            assertTrue(e.getMessage().contains("category column 9") && e.getMessage().contains("5 columns"), e.getMessage());
        }

        @Test
        void aValueColumnThatDoesNotExistIsNamed() {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(7, "Ghost"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(rows(3)));
            assertTrue(e.getMessage().contains("value column 7") && e.getMessage().contains("Ghost"), e.getMessage());
        }

        @Test
        void theFirstColumnPastTheLastIsRefusedForEveryKindOfColumn() {
            // the sheet has 5 columns: index 4 is the last one, index 5 is the first that does not exist
            var category = writer(c -> c.type(ChartType.BAR).categoryColumn(5).valueColumn(1, "A"));
            assertThrows(ExcelWriteException.class, () -> category.write(rows(3)).writeTo(java.io.OutputStream.nullOutputStream()));
            var value = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(5, "A"));
            assertThrows(ExcelWriteException.class, () -> value.write(rows(3)).writeTo(java.io.OutputStream.nullOutputStream()));
            var line = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "A").lineColumn(5, "L"));
            assertThrows(ExcelWriteException.class, () -> line.write(rows(3)).writeTo(java.io.OutputStream.nullOutputStream()));
            var secondary = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "A").secondaryLineColumn(5, "L"));
            assertThrows(ExcelWriteException.class, () -> secondary.write(rows(3)).writeTo(java.io.OutputStream.nullOutputStream()));
        }

        @Test
        void aSecondaryLineColumnThatDoesNotExistIsNamed() {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(0).valueColumn(1, "A").secondaryLineColumn(8, "Ghost"));
            var e = assertThrows(ExcelWriteException.class, () -> w.write(rows(3)));
            assertTrue(e.getMessage().contains("Ghost"), e.getMessage());
        }

        @Test
        void theLastColumnIsStillFine() throws IOException {
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(4).valueColumn(4, "D").lineColumn(4, "D2"));
            try (var wb = open(w.write(rows(3)))) {
                assertEquals(2, seen(wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea()).size());
            }
        }

        @Test
        void aFailedChartLeavesNoTempFileBehind() {
            // TempFileLeakGuard fails the test when a refused write leaves temp files.
            var w = writer(c -> c.type(ChartType.BAR).categoryColumn(9).valueColumn(1, "A"));
            assertThrows(ExcelWriteException.class, () -> w.write(rows(3)));
        }
    }

    @Nested
    class Positions {
        @Test
        void aWrongPositionIsRefused() {
            var config = new ExcelChartConfig();
            assertThrows(IllegalArgumentException.class, () -> config.position(-1, 0, 3, 3));
            assertThrows(IllegalArgumentException.class, () -> config.position(4, 0, 3, 3));
            assertThrows(IllegalArgumentException.class, () -> config.position(0, 5, 3, 4));
            assertDoesNotThrow(() -> config.position(0, -1, 3, -1));
            assertDoesNotThrow(() -> config.position(2, 1, 2, 1));
        }
    }
}
