package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/** The distance and the number format of the labels of a chart's value axis: both set, and neither written when not set. */
class ChartValueAxisScaleTest {

    private record Item(String name, int qty) {
    }

    private static List<CTValAx> valueAxes(ExcelChartConfig.ChartType type, Consumer<ExcelChartConfig> config) throws Exception {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name)
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .chart(ch -> {
                    ch.type(type).categoryColumn(0).valueColumn(1, "Qty");
                    config.accept(ch);
                })
                .write(IntStream.range(0, 4).mapToObj(i -> new Item("item " + i, i)).toList()).writeTo(out);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            return List.copyOf(chart.getCTChart().getPlotArea().getValAxList());
        }
    }

    @Test
    void withoutTheOptionsNothingIsWrittenAndTheLabelsFollowTheCells() throws Exception {
        for (var type : List.of(ExcelChartConfig.ChartType.BAR, ExcelChartConfig.ChartType.LINE, ExcelChartConfig.ChartType.AREA)) {
            CTValAx axis = valueAxes(type, ch -> { }).get(0);
            assertFalse(axis.isSetMajorUnit(), type + ": the distance is Excel's");
            assertFalse(axis.isSetNumFmt(), type + ": the labels follow the cells");
        }
    }

    @Test
    void theMajorUnitIsWrittenOnTheValueAxisOfBarLineAndAreaCharts() throws Exception {
        for (var type : List.of(ExcelChartConfig.ChartType.BAR, ExcelChartConfig.ChartType.LINE, ExcelChartConfig.ChartType.AREA)) {
            CTValAx axis = valueAxes(type, ch -> ch.valueAxisMajorUnit(1)).get(0);
            assertTrue(axis.isSetMajorUnit(), type.toString());
            assertEquals(1.0, axis.getMajorUnit().getVal(), type.toString());
        }
    }

    @Test
    void theNumberFormatIsWrittenAndIsNotLinkedToTheCells() throws Exception {
        CTValAx axis = valueAxes(ExcelChartConfig.ChartType.BAR, ch -> ch.valueAxisNumberFormat("0.0")).get(0);
        assertTrue(axis.isSetNumFmt());
        assertEquals("0.0", axis.getNumFmt().getFormatCode());
        assertFalse(axis.getNumFmt().getSourceLinked(), "the format given is meant, not the cells'");
    }

    @Test
    void inAScatterChartTheOptionsGoToTheVerticalAxis() throws Exception {
        var axes = valueAxes(ExcelChartConfig.ChartType.SCATTER, ch -> ch.valueAxisMajorUnit(2).valueAxisNumberFormat("0"));
        assertEquals(2, axes.size());
        var withUnit = axes.stream().filter(CTValAx::isSetMajorUnit).toList();
        assertEquals(1, withUnit.size(), "only one of the two value axes");
        assertEquals("l", withUnit.get(0).getAxPos().getVal().toString(), "the left (vertical) axis");
        assertEquals(2.0, withUnit.get(0).getMajorUnit().getVal());
    }

    @Test
    void invalidValuesAreRefused() {
        var config = new ExcelChartConfig();
        assertThrows(IllegalArgumentException.class, () -> config.valueAxisMajorUnit(0));
        assertThrows(IllegalArgumentException.class, () -> config.valueAxisMajorUnit(-1));
        assertThrows(IllegalArgumentException.class, () -> config.valueAxisMajorUnit(Double.NaN));
        assertThrows(IllegalArgumentException.class, () -> config.valueAxisMajorUnit(Double.POSITIVE_INFINITY));
        assertThrows(IllegalArgumentException.class, () -> config.valueAxisNumberFormat(" "));
        assertThrows(NullPointerException.class, () -> config.valueAxisNumberFormat(null));
    }
}
