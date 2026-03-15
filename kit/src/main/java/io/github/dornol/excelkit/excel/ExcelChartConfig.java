package io.github.dornol.excelkit.excel;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

/**
 * Builder for creating charts in an Excel sheet.
 * <p>
 * Supports bar, line, and pie charts. Charts are rendered using Apache POI's XDDF chart API
 * and reference cell ranges for their data.
 *
 * <pre>{@code
 * new ExcelWriter<Product>()
 *     .addColumn("Name", Product::getName)
 *     .addColumn("Sales", Product::getSales, c -> c.type(ExcelDataType.INTEGER))
 *     .chart(chart -> chart
 *         .type(ExcelChartConfig.ChartType.BAR)
 *         .title("Sales by Product")
 *         .categoryColumn(0)
 *         .valueColumn(1, "Sales")
 *         .categoryAxisTitle("Product")
 *         .valueAxisTitle("Amount")
 *         .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM)
 *         .barDirection(ExcelChartConfig.BarDirection.HORIZONTAL)
 *         .barGrouping(ExcelChartConfig.BarGrouping.STACKED)
 *         .showDataLabels(true)
 *         .position(3, 0, 20, 8))
 *     .write(stream)
 *     .consumeOutputStream(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelChartConfig {

    /**
     * Supported chart types.
     */
    public enum ChartType {
        BAR,
        LINE,
        PIE
    }

    /**
     * Legend position relative to the chart area.
     *
     * @since 0.7.0
     */
    public enum LegendPosition {
        BOTTOM,
        LEFT,
        RIGHT,
        TOP,
        TOP_RIGHT
    }

    /**
     * Bar chart grouping mode. Determines how multiple series are arranged.
     *
     * @since 0.7.0
     */
    public enum BarGrouping {
        /** Each series is drawn side by side. */
        STANDARD,
        /** Series are stacked on top of each other. */
        STACKED,
        /** Series are stacked and scaled to 100%. */
        PERCENT_STACKED
    }

    /**
     * Bar chart direction. Controls whether bars are drawn vertically (columns) or horizontally.
     *
     * @since 0.7.0
     */
    public enum BarDirection {
        /** Bars are drawn as vertical columns. */
        VERTICAL,
        /** Bars are drawn as horizontal bars. */
        HORIZONTAL
    }

    private ChartType chartType = ChartType.BAR;
    private String title;
    private int categoryColumnIndex = 0;
    private final List<ValueSeries> valueSeries = new ArrayList<>();
    private int anchorCol1 = 0;
    private int anchorRow1 = -1;
    private int anchorCol2 = 8;
    private int anchorRow2 = -1;
    private String categoryAxisTitle;
    private String valueAxisTitle;
    private LegendPosition legendPosition;
    private boolean showDataLabels = false;
    private BarGrouping barGrouping = BarGrouping.STANDARD;
    private BarDirection barDirection = BarDirection.VERTICAL;

    /**
     * Sets the chart type.
     *
     * @param type the chart type
     * @return this config for chaining
     */
    public ExcelChartConfig type(ChartType type) {
        this.chartType = type;
        return this;
    }

    /**
     * Sets the chart title.
     *
     * @param title the chart title
     * @return this config for chaining
     */
    public ExcelChartConfig title(String title) {
        this.title = title;
        return this;
    }

    /**
     * Sets the category (X-axis) column index (0-based).
     *
     * @param columnIndex 0-based column index for categories
     * @return this config for chaining
     */
    public ExcelChartConfig categoryColumn(int columnIndex) {
        this.categoryColumnIndex = columnIndex;
        return this;
    }

    /**
     * Adds a value (Y-axis) series from the specified column.
     *
     * @param columnIndex 0-based column index for values
     * @param seriesTitle title of this data series
     * @return this config for chaining
     */
    public ExcelChartConfig valueColumn(int columnIndex, String seriesTitle) {
        this.valueSeries.add(new ValueSeries(columnIndex, seriesTitle));
        return this;
    }

    /**
     * Sets the chart position in the sheet.
     *
     * @param col1 starting column (0-based)
     * @param row1 starting row (0-based)
     * @param col2 ending column (0-based)
     * @param row2 ending row (0-based)
     * @return this config for chaining
     */
    public ExcelChartConfig position(int col1, int row1, int col2, int row2) {
        this.anchorCol1 = col1;
        this.anchorRow1 = row1;
        this.anchorCol2 = col2;
        this.anchorRow2 = row2;
        return this;
    }

    /**
     * Sets the category axis (X-axis) title.
     *
     * @param title the axis title, or {@code null} to omit
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig categoryAxisTitle(String title) {
        this.categoryAxisTitle = title;
        return this;
    }

    /**
     * Sets the value axis (Y-axis) title.
     *
     * @param title the axis title, or {@code null} to omit
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig valueAxisTitle(String title) {
        this.valueAxisTitle = title;
        return this;
    }

    /**
     * Sets the legend position. If {@code null} (the default), no legend is added.
     *
     * @param position the legend position, or {@code null} for no legend
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig legendPosition(LegendPosition position) {
        this.legendPosition = position;
        return this;
    }

    /**
     * Sets whether data labels (values) are shown on chart data points.
     * <p>
     * Note: data label rendering is best-effort and depends on the chart type
     * and Apache POI's XDDF support.
     *
     * @param show {@code true} to show data labels
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig showDataLabels(boolean show) {
        this.showDataLabels = show;
        return this;
    }

    /**
     * Sets the bar grouping mode for bar charts.
     * Ignored for non-bar chart types.
     *
     * @param grouping the bar grouping mode
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig barGrouping(BarGrouping grouping) {
        this.barGrouping = grouping;
        return this;
    }

    /**
     * Sets the bar direction for bar charts.
     * {@link BarDirection#VERTICAL} renders columns; {@link BarDirection#HORIZONTAL} renders horizontal bars.
     * Ignored for non-bar chart types.
     *
     * @param direction the bar direction
     * @return this config for chaining
     * @since 0.7.0
     */
    public ExcelChartConfig barDirection(BarDirection direction) {
        this.barDirection = direction;
        return this;
    }

    /**
     * Creates the chart on the given sheet.
     * Package-private, called by the writer after data is written.
     *
     * @param sheet      the current SXSSFSheet
     * @param dataEndRow the last data row index (0-based)
     * @param headerRow  the header row index (0-based)
     */
    void apply(SXSSFSheet sheet, int headerRow, int dataEndRow) {
        if (valueSeries.isEmpty()) return;

        // Access underlying XSSFSheet for chart support
        XSSFSheet xssfSheet;
        try {
            var field = SXSSFSheet.class.getDeclaredField("_sh");
            field.setAccessible(true);
            xssfSheet = (XSSFSheet) field.get(sheet);
        } catch (Exception e) {
            throw new ExcelWriteException("Failed to access underlying XSSFSheet for chart creation", e);
        }

        int row1 = anchorRow1 >= 0 ? anchorRow1 : dataEndRow + 2;
        int row2 = anchorRow2 >= 0 ? anchorRow2 : row1 + 15;

        XSSFDrawing drawing = xssfSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                anchorCol1, row1, anchorCol2, row2);
        XSSFChart chart = drawing.createChart(anchor);

        if (title != null) {
            chart.setTitleText(title);
        }

        int dataStartRow = headerRow + 1;
        String sheetName = xssfSheet.getSheetName();

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                xssfSheet,
                new org.apache.poi.ss.util.CellRangeAddress(
                        dataStartRow, dataEndRow,
                        categoryColumnIndex, categoryColumnIndex));

        switch (chartType) {
            case BAR -> createBarChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case LINE -> createLineChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case PIE -> createPieChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
        }

        applyLegend(chart);
        applyDataLabels(chart);
    }

    private void createBarChart(XSSFChart chart, XDDFDataSource<String> categories,
                                XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        applyAxisTitles(categoryAxis, valueAxis);

        XDDFBarChartData data = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, categoryAxis, valueAxis);
        data.setBarDirection(barDirection == BarDirection.HORIZONTAL
                ? org.apache.poi.xddf.usermodel.chart.BarDirection.BAR
                : org.apache.poi.xddf.usermodel.chart.BarDirection.COL);
        applyBarGrouping(data);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    private void createLineChart(XSSFChart chart, XDDFDataSource<String> categories,
                                 XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        applyAxisTitles(categoryAxis, valueAxis);

        XDDFLineChartData data = (XDDFLineChartData) chart.createData(
                ChartTypes.LINE, categoryAxis, valueAxis);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    private void createPieChart(XSSFChart chart, XDDFDataSource<String> categories,
                                XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    private void applyAxisTitles(XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis) {
        if (categoryAxisTitle != null) {
            categoryAxis.setTitle(categoryAxisTitle);
        }
        if (valueAxisTitle != null) {
            valueAxis.setTitle(valueAxisTitle);
        }
    }

    private void applyLegend(XSSFChart chart) {
        if (legendPosition == null) {
            return;
        }
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(mapLegendPosition(legendPosition));
    }

    private static org.apache.poi.xddf.usermodel.chart.LegendPosition mapLegendPosition(LegendPosition position) {
        return switch (position) {
            case BOTTOM -> org.apache.poi.xddf.usermodel.chart.LegendPosition.BOTTOM;
            case LEFT -> org.apache.poi.xddf.usermodel.chart.LegendPosition.LEFT;
            case RIGHT -> org.apache.poi.xddf.usermodel.chart.LegendPosition.RIGHT;
            case TOP -> org.apache.poi.xddf.usermodel.chart.LegendPosition.TOP;
            case TOP_RIGHT -> org.apache.poi.xddf.usermodel.chart.LegendPosition.TOP_RIGHT;
        };
    }

    private void applyBarGrouping(XDDFBarChartData data) {
        if (barGrouping == null) {
            return;
        }
        switch (barGrouping) {
            case STANDARD -> data.setBarGrouping(org.apache.poi.xddf.usermodel.chart.BarGrouping.STANDARD);
            case STACKED -> data.setBarGrouping(org.apache.poi.xddf.usermodel.chart.BarGrouping.STACKED);
            case PERCENT_STACKED -> data.setBarGrouping(org.apache.poi.xddf.usermodel.chart.BarGrouping.PERCENT_STACKED);
        }
    }

    /**
     * Applies data labels to the chart if {@link #showDataLabels} is enabled.
     * <p>
     * This uses the low-level CT (XML bean) API because POI's XDDF layer does not
     * expose a convenient method for toggling data labels on all series.
     * This is a best-effort feature.
     */
    private void applyDataLabels(XSSFChart chart) {
        if (!showDataLabels) {
            return;
        }
        try {
            var ctChart = chart.getCTChart();
            var plotArea = ctChart.getPlotArea();

            // Apply data labels to bar charts
            for (var barChart : plotArea.getBarChartList()) {
                var dLbls = barChart.isSetDLbls() ? barChart.getDLbls() : barChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }

            // Apply data labels to line charts
            for (var lineChart : plotArea.getLineChartList()) {
                var dLbls = lineChart.isSetDLbls() ? lineChart.getDLbls() : lineChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }

            // Apply data labels to pie charts
            for (var pieChart : plotArea.getPieChartList()) {
                var dLbls = pieChart.isSetDLbls() ? pieChart.getDLbls() : pieChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }
        } catch (Exception e) {
            // Best-effort: if CT API is unavailable or fails, skip data labels silently
        }
    }

    private void addSeries(XDDFChartData chartData, XDDFDataSource<String> categories,
                           XSSFSheet sheet, int dataStart, int dataEnd) {
        for (ValueSeries vs : valueSeries) {
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet,
                    new org.apache.poi.ss.util.CellRangeAddress(
                            dataStart, dataEnd,
                            vs.columnIndex, vs.columnIndex));
            XDDFChartData.Series series = chartData.addSeries(categories, values);
            series.setTitle(vs.title, null);
        }
    }

    private record ValueSeries(int columnIndex, String title) {
    }
}
