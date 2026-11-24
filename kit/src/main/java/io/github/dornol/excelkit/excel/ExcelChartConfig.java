package io.github.dornol.excelkit.excel;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Builder for creating charts in an Excel sheet.
 * <p>
 * Supports bar, line, and pie charts. Charts are rendered using Apache POI's XDDF chart API
 * and reference cell ranges for their data.
 *
 * <pre>{@code
 * ExcelWriter.<Product>create()
 *     .column("Name", Product::getName)
 *     .column("Sales", Product::getSales, c -> c.type(ExcelDataType.INTEGER))
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
 *     .write(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelChartConfig {

    /** Creates a new chart configuration with defaults. */
    public ExcelChartConfig() {}

    /**
     * Supported chart types.
     */
    public enum ChartType {
        /** Bar chart. */
        BAR,
        /** Line chart. */
        LINE,
        /** Pie chart. */
        PIE,
        /** Scatter (XY) chart. */
        SCATTER,
        /** Area chart. */
        AREA,
        /** Doughnut chart. */
        DOUGHNUT
    }

    /**
     * Legend position relative to the chart area.
     *
     * @since 0.6.1
     */
    public enum LegendPosition {
        /** Bottom position. */
        BOTTOM,
        /** Left position. */
        LEFT,
        /** Right position. */
        RIGHT,
        /** Top position. */
        TOP,
        /** Top-right position. */
        TOP_RIGHT
    }

    /**
     * Bar chart grouping mode. Determines how multiple series are arranged.
     *
     * @since 0.6.1
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
     * @since 0.6.1
     */
    public enum BarDirection {
        /** Bars are drawn as vertical columns. */
        VERTICAL,
        /** Bars are drawn as horizontal bars. */
        HORIZONTAL
    }

    private ChartType chartType = ChartType.BAR;
    private @Nullable String title;
    private int categoryColumnIndex = 0;
    private final List<ValueSeries> valueSeries = new ArrayList<>();
    /** Series drawn as lines over the main chart (a combination chart), on the primary or the secondary value axis. */
    private final List<ValueSeries> lineSeries = new ArrayList<>();
    private @Nullable String secondaryValueAxisTitle;
    private int anchorCol1 = 0;
    private int anchorRow1 = -1;
    private int anchorCol2 = 8;
    private int anchorRow2 = -1;
    private @Nullable String categoryAxisTitle;
    private @Nullable String valueAxisTitle;
    private @Nullable LegendPosition legendPosition;
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
        Objects.requireNonNull(type, "type cannot be null");
        this.chartType = type;
        return this;
    }

    /**
     * Sets the chart title.
     *
     * @param title the chart title
     * @return this config for chaining
     */
    public ExcelChartConfig title(@Nullable String title) {
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
        if (columnIndex < 0) {
            throw new IllegalArgumentException("categoryColumn index must be non-negative");
        }
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
        Objects.requireNonNull(seriesTitle, "seriesTitle cannot be null");
        if (columnIndex < 0) {
            throw new IllegalArgumentException("valueColumn index must be non-negative");
        }
        this.valueSeries.add(new ValueSeries(columnIndex, seriesTitle, false));
        return this;
    }

    /**
     * Adds a series drawn as a line over the chart, on the same value axis as the main series: with {@link ChartType#BAR}
     * columns a line over the bars, for a target or a trend that has the same unit. The chart must be a vertical
     * {@code BAR}, a {@code LINE} or an {@code AREA} chart, and have at least one {@link #valueColumn} (the main series).
     *
     * @param columnIndex 0-based column index for values
     * @param seriesTitle title of this data series
     * @return this config for chaining
     * @since 0.29.0
     */
    public ExcelChartConfig lineColumn(int columnIndex, String seriesTitle) {
        Objects.requireNonNull(seriesTitle, "seriesTitle cannot be null");
        if (columnIndex < 0) {
            throw new IllegalArgumentException("lineColumn index must be non-negative");
        }
        this.lineSeries.add(new ValueSeries(columnIndex, seriesTitle, false));
        return this;
    }

    /**
     * Adds a series drawn as a line over the chart on a second value axis on the right, for a quantity with another
     * unit or scale than the main series (columns of revenue with a line of the growth rate). See {@link #lineColumn}
     * for the chart types it needs; {@link #secondaryValueAxisTitle(String)} titles the axis.
     *
     * @param columnIndex 0-based column index for values
     * @param seriesTitle title of this data series
     * @return this config for chaining
     * @since 0.29.0
     */
    public ExcelChartConfig secondaryLineColumn(int columnIndex, String seriesTitle) {
        Objects.requireNonNull(seriesTitle, "seriesTitle cannot be null");
        if (columnIndex < 0) {
            throw new IllegalArgumentException("secondaryLineColumn index must be non-negative");
        }
        this.lineSeries.add(new ValueSeries(columnIndex, seriesTitle, true));
        return this;
    }

    /**
     * Sets the title of the secondary value axis (the one on the right that {@link #secondaryLineColumn} uses).
     *
     * @param title the axis title, or {@code null} to omit
     * @return this config for chaining
     * @since 0.29.0
     */
    public ExcelChartConfig secondaryValueAxisTitle(@Nullable String title) {
        this.secondaryValueAxisTitle = title;
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
        if (col1 < 0 || col2 < col1) {
            throw new IllegalArgumentException("position needs 0 <= col1 <= col2: " + col1 + ", " + col2);
        }
        // A negative row means "below the data"; two given rows must run downwards.
        if (row1 >= 0 && row2 >= 0 && row2 < row1) {
            throw new IllegalArgumentException("position needs row1 <= row2: " + row1 + ", " + row2);
        }
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
     * @since 0.6.1
     */
    public ExcelChartConfig categoryAxisTitle(@Nullable String title) {
        this.categoryAxisTitle = title;
        return this;
    }

    /**
     * Sets the value axis (Y-axis) title.
     *
     * @param title the axis title, or {@code null} to omit
     * @return this config for chaining
     * @since 0.6.1
     */
    public ExcelChartConfig valueAxisTitle(@Nullable String title) {
        this.valueAxisTitle = title;
        return this;
    }

    /**
     * Sets the legend position. If {@code null} (the default), no legend is added.
     *
     * @param position the legend position, or {@code null} for no legend
     * @return this config for chaining
     * @since 0.6.1
     */
    public ExcelChartConfig legendPosition(@Nullable LegendPosition position) {
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
     * @since 0.6.1
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
     * @since 0.6.1
     */
    public ExcelChartConfig barGrouping(BarGrouping grouping) {
        Objects.requireNonNull(grouping, "grouping cannot be null");
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
     * @since 0.6.1
     */
    public ExcelChartConfig barDirection(BarDirection direction) {
        Objects.requireNonNull(direction, "direction cannot be null");
        this.barDirection = direction;
        return this;
    }

    /**
     * Creates the chart on the given sheet.
     * Package-private, called by the writer after data is written.
     *
     * @param sheet        the current SXSSFSheet
     * @param dataEndRow   the last data row index (0-based)
     * @param headerRow    the header row index (0-based)
     * @param nextFreeRow  the first row after the data, summary rows and {@code afterData} content
     */
    void apply(SXSSFSheet sheet, int headerRow, int dataEndRow, int nextFreeRow, int columnCount) {
        if (categoryColumnIndex >= columnCount) {
            throw new ExcelWriteException("Chart category column " + categoryColumnIndex + " does not exist: the sheet has "
                    + columnCount + " columns");
        }
        for (ValueSeries vs : valueSeries) {
            if (vs.columnIndex >= columnCount) {
                throw new ExcelWriteException("Chart value column " + vs.columnIndex + " ('" + vs.title
                        + "') does not exist: the sheet has " + columnCount + " columns");
            }
        }
        for (ValueSeries vs : lineSeries) {
            if (vs.columnIndex >= columnCount) {
                throw new ExcelWriteException("Chart line column " + vs.columnIndex + " ('" + vs.title
                        + "') does not exist: the sheet has " + columnCount + " columns");
            }
        }
        if (!lineSeries.isEmpty()) {
            if (chartType != ChartType.BAR && chartType != ChartType.LINE && chartType != ChartType.AREA) {
                throw new ExcelWriteException("Line series (lineColumn, secondaryLineColumn) go over a BAR, LINE or AREA chart, "
                        + "not " + chartType);
            }
            if (chartType == ChartType.BAR && barDirection == BarDirection.HORIZONTAL) {
                throw new ExcelWriteException("Line series (lineColumn, secondaryLineColumn) need vertical bars: "
                        + "a line cannot run over horizontal ones");
            }
            if (valueSeries.isEmpty()) {
                throw new ExcelWriteException("Line series (lineColumn, secondaryLineColumn) go over the main series: "
                        + "add at least one valueColumn");
            }
        }
        if (valueSeries.isEmpty()) return;
        // Without data rows there is nothing to plot, and the range would be reversed.
        if (dataEndRow <= headerRow) return;

        // Access underlying XSSFSheet for chart support
        XSSFSheet xssfSheet = SXSSFSheetHelper.getXSSFSheet(sheet);

        // Below the data and any summary/afterData rows, leaving one blank row.
        int row1 = anchorRow1 >= 0 ? anchorRow1 : Math.max(dataEndRow + 1, nextFreeRow) + 1;
        int row2 = anchorRow2 >= 0 ? anchorRow2 : row1 + 15;

        XSSFDrawing drawing = xssfSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                anchorCol1, row1, anchorCol2, row2);
        XSSFChart chart = drawing.createChart(anchor);

        if (title != null) {
            chart.setTitleText(title);
        }

        int dataStartRow = headerRow + 1;

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                xssfSheet,
                new CellRangeAddress(
                        dataStartRow, dataEndRow,
                        categoryColumnIndex, categoryColumnIndex));

        switch (chartType) {
            case BAR -> createBarChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case LINE -> createLineChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case PIE -> createPieChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case SCATTER -> createScatterChart(chart, xssfSheet, dataStartRow, dataEndRow);
            case AREA -> createAreaChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
            case DOUGHNUT -> createDoughnutChart(chart, categories, xssfSheet, dataStartRow, dataEndRow);
        }

        applyLegend(chart);
        applyDataLabels(chart);
    }

    private void createBarChart(XSSFChart chart, XDDFDataSource<String> categories,
                                XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // The columns sit between the tick marks: with the default (on them) the first and the last one are cut in half.
        valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        applyAxisTitles(categoryAxis, valueAxis);

        XDDFBarChartData data = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, categoryAxis, valueAxis);
        data.setBarDirection(barDirection == BarDirection.HORIZONTAL
                ? org.apache.poi.xddf.usermodel.chart.BarDirection.BAR
                : org.apache.poi.xddf.usermodel.chart.BarDirection.COL);
        applyBarGrouping(data);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
        addLineSeries(chart, categoryAxis, valueAxis, categories, sheet, dataStart, dataEnd);
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
        addLineSeries(chart, categoryAxis, valueAxis, categories, sheet, dataStart, dataEnd);
    }

    private void createPieChart(XSSFChart chart, XDDFDataSource<String> categories,
                                XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    private void createScatterChart(XSSFChart chart, XSSFSheet sheet,
                                    int dataStart, int dataEnd) {
        XDDFValueAxis xAxis = chart.createValueAxis(AxisPosition.BOTTOM);
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        if (categoryAxisTitle != null) xAxis.setTitle(categoryAxisTitle);
        if (valueAxisTitle != null) yAxis.setTitle(valueAxisTitle);

        XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(
                ChartTypes.SCATTER, xAxis, yAxis);

        // For scatter, X-axis is numeric
        XDDFNumericalDataSource<Double> xValues = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet,
                new CellRangeAddress(
                        dataStart, dataEnd,
                        categoryColumnIndex, categoryColumnIndex));

        for (ValueSeries vs : valueSeries) {
            XDDFNumericalDataSource<Double> yValues = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet,
                    new CellRangeAddress(
                            dataStart, dataEnd,
                            vs.columnIndex, vs.columnIndex));
            XDDFChartData.Series series = data.addSeries(xValues, yValues);
            series.setTitle(vs.title, null);
        }

        chart.plot(data);
    }

    private void createAreaChart(XSSFChart chart, XDDFDataSource<String> categories,
                                 XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        applyAxisTitles(categoryAxis, valueAxis);

        XDDFAreaChartData data = (XDDFAreaChartData) chart.createData(
                ChartTypes.AREA, categoryAxis, valueAxis);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
        addLineSeries(chart, categoryAxis, valueAxis, categories, sheet, dataStart, dataEnd);
    }

    private void createDoughnutChart(XSSFChart chart, XDDFDataSource<String> categories,
                                     XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartData data = chart.createData(ChartTypes.DOUGHNUT, null, null);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    /**
     * Draws the line series over the main chart: those on the primary axis share the main category and value axes; those
     * on the secondary axis get their own value axis on the right (crossing at its maximum) and a category axis that is
     * not shown, which is how Excel stores a chart with a second axis.
     */
    private void addLineSeries(XSSFChart chart, XDDFChartAxis categoryAxis, XDDFValueAxis valueAxis,
                               XDDFDataSource<String> categories, XSSFSheet sheet, int dataStart, int dataEnd) {
        List<ValueSeries> primary = lineSeries.stream().filter(s -> !s.secondaryAxis()).toList();
        List<ValueSeries> secondary = lineSeries.stream().filter(ValueSeries::secondaryAxis).toList();
        if (!primary.isEmpty()) {
            XDDFLineChartData lines = (XDDFLineChartData) chart.createData(ChartTypes.LINE, categoryAxis, valueAxis);
            addSeries(lines, primary, categories, sheet, dataStart, dataEnd);
            chart.plot(lines);
        }
        if (!secondary.isEmpty()) {
            XDDFCategoryAxis secondCategoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            secondCategoryAxis.setVisible(false);
            XDDFValueAxis secondValueAxis = chart.createValueAxis(AxisPosition.RIGHT);
            secondValueAxis.setCrosses(AxisCrosses.MAX);
            // like the main axis, so that the points of a line fall in the middle of the columns
            secondValueAxis.setCrossBetween(chartType == ChartType.BAR ? AxisCrossBetween.BETWEEN : AxisCrossBetween.MIDPOINT_CATEGORY);
            if (secondaryValueAxisTitle != null) secondValueAxis.setTitle(secondaryValueAxisTitle);
            secondCategoryAxis.crossAxis(secondValueAxis);
            secondValueAxis.crossAxis(secondCategoryAxis);
            XDDFLineChartData lines = (XDDFLineChartData) chart.createData(ChartTypes.LINE, secondCategoryAxis, secondValueAxis);
            addSeries(lines, secondary, categories, sheet, dataStart, dataEnd);
            chart.plot(lines);
        }
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

            // Apply data labels to scatter charts
            for (var scatterChart : plotArea.getScatterChartList()) {
                var dLbls = scatterChart.isSetDLbls() ? scatterChart.getDLbls() : scatterChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }

            // Apply data labels to area charts
            for (var areaChart : plotArea.getAreaChartList()) {
                var dLbls = areaChart.isSetDLbls() ? areaChart.getDLbls() : areaChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }

            // Apply data labels to doughnut charts
            for (var doughnutChart : plotArea.getDoughnutChartList()) {
                var dLbls = doughnutChart.isSetDLbls() ? doughnutChart.getDLbls() : doughnutChart.addNewDLbls();
                dLbls.addNewShowVal().setVal(true);
                dLbls.addNewShowCatName().setVal(false);
                dLbls.addNewShowSerName().setVal(false);
            }
        } catch (RuntimeException e) {
            throw new ExcelWriteException("Failed to apply chart data labels", e);
        }
    }

    private void addSeries(XDDFChartData chartData, XDDFDataSource<String> categories,
                           XSSFSheet sheet, int dataStart, int dataEnd) {
        addSeries(chartData, valueSeries, categories, sheet, dataStart, dataEnd);
    }

    private void addSeries(XDDFChartData chartData, List<ValueSeries> seriesList, XDDFDataSource<String> categories,
                           XSSFSheet sheet, int dataStart, int dataEnd) {
        for (ValueSeries vs : seriesList) {
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet,
                    new CellRangeAddress(
                            dataStart, dataEnd,
                            vs.columnIndex, vs.columnIndex));
            XDDFChartData.Series series = chartData.addSeries(categories, values);
            series.setTitle(vs.title, null);
            // A smoothed line bends past the points (the growth of a month shows below its real value).
            if (series instanceof XDDFLineChartData.Series line) line.setSmooth(false);
        }
    }

    private record ValueSeries(int columnIndex, String title, boolean secondaryAxis) {
    }
}
