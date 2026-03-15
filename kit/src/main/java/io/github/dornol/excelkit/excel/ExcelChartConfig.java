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

    private ChartType chartType = ChartType.BAR;
    private String title;
    private int categoryColumnIndex = 0;
    private final List<ValueSeries> valueSeries = new ArrayList<>();
    private int anchorCol1 = 0;
    private int anchorRow1 = -1;
    private int anchorCol2 = 8;
    private int anchorRow2 = -1;

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
    }

    private void createBarChart(XSSFChart chart, XDDFDataSource<String> categories,
                                XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFBarChartData data = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, categoryAxis, valueAxis);
        data.setBarDirection(BarDirection.COL);

        addSeries(data, categories, sheet, dataStart, dataEnd);
        chart.plot(data);
    }

    private void createLineChart(XSSFChart chart, XDDFDataSource<String> categories,
                                 XSSFSheet sheet, int dataStart, int dataEnd) {
        XDDFChartAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

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
