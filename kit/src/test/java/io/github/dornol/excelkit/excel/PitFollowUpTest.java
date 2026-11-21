package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Edges that mutation testing found unpinned in code written during the audits. */
class PitFollowUpTest {

    record Row(String name, int value) {
    }

    @Test
    void chartPositionAcceptsEqualEdgesAndRefusesOnlyWhatRunsBackwards() {
        ExcelChartConfig c = new ExcelChartConfig();
        assertDoesNotThrow(() -> c.position(0, 5, 0, 5), "a one-cell chart is not backwards");
        assertDoesNotThrow(() -> c.position(3, 5, 3, 6));
        assertDoesNotThrow(() -> c.position(0, -1, 8, 4), "row1 negative: below the data");
        assertDoesNotThrow(() -> c.position(0, 4, 8, -1), "row2 negative: derived from row1");
        assertThrows(IllegalArgumentException.class, () -> c.position(0, 5, 8, 4), "one row up is backwards");
        assertThrows(IllegalArgumentException.class, () -> c.position(4, 0, 3, 5), "one column left is backwards");
    }

    @Test
    void percentStackedBarsAreWrittenAsPercentStacked() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().column("n", Row::name).column("v", Row::value)
                .chart(c -> c.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "V")
                        .barGrouping(ExcelChartConfig.BarGrouping.PERCENT_STACKED))
                .write(List.of(new Row("a", 1), new Row("b", 2))).writeTo(out);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
            assertTrue(chart.getCTChart().xmlText().contains("percentStacked"), chart.getCTChart().xmlText());
        }
    }

    @Test
    void csvDelimiterRefusalNamesTheCharacter() {
        assertTrue(assertThrows(IllegalArgumentException.class, () -> CsvWriter.<Row>create().delimiter('"'))
                .getMessage().contains("got '\"'"));
        assertTrue(assertThrows(IllegalArgumentException.class, () -> CsvWriter.<Row>create().delimiter('\n'))
                .getMessage().contains("got U+000A"));
        assertTrue(assertThrows(IllegalArgumentException.class, () -> CsvWriter.<Row>create().delimiter('\0'))
                .getMessage().contains("got U+0000"));
    }

    @Test
    void richTextHandedToTheBareCellWriterIsItsPlainText() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var cell = wb.createSheet("S").createRow(0).createCell(0);
            ExcelTemplateWriter.setCellValue(cell, new ExcelRichText().text("a").bold("b"));
            assertEquals("ab", cell.getStringCellValue());
        }
    }
}
