package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A chart is placed below everything the sheet holds: several summary rows and content written by
 * {@code afterData}, on the last sheet and, with rollover, on every earlier sheet too. Both writers, with and
 * without rollover.
 */
class ChartPlacementTest {

    private record Item(String name, int qty) {
    }

    private static final int ROWS = 10;

    private static void checkEverySheet(byte[] xlsx, int expectedSheets, boolean afterData, String at) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            assertEquals(expectedSheets, wb.getNumberOfSheets(), at);
            for (int s = 0; s < expectedSheets; s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                assertNotNull(sheet.getDrawingPatriarch(), at + " sheet " + s + " has no chart");
                int top = sheet.getDrawingPatriarch().getCTDrawing().getTwoCellAnchorArray(0).getFrom().getRow();
                assertTrue(top > sheet.getLastRowNum(),
                        at + " sheet " + s + ": the chart starts at row " + top + " but content reaches row " + sheet.getLastRowNum());
                // the summary really produced several rows, so the check above means something
                boolean marker = false;
                for (var row : sheet) {
                    var c = row.getCell(0);
                    if (c != null && c.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                            && c.getStringCellValue().equals("AFTER")) marker = true;
                }
                assertEquals(afterData, marker, at + " sheet " + s + " afterData content");
                int dataRows = 0;
                for (var row : sheet) {
                    var c = row.getCell(0);
                    if (c != null && c.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING
                            && c.getStringCellValue().startsWith("item ")) dataRows++;
                }
                assertTrue(sheet.getLastRowNum() >= dataRows + 3, at + " sheet " + s + " is too short to show several summary rows");
            }
        }
    }

    @ParameterizedTest(name = "ExcelWriter maxRows={0} afterData={1}")
    @CsvSource({"0,false", "0,true", "4,false", "4,true", "5,true"})
    void excelWriterPlacesTheChartBelowEverything(int maxRows, boolean afterData) throws Exception {
        ExcelWriter<Item> writer = ExcelWriter.<Item>create()
                .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Qty").average("Qty").max("Qty"))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"));
        if (maxRows > 0) writer.maxRows(maxRows);
        if (afterData) writer.afterData(ctx -> {
            ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("AFTER");
            return ctx.getCurrentRow() + 2;
        });
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(IntStream.range(0, ROWS).mapToObj(i -> new Item("item " + i, i)).toList()).writeTo(out);

        checkEverySheet(out.toByteArray(), maxRows > 0 ? (int) Math.ceil(ROWS / (double) maxRows) : 1, afterData,
                "ExcelWriter maxRows=" + maxRows + " afterData=" + afterData);
    }

    @ParameterizedTest(name = "ExcelSheetWriter maxRows={0} afterData={1}")
    @CsvSource({"0,false", "0,true", "4,false", "4,true", "5,true"})
    void sheetWriterPlacesTheChartBelowEverything(int maxRows, boolean afterData) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            ExcelSheetWriter<Item> writer = book.<Item>sheet("Data")
                    .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Qty").average("Qty").max("Qty"))
                    .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"));
            if (maxRows > 0) writer.maxRows(maxRows);
            if (afterData) writer.afterData(ctx -> {
                ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("AFTER");
                return ctx.getCurrentRow() + 2;
            });
            writer.write(IntStream.range(0, ROWS).mapToObj(i -> new Item("item " + i, i)));
            book.finish().writeTo(out);
        }
        checkEverySheet(out.toByteArray(), maxRows > 0 ? (int) Math.ceil(ROWS / (double) maxRows) : 1, afterData,
                "ExcelSheetWriter maxRows=" + maxRows + " afterData=" + afterData);
    }
}
