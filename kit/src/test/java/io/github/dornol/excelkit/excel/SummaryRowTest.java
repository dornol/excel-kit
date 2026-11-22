package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class SummaryRowTest {

    record Item(String name, int price, int qty) {}

    @Test
    void summary_singleSum_generatesFormulaAndLabel() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .column("Qty", i -> i.qty(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Price").sum("Qty"))
                .write(Stream.of(new Item("A", 100, 5), new Item("B", 200, 3)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: header, Row 1-2: data, Row 3: summary
            var summaryRow = sheet.getRow(3);
            assertNotNull(summaryRow, "Summary row should exist at row 3");
            assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
            assertEquals("SUM(C2:C3)", summaryRow.getCell(2).getCellFormula());
        }
    }

    @Test
    void summary_multipleColumnsSum_allFormulasGenerated() throws Exception {
        record Data(String a, int b, int c, int d) {}
        var handler = ExcelWriter.<Data>create()
                .column("A", Data::a)
                .column("B", x -> x.b(), c -> c.type(ExcelDataType.INTEGER))
                .column("C", x -> x.c(), c -> c.type(ExcelDataType.INTEGER))
                .column("D", x -> x.d(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Sum").sum("B").sum("C").sum("D"))
                .write(Stream.of(new Data("x", 1, 2, 3), new Data("y", 4, 5, 6)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var summaryRow = wb.getSheetAt(0).getRow(3);
            assertEquals("Sum", summaryRow.getCell(0).getStringCellValue());
            assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
            assertEquals("SUM(C2:C3)", summaryRow.getCell(2).getCellFormula());
            assertEquals("SUM(D2:D3)", summaryRow.getCell(3).getCellFormula());
        }
    }

    @Test
    void summary_multipleOps_separateRowsWithCorrectLabels() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.sum("Price").average("Price"))
                .write(Stream.of(new Item("A", 100, 5)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: header, Row 1: data, Row 2: SUM, Row 3: AVERAGE
            var sumRow = sheet.getRow(2);
            assertEquals("Sum", sumRow.getCell(0).getStringCellValue());
            assertEquals("SUM(B2:B2)", sumRow.getCell(1).getCellFormula());

            var avgRow = sheet.getRow(3);
            assertEquals("Average", avgRow.getCell(0).getStringCellValue());
            assertEquals("AVERAGE(B2:B2)", avgRow.getCell(1).getCellFormula());
        }
    }

    @Test
    void summary_countMinMax_generateCorrectFormulas() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.count("Price").min("Price").max("Price"))
                .write(Stream.of(new Item("A", 10, 1), new Item("B", 20, 2), new Item("C", 30, 3)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: header, Row 1-3: data, Row 4: COUNT, Row 5: MIN, Row 6: MAX
            var countRow = sheet.getRow(4);
            assertEquals("Count", countRow.getCell(0).getStringCellValue());
            assertEquals("COUNT(B2:B4)", countRow.getCell(1).getCellFormula());

            var minRow = sheet.getRow(5);
            assertEquals("Min", minRow.getCell(0).getStringCellValue());
            assertEquals("MIN(B2:B4)", minRow.getCell(1).getCellFormula());

            var maxRow = sheet.getRow(6);
            assertEquals("Max", maxRow.getCell(0).getStringCellValue());
            assertEquals("MAX(B2:B4)", maxRow.getCell(1).getCellFormula());
        }
    }

    @Test
    void summary_customLabelText_singleOp() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Grand Total").sum("Price"))
                .write(Stream.of(new Item("A", 100, 5)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var summaryRow = wb.getSheetAt(0).getRow(2);
            assertEquals("Grand Total", summaryRow.getCell(0).getStringCellValue(),
                    "Custom label should be used when only one op is configured");
        }
    }

    @Test
    void summary_labelInSpecificColumn_verifyPlacement() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .column("Qty", i -> i.qty(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Price", "Subtotal").sum("Qty"))
                .write(Stream.of(new Item("A", 100, 5)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var summaryRow = wb.getSheetAt(0).getRow(2);
            // Label should be in column index 1 (the "Price" column)
            assertEquals("Subtotal", summaryRow.getCell(1).getStringCellValue());
            assertEquals("SUM(C2:C2)", summaryRow.getCell(2).getCellFormula());
        }
    }

    @Test
    void excelSheetWriter_withSummary_generatesFormulaRow() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<Item>sheet("Sales")
                    .column("Name", Item::name)
                    .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Price"))
                    .write(Stream.of(new Item("X", 50, 1), new Item("Y", 75, 2)));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            var summaryRow = sheet.getRow(3);
            assertNotNull(summaryRow);
            assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            assertEquals("SUM(B2:B3)", summaryRow.getCell(1).getCellFormula());
        }
    }

    @Test
    void summary_withBeforeHeader_formulaRangeAccountsForOffset() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .beforeHeader(ctx -> {
                    // Write a title row before the header
                    var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("Report Title");
                    return ctx.getCurrentRow() + 1;
                })
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Price"))
                .write(Stream.of(new Item("A", 100, 5), new Item("B", 200, 3)));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: title, Row 1: header, Row 2-3: data, Row 4: summary
            var summaryRow = sheet.getRow(4);
            assertNotNull(summaryRow, "Summary row should be at row 4 after beforeHeader offset");
            assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            // Data starts at row 3 (1-based), ends at row 4 (1-based)
            assertEquals("SUM(B3:B4)", summaryRow.getCell(1).getCellFormula(),
                    "Formula range should account for the beforeHeader offset");
        }
    }

    @Test
    void summary_emptyData_summaryRowStillCreated() throws Exception {
        var handler = ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Price", i -> i.price(), c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Price"))
                .write(Stream.empty());

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: header, Row 1: summary (no data rows)
            var summaryRow = sheet.getRow(1);
            assertNotNull(summaryRow, "Summary row should exist even with no data");
            assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            // No data rows: no formula is written, since B2:B1 would include the header row
            assertNull(summaryRow.getCell(1));
        }
    }
}
