package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CellMergeTest {

    @Test
    void mergeCells_indexBased_viaBeforeHeader() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .beforeHeader(ctx -> {
                    // Create a title row spanning columns A-C (indices 0-2)
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Report Title");
                    ctx.mergeCells(0, 0, 0, 2);
                    return ctx.getCurrentRow() + 1;
                })
                .column("Name", (row, c) -> row)
                .column("Age", (row, c) -> row.length())
                .column("City", (row, c) -> row.toUpperCase())
                .write(Stream.of("Alice", "Bob"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1, sheet.getNumMergedRegions(), "Should have exactly one merged region");
            var region = sheet.getMergedRegion(0);
            assertEquals(0, region.getFirstRow());
            assertEquals(0, region.getLastRow());
            assertEquals(0, region.getFirstColumn());
            assertEquals(2, region.getLastColumn());
        }
    }

    @Test
    void mergeCells_stringBased_viaAfterData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Name", (row, c) -> row)
                .column("Value", (row, c) -> row.length())
                .afterData(ctx -> {
                    int row = ctx.getCurrentRow();
                    ctx.getSheet().createRow(row).createCell(0).setCellValue("Total");
                    // Merge using Excel notation — row numbers in Excel are 1-based
                    String range = "A" + (row + 1) + ":B" + (row + 1);
                    ctx.mergeCells(range);
                    return row + 1;
                })
                .write(Stream.of("Alice", "Bob"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1, sheet.getNumMergedRegions(), "Should have exactly one merged region");
            // afterData row: header(0) + 2 data rows = row index 3 (Excel row 4)
            var region = sheet.getMergedRegion(0);
            assertEquals(3, region.getFirstRow());
            assertEquals(3, region.getLastRow());
            assertEquals(0, region.getFirstColumn());
            assertEquals(1, region.getLastColumn());
        }
    }

    @Test
    void mergeCells_shouldSupportChaining() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(0).createCell(0).setCellValue("Title");
                    ctx.getSheet().createRow(1).createCell(0).setCellValue("Subtitle");
                    ctx.mergeCells(0, 0, 0, 2)
                       .mergeCells(1, 1, 0, 1);
                    return ctx.getCurrentRow() + 2;
                })
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row)
                .column("C", (row, c) -> row)
                .write(Stream.of("x"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(2, sheet.getNumMergedRegions(), "Chained mergeCells should create two merged regions");
        }
    }

    @Test
    void mergeCells_viaAfterAll() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Name", (row, c) -> row)
                .column("Value", (row, c) -> row.length())
                .afterAll(ctx -> {
                    int row = ctx.getCurrentRow();
                    ctx.getSheet().createRow(row).createCell(0).setCellValue("Grand Total");
                    ctx.mergeCells(row, row, 0, 1);
                    return row + 1;
                })
                .write(Stream.of("Alice", "Bob"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1, sheet.getNumMergedRegions(), "afterAll callback should create one merged region");
            var region = sheet.getMergedRegion(0);
            assertEquals(3, region.getFirstRow());
            assertEquals(3, region.getLastRow());
            assertEquals(0, region.getFirstColumn());
            assertEquals(1, region.getLastColumn());
        }
    }

    @Test
    void mergeCells_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = new ExcelWorkbook()) {
            workbook.<String>sheet("TestSheet")
                    .beforeHeader(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Report");
                        ctx.mergeCells(0, 0, 0, 2);
                        return ctx.getCurrentRow() + 1;
                    })
                    .column("Name", (row, c) -> row)
                    .column("Age", (row, c) -> row.length())
                    .column("City", (row, c) -> row.toUpperCase())
                    .afterData(ctx -> {
                        int row = ctx.getCurrentRow();
                        ctx.getSheet().createRow(row).createCell(0).setCellValue("Footer");
                        ctx.mergeCells(row, row, 0, 2);
                        return row + 1;
                    })
                    .write(Stream.of("Alice", "Bob"));
            workbook.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(2, sheet.getNumMergedRegions(), "Should have two merged regions (beforeHeader + afterData)");
            // First merged region: beforeHeader title row
            var region0 = sheet.getMergedRegion(0);
            assertEquals(0, region0.getFirstRow());
            assertEquals(0, region0.getLastRow());
            assertEquals(0, region0.getFirstColumn());
            assertEquals(2, region0.getLastColumn());
            // Second merged region: afterData footer row
            var region1 = sheet.getMergedRegion(1);
            assertEquals(0, region1.getFirstColumn());
            assertEquals(2, region1.getLastColumn());
        }
    }

    @Test
    void mergeCells_preservesCellValue() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .beforeHeader(ctx -> {
                    var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("Merged Title");
                    row.createCell(1).setCellValue("");
                    row.createCell(2).setCellValue("");
                    ctx.mergeCells(0, 0, 0, 2);
                    return ctx.getCurrentRow() + 1;
                })
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row)
                .column("C", (row, c) -> row)
                .write(Stream.of("x"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals(1, sheet.getNumMergedRegions());
            // The first cell of the merged region should preserve its value
            String cellValue = sheet.getRow(0).getCell(0).getStringCellValue();
            assertEquals("Merged Title", cellValue, "First cell of merged region should preserve its value");
        }
    }

    @Test
    void mergeCells_withSheetRollover() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>(3)
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Title");
                    ctx.mergeCells(ctx.getCurrentRow(), ctx.getCurrentRow(), 0, 1);
                    return ctx.getCurrentRow() + 1;
                })
                .column("Name", (row, c) -> row)
                .column("Value", (row, c) -> row.length())
                .write(IntStream.rangeClosed(1, 6).mapToObj(i -> "Item" + i))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getNumberOfSheets() >= 2, "Should have at least 2 sheets after rollover");
            // Both sheets should have a merged region from the beforeHeader callback
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                var sheet = wb.getSheetAt(i);
                assertTrue(sheet.getNumMergedRegions() >= 1,
                        "Sheet " + i + " should have at least one merged region");
                var region = sheet.getMergedRegion(0);
                assertEquals(0, region.getFirstColumn());
                assertEquals(1, region.getLastColumn());
            }
        }
    }
}
