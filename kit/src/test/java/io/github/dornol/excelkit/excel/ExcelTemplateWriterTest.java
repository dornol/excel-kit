package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelTemplateWriter} and {@link TemplateListWriter}.
 */
class ExcelTemplateWriterTest {

    /**
     * Creates a minimal template in memory and returns it as an InputStream.
     */
    private InputStream createTemplate() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Sheet1");
            // Row 0: title
            sheet.createRow(0).createCell(0).setCellValue("Report Title");
            // Row 1: empty
            // Row 2: labels
            XSSFRow labelRow = sheet.createRow(2);
            labelRow.createCell(0).setCellValue("Client:");
            labelRow.createCell(1).setCellValue(""); // placeholder
            // Row 3: date label
            XSSFRow dateRow = sheet.createRow(3);
            dateRow.createCell(0).setCellValue("Date:");
            // Row 4: column headers
            XSSFRow headerRow = sheet.createRow(4);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Qty");
            headerRow.createCell(2).setCellValue("Amount");

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return new ByteArrayInputStream(bos.toByteArray());
        }
    }

    private InputStream createMultiSheetTemplate() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Users").createRow(0).createCell(0).setCellValue("Name");
            wb.createSheet("Orders").createRow(0).createCell(0).setCellValue("OrderID");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return new ByteArrayInputStream(bos.toByteArray());
        }
    }

    private XSSFWorkbook readOutput(ByteArrayOutputStream bos) throws IOException {
        return new XSSFWorkbook(new ByteArrayInputStream(bos.toByteArray()));
    }

    // ============================================================
    // Cell-level writes
    // ============================================================
    @Nested
    class CellWriteTests {

        @Test
        void cell_writesStringValue() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", "Acme Corp").finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("Acme Corp", wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
            }
        }

        @Test
        void cell_writesNumberValue() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("C3", 42).finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals(42.0, wb.getSheetAt(0).getRow(2).getCell(2).getNumericCellValue());
            }
        }

        @Test
        void cell_writesBooleanValue() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("C3", true).finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertTrue(wb.getSheetAt(0).getRow(2).getCell(2).getBooleanCellValue());
            }
        }

        @Test
        void cell_writesLocalDateValue() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B4", LocalDate.of(2026, 3, 19)).finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFCell cell = wb.getSheetAt(0).getRow(3).getCell(1);
                assertNotNull(cell.getLocalDateTimeCellValue());
            }
        }

        @Test
        void cell_writesLocalDateTimeValue() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B4", LocalDateTime.of(2026, 3, 19, 14, 30))
                        .finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertNotNull(wb.getSheetAt(0).getRow(3).getCell(1));
            }
        }

        @Test
        void cell_writesNullAsBlank() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", null).finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFCell cell = wb.getSheetAt(0).getRow(2).getCell(1);
                assertEquals(CellType.BLANK, cell.getCellType());
            }
        }

        @Test
        void cell_byRowCol_works() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell(2, 1, "TestValue").finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("TestValue", wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
            }
        }

        @Test
        void cell_multipleCells_topDown() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", "Client A")
                        .cell("B4", "2026-03-19")
                        .finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("Client A", wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
                assertEquals("2026-03-19", wb.getSheetAt(0).getRow(3).getCell(1).getStringCellValue());
            }
        }
    }

    // ============================================================
    // Row-order enforcement
    // ============================================================
    @Nested
    class RowOrderTests {

        @Test
        void cell_reverseOrder_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B4", "first");
                assertThrows(ExcelWriteException.class, () -> w.cell("B3", "second"));
            }
        }

        @Test
        void cell_sameRow_allowed() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("A3", "col1")
                        .cell("B3", "col2")
                        .finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("col1", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
                assertEquals("col2", wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
            }
        }
    }

    // ============================================================
    // List streaming
    // ============================================================
    @Nested
    class ListWriteTests {

        @Test
        void list_writesDataFromStartRow() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .column("Qty", s -> s.length())
                        .column("Amount", s -> s.length() * 100)
                        .write(Stream.of("Widget", "Gadget"));
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("Widget", sheet.getRow(5).getCell(0).getStringCellValue());
                assertEquals("Gadget", sheet.getRow(6).getCell(0).getStringCellValue());
            }
        }

        @Test
        void list_withHeaders_writesHeadersThenData() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Product", s -> s)
                        .column("Count", s -> s.length())
                        .writeWithHeaders(Stream.of("A", "B"));
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("Product", sheet.getRow(5).getCell(0).getStringCellValue());
                assertEquals("Count", sheet.getRow(5).getCell(1).getStringCellValue());
                assertEquals("A", sheet.getRow(6).getCell(0).getStringCellValue());
            }
        }

        @Test
        void list_afterData_callback() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .afterData(ctx -> {
                            ctx.getSheet().createRow(ctx.getCurrentRow())
                                    .createCell(0).setCellValue("Total: 2");
                            return ctx.getCurrentRow() + 1;
                        })
                        .write(Stream.of("A", "B"));
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("Total: 2", sheet.getRow(7).getCell(0).getStringCellValue());
            }
        }

        @Test
        void list_noColumns_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                TemplateListWriter<String> lw = w.<String>list(5);
                assertThrows(ExcelWriteException.class, () -> lw.write(Stream.of("x")));
            }
        }

        @Test
        void list_largeDataset() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate(), 100)) {
                w.<Integer>list(5)
                        .column("ID", i -> i)
                        .column("Value", i -> i * 10)
                        .write(IntStream.range(0, 10_000).boxed());
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertNotNull(wb.getSheetAt(0).getRow(10004));
            }
        }
    }

    // ============================================================
    // Mixed mode (cell + list)
    // ============================================================
    @Nested
    class MixedModeTests {

        @Test
        void cellThenList_works() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", "Acme Corp")
                        .cell("B4", "2026-03-19");

                w.<String>list(5)
                        .column("Name", s -> s)
                        .column("Qty", s -> s.length())
                        .write(Stream.of("ItemA", "ItemB"));

                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                // Template preserved
                assertEquals("Report Title", sheet.getRow(0).getCell(0).getStringCellValue());
                // Cell writes
                assertEquals("Acme Corp", sheet.getRow(2).getCell(1).getStringCellValue());
                // List data
                assertEquals("ItemA", sheet.getRow(5).getCell(0).getStringCellValue());
                assertEquals("ItemB", sheet.getRow(6).getCell(0).getStringCellValue());
            }
        }

        @Test
        void cellThenListThenCell_works() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", "Client");

                w.<String>list(5)
                        .column("Name", s -> s)
                        .write(Stream.of("X", "Y"));

                // Write cell after list (row 10, well after data ends at row 7)
                w.cell("A10", "Footer note");

                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("Client", sheet.getRow(2).getCell(1).getStringCellValue());
                assertEquals("X", sheet.getRow(5).getCell(0).getStringCellValue());
                assertEquals("Footer note", sheet.getRow(9).getCell(0).getStringCellValue());
            }
        }
    }

    // ============================================================
    // Multi-sheet
    // ============================================================
    @Nested
    class MultiSheetTests {

        @Test
        void sheet_switchAndWrite() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createMultiSheetTemplate())) {
                w.sheet(0).cell("A2", "Alice");
                w.sheet(1).cell("A2", "ORD-001");
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("ORD-001", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void sheet_byName() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createMultiSheetTemplate())) {
                w.sheet("Orders").cell("A2", "ORD-100");
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("ORD-100", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void sheet_invalidIndex_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createMultiSheetTemplate())) {
                assertThrows(ExcelWriteException.class, () -> w.sheet(5));
            }
        }

        @Test
        void sheet_invalidName_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createMultiSheetTemplate())) {
                assertThrows(ExcelWriteException.class, () -> w.sheet("NoSuchSheet"));
            }
        }
    }

    // ============================================================
    // Template preservation
    // ============================================================
    @Nested
    class TemplatePreservationTests {

        @Test
        void existingContent_preserved() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", "NewClient").finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                // Original template content still there
                assertEquals("Report Title", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Client:", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("Date:", sheet.getRow(3).getCell(0).getStringCellValue());
                // Headers from template
                assertEquals("Name", sheet.getRow(4).getCell(0).getStringCellValue());
                assertEquals("Qty", sheet.getRow(4).getCell(1).getStringCellValue());
                assertEquals("Amount", sheet.getRow(4).getCell(2).getStringCellValue());
            }
        }

        @Test
        void mergedRegions_preserved() throws IOException {
            // Create template with merged region
            ByteArrayOutputStream templateBos = new ByteArrayOutputStream();
            try (XSSFWorkbook twb = new XSSFWorkbook()) {
                XSSFSheet sheet = twb.createSheet("Sheet1");
                sheet.createRow(0).createCell(0).setCellValue("Merged Title");
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
                sheet.createRow(2).createCell(0).setCellValue("Data:");
                twb.write(templateBos);
            }

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(
                    new ByteArrayInputStream(templateBos.toByteArray()))) {
                w.cell("A3", "value").finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals(1, wb.getSheetAt(0).getNumMergedRegions());
            }
        }
    }

    // ============================================================
    // TemplateListWriter options
    // ============================================================
    @Nested
    class ListOptionsTests {

        @Test
        void list_rowHeight_doesNotThrow() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .rowHeight(25)
                        .write(Stream.of("A"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }

        @Test
        void list_rowColor_appliesColor() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .rowColor(s -> "error".equals(s) ? ExcelColor.LIGHT_RED : null)
                        .write(Stream.of("ok", "error"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }

        @Test
        void list_summary_appendsFormulaRow() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<Integer>list(5)
                        .column("Name", i -> "Item" + i)
                        .column("Qty", i -> i, c -> c.type(ExcelDataType.INTEGER))
                        .summary(s -> s.label("Total").sum("Qty"))
                        .write(Stream.of(10, 20, 30));
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                // Summary row should be at row 8 (5 + 3 data rows)
                XSSFRow summaryRow = sheet.getRow(8);
                assertNotNull(summaryRow);
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            }
        }

        @Test
        void list_defaultStyle_appliedToColumns() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .defaultStyle(ds -> ds.bold(true))
                        .write(Stream.of("Bold"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }

        @Test
        void list_onProgress_callbackFires() throws IOException {
            int[] count = {0};
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<Integer>list(5)
                        .column("ID", i -> i)
                        .onProgress(10, (total, cursor) -> count[0]++)
                        .write(IntStream.range(0, 50).boxed());
                w.finish();
            }
            assertEquals(5, count[0]); // 50 rows / 10 interval
        }

        @Test
        void list_onProgress_invalidInterval_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                TemplateListWriter<String> lw = w.<String>list(5).column("X", s -> s);
                assertThrows(IllegalArgumentException.class, () -> lw.onProgress(0, (t, c) -> {}));
            }
        }

        @Test
        void list_autoWidthSampleRows_accepted() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .autoWidthSampleRows(50)
                        .write(Stream.of("A"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }

        @Test
        void list_autoWidthSampleRows_negative_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                TemplateListWriter<String> lw = w.<String>list(5).column("X", s -> s);
                assertThrows(IllegalArgumentException.class, () -> lw.autoWidthSampleRows(-1));
            }
        }

        @Test
        void list_columnWithCursorFunction() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("No", (ExcelRowFunction<String, Object>) (s, cursor) -> cursor.getCurrentTotal())
                        .column("Name", s -> s)
                        .write(Stream.of("A", "B", "C"));
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                // Row numbers: 1, 2, 3
                assertNotNull(wb.getSheetAt(0).getRow(5));
            }
        }

        @Test
        void list_columnWithConfig() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s, c -> c.bold(true).fontSize(14))
                        .write(Stream.of("Styled"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }

        @Test
        void list_cursorColumnWithConfig() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("No", (ExcelRowFunction<String, Object>) (s, c) -> c.getCurrentTotal(),
                                cfg -> cfg.type(ExcelDataType.INTEGER))
                        .write(Stream.of("A"));
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    // ============================================================
    // Edge cases
    // ============================================================
    @Nested
    class EdgeCaseTests {

        @Test
        void cell_objectValue_fallsBackToString() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("B3", new Object() {
                    @Override
                    public String toString() {
                        return "custom-object";
                    }
                }).finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("custom-object", wb.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
            }
        }

        @Test
        void list_emptyStream_writesNothing() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.<String>list(5)
                        .column("Name", s -> s)
                        .write(Stream.empty());
                w.finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                // Row 5 should not exist (no data written)
                assertNull(wb.getSheetAt(0).getRow(5));
            }
        }

        @Test
        void cell_writesToNewRowBeyondTemplate() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.cell("A100", "far away").finish().consumeOutputStream(bos);
            }
            try (XSSFWorkbook wb = readOutput(bos)) {
                assertEquals("far away", wb.getSheetAt(0).getRow(99).getCell(0).getStringCellValue());
            }
        }

        @Test
        void list_afterFinish_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.finish();
                assertThrows(ExcelWriteException.class, () -> w.list(5));
            }
        }
    }

    // ============================================================
    // Lifecycle
    // ============================================================
    @Nested
    class LifecycleTests {

        @Test
        void finish_twice_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.finish();
                assertThrows(ExcelWriteException.class, w::finish);
            }
        }

        @Test
        void close_withoutFinish_doesNotThrow() throws IOException {
            ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate());
            assertDoesNotThrow(w::close);
        }

        @Test
        void cell_afterFinish_throws() throws IOException {
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.finish();
                assertThrows(ExcelWriteException.class, () -> w.cell("A1", "x"));
            }
        }

        @Test
        void finish_producesNonEmptyOutput() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelTemplateWriter w = new ExcelTemplateWriter(createTemplate())) {
                w.finish().consumeOutputStream(bos);
            }
            assertTrue(bos.toByteArray().length > 0);
        }
    }
}
