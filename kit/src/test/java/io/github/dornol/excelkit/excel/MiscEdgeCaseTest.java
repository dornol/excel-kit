package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ExcelKitException;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Miscellaneous edge case tests for remaining uncovered branches:
 * - ExcelHandler (consume twice, password null)
 * - ExcelReader (skipColumns negative, header not found)
 * - ExcelSummary (all Op types)
 * - ExcelWorkbook edge cases
 * - ExcelWriter defaultStyle with applyDefaults
 */
class MiscEdgeCaseTest {

    record Item(String name, int value) {}

    // ============================================================
    // ExcelHandler edge cases
    // ============================================================
    @Nested
    class ExcelHandlerTests {

        @Test
        void consumeOutputStream_twice_throws() throws IOException {
            ByteArrayOutputStream out1 = new ByteArrayOutputStream();
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            handler.consumeOutputStream(out1);
            assertThrows(ExcelWriteException.class, () -> handler.consumeOutputStream(new ByteArrayOutputStream()));
        }

        @Test
        void consumeOutputStreamWithPassword_nullPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), (String) null));
        }

        @Test
        void consumeOutputStreamWithPassword_blankPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), "  "));
        }

        @Test
        void consumeOutputStreamWithPassword_charArray_nullPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), (char[]) null));
        }

        @Test
        void consumeOutputStreamWithPassword_charArray_emptyPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), new char[0]));
        }
    }

    // ============================================================
    // ExcelReader edge cases
    // ============================================================
    @Nested
    class ExcelReaderTests {

        static class MutableItem {
            String name;
            int value;
        }

        @Test
        void skipColumns_negative_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            assertThrows(IllegalArgumentException.class, () -> reader.skipColumns(-1));
        }

        @Test
        void onProgress_zeroInterval_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            assertThrows(IllegalArgumentException.class,
                    () -> reader.onProgress(0, (c, cur) -> {}));
        }

        @Test
        void headerNotFound_throws() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            reader.addColumn("Name", (item, cell) -> {});
            reader.addColumn("NonExistentHeader", (item, cell) -> {});

            assertThrows(ExcelKitException.class,
                    () -> reader.build(new ByteArrayInputStream(out.toByteArray()))
                            .read(r -> {}));
        }

        @Test
        void getSheetHeaders_withHeaderRowIndex() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            List<String> headers = ExcelReader.getSheetHeaders(
                    new ByteArrayInputStream(out.toByteArray()), 0, 0);

            assertEquals(2, headers.size());
            assertEquals("Name", headers.get(0));
            assertEquals("Value", headers.get(1));
        }
    }

    // ============================================================
    // ExcelSummary all Op types
    // ============================================================
    @Nested
    class ExcelSummaryTests {

        @Test
        void allSummaryOps_shouldWriteFormulaRows() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Summary")
                            .sum("Value")
                            .average("Value")
                            .count("Value")
                            .min("Value")
                            .max("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Header(0) + 2 data rows(1,2) + 5 summary rows(3-7)
                // Check SUM row has formula
                var sumRow = sheet.getRow(3);
                assertNotNull(sumRow, "SUM summary row should exist");
                var sumCell = sumRow.getCell(1);
                assertEquals(CellType.FORMULA, sumCell.getCellType(), "Summary cell should be formula");
                assertTrue(sumCell.getCellFormula().startsWith("SUM("), "Should be SUM formula");
            }
        }

        @Test
        void summary_singleOp_withLabel_shouldUseLabelText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Total")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(3);
                // Single op with label → uses labelText "Total"
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            }
        }

        @Test
        void summary_labelInNonExistentColumn_shouldFallbackToFirstColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("NonExistent", "Total:")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(2);
                // NonExistent column → falls back to idx=0
                assertEquals("Total:", summaryRow.getCell(0).getStringCellValue());
            }
        }

        @Test
        void summary_labelInColumn_shouldWriteLabelAndFormula() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Name", "Total:")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                var summaryRow = sheet.getRow(3);
                assertNotNull(summaryRow);
                assertEquals("Total:", summaryRow.getCell(0).getStringCellValue());
                assertEquals(CellType.FORMULA, summaryRow.getCell(1).getCellType());
            }
        }
    }

    // ============================================================
    // ExcelWriter defaultStyle
    // ============================================================
    @Nested
    class ExcelWriterDefaultStyleTests {

        @Test
        void defaultStyle_shouldApplyBoldToAllColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .defaultStyle(d -> d.bold(true).fontSize(12))
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var dataRow = wb.getSheetAt(0).getRow(1);
                // Both columns should have bold font from defaultStyle
                var font0 = wb.getFontAt(dataRow.getCell(0).getCellStyle().getFontIndex());
                assertTrue(font0.getBold(), "Name column should be bold from default style");
                var font1 = wb.getFontAt(dataRow.getCell(1).getCellStyle().getFontIndex());
                assertTrue(font1.getBold(), "Value column should be bold from default style");
            }
        }

        @Test
        void defaultStyle_columnOverrides_shouldNotBeBold() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .defaultStyle(d -> d.bold(true))
                    .addColumn("Name", Item::name, c -> c.bold(false))
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var dataRow = wb.getSheetAt(0).getRow(1);
                // Name column overrides bold=false
                var font0 = wb.getFontAt(dataRow.getCell(0).getCellStyle().getFontIndex());
                assertFalse(font0.getBold(), "Name column should override bold to false");
                // Value column inherits bold=true from default
                var font1 = wb.getFontAt(dataRow.getCell(1).getCellStyle().getFontIndex());
                assertTrue(font1.getBold(), "Value column should inherit bold from default");
            }
        }
    }

    // ============================================================
    // ExcelWorkbook edge cases
    // ============================================================
    @Nested
    class ExcelWorkbookTests {

        @Test
        void protectWorkbook_shouldSetWorkbookProtection() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<Item>sheet("Data")
                        .column("Name", Item::name)
                        .write(Stream.of(new Item("A", 1)));
                wb.protectWorkbook("password123");
                wb.finish().consumeOutputStream(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getCTWorkbook().isSetWorkbookProtection(),
                        "Workbook protection should be set");
            }
        }
    }

    // ============================================================
    // ExcelWriter header font customization
    // ============================================================
    @Nested
    class ExcelWriterHeaderFontTests {

        @Test
        void headerFontName_shouldApplyCustomFont() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .headerFontName("Arial")
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Arial", font.getFontName());
            }
        }

        @Test
        void headerFontSize_shouldApplyCustomSize() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .headerFontSize(16)
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals(16, font.getFontHeightInPoints());
            }
        }

        @Test
        void headerFontNameAndSize_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .headerFontName("Times New Roman")
                    .headerFontSize(14)
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Times New Roman", font.getFontName());
                assertEquals(14, font.getFontHeightInPoints());
            }
        }
    }

    // ============================================================
    // ExcelWriter write with no data
    // ============================================================
    @Nested
    class EmptyDataTests {

        @Test
        void write_emptyStream_shouldCreateHeaderOnly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.empty())
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Header row exists
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Value", sheet.getRow(0).getCell(1).getStringCellValue());
                // No data rows
                assertNull(sheet.getRow(1), "Should have no data rows");
            }
        }
    }

    // ============================================================
    // readStrict with error messages
    // ============================================================
    @Nested
    class ReadStrictTests {

        @Test
        void readStrict_emptyMessages_shouldShowUnknownError() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 10)))
                    .consumeOutputStream(out);

            // Read with a mapper that always succeeds
            List<Item> results = new ArrayList<>();
            ExcelReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).build(new ByteArrayInputStream(out.toByteArray()))
                    .readStrict(results::add);

            assertEquals(1, results.size());
        }
    }
}
