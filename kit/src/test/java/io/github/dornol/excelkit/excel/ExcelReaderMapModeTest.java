package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelReader#forMap()} — the v0.12.0 replacement for
 * the deleted {@code ExcelMapReader}. Verifies header auto-discovery,
 * fluent-API integration, and the mixed-mode runtime guards that prevent
 * calling setter-style column methods on a map-mode reader.
 */
class ExcelReaderMapModeTest {

    /**
     * Writes a simple two-column Excel file (Name/Age) to a byte array for test use.
     */
    private static byte[] writeSampleExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.forMap("Name", "Age", "City")
                .write(Stream.of(
                        Map.of("Name", "Alice", "Age", 30, "City", "Seoul"),
                        Map.of("Name", "Bob", "Age", 25, "City", "Tokyo")))
                .write(out);
        return out.toByteArray();
    }

    @Nested
    @DisplayName("Factory and return type")
    class FactoryAndType {

        @Test
        void forMap_returnsNonNull() {
            assertNotNull(ExcelReader.forMap());
        }

        @Test
        void forMap_returnsNewInstanceEachCall() {
            assertNotSame(ExcelReader.forMap(), ExcelReader.forMap());
        }
    }

    @Nested
    @DisplayName("Reading — header auto-discover")
    class HeaderAutoDiscover {

        @Test
        void forMap_readsAllColumnsFromHeaderRow() throws IOException {
            byte[] data = writeSampleExcel();
            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
            assertEquals("Seoul", results.get(0).get("City"));
            assertEquals("Bob", results.get(1).get("Name"));
            assertEquals("Tokyo", results.get(1).get("City"));
        }

        @Test
        void forMap_preservesHeaderOrder_LinkedHashMap() throws IOException {
            byte[] data = writeSampleExcel();
            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(List.of("Name", "Age", "City"),
                    new ArrayList<>(results.get(0).keySet()),
                    "map key order should match header-row column order");
        }
    }

    @Nested
    @DisplayName("Fluent API compatibility")
    class FluentApi {

        @Test
        void forMap_sheetIndex_readsSpecifiedSheet() throws IOException {
            // Write two separate files and then verify sheetIndex(0) reads the first.
            // Multi-sheet Excel writing needs ExcelWorkbook; for simplicity, verify sheetIndex(0) works.
            byte[] data = writeSampleExcel();
            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .sheetIndex(0)
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
        }

        @Test
        void forMap_headerRowIndex_skipsRowsBeforeHeader() throws IOException {
            // Build a file where headers are on row 2 (rowNum=2). Use ExcelWorkbook for custom placement.
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String[]>sheet("Sheet1")
                        .beforeHeader(ctx -> {
                            var sheet = ctx.getSheet();
                            var row0 = sheet.createRow(0);
                            row0.createCell(0).setCellValue("TITLE: Users");
                            var row1 = sheet.createRow(1);
                            row1.createCell(0).setCellValue("(subtitle)");
                            return 2;  // header goes on row 2
                        })
                        .column("Name", a -> a[0])
                        .column("Age", a -> a[1])
                        .write(Stream.of(new String[]{"Alice", "30"}, new String[]{"Bob", "25"}));
                wb.finish().write(out);
            }

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .headerRowIndex(2)
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("25", results.get(1).get("Age"));
        }

        @Test
        void forMap_onProgress_firesAtExpectedIntervals() throws IOException {
            // ExcelMapReader didn't expose onProgress; ExcelReader.forMap() does (inherited).
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.forMap("n")
                    .write(Stream.of(
                            Map.of("n", "1"), Map.of("n", "2"), Map.of("n", "3"),
                            Map.of("n", "4"), Map.of("n", "5"), Map.of("n", "6")))
                    .write(out);

            AtomicLong lastCount = new AtomicLong(0);
            AtomicInteger callCount = new AtomicInteger(0);
            ExcelReader.forMap()
                    .onProgress(2, (count, data) -> {
                        lastCount.set(count);
                        callCount.incrementAndGet();
                    })
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> {});

            assertEquals(3, callCount.get(), "onProgress fires at rows 2, 4, 6");
            assertEquals(6, lastCount.get());
        }

        @Test
        void forMap_readAsStream_producesSameRowsAsRead() throws IOException {
            byte[] data = writeSampleExcel();

            List<Map<String, String>> viaRead = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> viaRead.add(r.data()));

            List<Map<String, String>> viaStream = new ArrayList<>();
            try (Stream<ReadResult<Map<String, String>>> s = ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .readAsStream()) {
                s.forEach(r -> viaStream.add(r.data()));
            }

            assertEquals(viaRead, viaStream);
        }
    }

    @Nested
    @DisplayName("Mixed-mode runtime guards")
    class MixedModeGuards {

        @Test
        void column_setterOnly_throwsIllegalStateException() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.column((row, cell) -> {}));
        }

        @Test
        void column_nameAndSetter_throwsIllegalStateException() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.column("Name", (row, cell) -> {}));
        }

        @Test
        void columnAt_throwsIllegalStateException() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.columnAt(0, (row, cell) -> {}));
        }

        @Test
        void skipColumn_throwsIllegalStateException() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class, reader::skipColumn);
        }

        @Test
        void skipColumns_throwsIllegalStateException() {
            var reader = ExcelReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.skipColumns(3));
        }

        @Test
        void guardMessage_identifiesMethodName() {
            var reader = ExcelReader.forMap();
            var ex = assertThrows(IllegalStateException.class,
                    () -> reader.column((row, cell) -> {}));
            assertTrue(ex.getMessage().contains("column(BiConsumer)"),
                    "error message should name the rejected method");
            assertTrue(ex.getMessage().contains("forMap()"),
                    "error message should point the user at forMap()");
        }

        @Test
        void setterMode_stillAllowsColumnCalls() {
            // Sanity check: plain ExcelReader (not from forMap) still works as before.
            var reader = new ExcelReader<>(Object::new, null);
            assertDoesNotThrow(() -> reader.column((row, cell) -> {}));
            assertDoesNotThrow(() -> reader.column("Name", (row, cell) -> {}));
            assertDoesNotThrow(() -> reader.columnAt(0, (row, cell) -> {}));
            assertDoesNotThrow(reader::skipColumn);
            assertDoesNotThrow(() -> reader.skipColumns(2));
        }
    }

    @Nested
    @DisplayName("Behavioral equivalence with deleted ExcelMapReader")
    class BehavioralEquivalence {

        /**
         * Builds an Excel file by hand with an explicit row layout (including nulls and
         * short rows) so we can test the edges that the normal ExcelWriter/XSSFWorkbook
         * round trip would paper over.
         */
        private byte[] writeHandCraftedWorkbook(HandCrafter crafter) throws IOException {
            try (XSSFWorkbook wb = new XSSFWorkbook();
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                Sheet sheet = wb.createSheet("Sheet1");
                crafter.craft(sheet);
                wb.write(out);
                return out.toByteArray();
            }
        }

        @FunctionalInterface
        private interface HandCrafter {
            void craft(Sheet sheet);
        }

        @Test
        @DisplayName("row with fewer cells than headers: trailing header keys are absent")
        void excelMapReader_fewerDataColumnsThanHeaders() throws IOException {
            byte[] data = writeHandCraftedWorkbook(sheet -> {
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Age");
                header.createCell(2).setCellValue("City");
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue("Alice");
                dataRow.createCell(1).setCellValue("30");
                // City column intentionally left unwritten — row has only 2 cells
            });

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
            assertFalse(results.get(0).containsKey("City"),
                    "trailing missing cell means the corresponding header key is absent "
                            + "(preserved from deleted ExcelMapReader)");
        }

        @Test
        @DisplayName("row with more cells than headers: the extras are ignored")
        void excelMapReader_moreDataColumnsThanHeaders() throws IOException {
            byte[] data = writeHandCraftedWorkbook(sheet -> {
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Age");
                // Only 2 headers
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue("Alice");
                dataRow.createCell(1).setCellValue("30");
                dataRow.createCell(2).setCellValue("extra-ignored");
                dataRow.createCell(3).setCellValue("also-ignored");
            });

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals(2, results.get(0).size(),
                    "map shape is bounded by header count, not data-cell count");
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
            assertFalse(results.get(0).containsKey("extra-ignored"));
        }

        @Test
        @DisplayName("blank header cell becomes an empty-string map key (CellData compact ctor coerces null → \"\")")
        void excelMapReader_blankHeaderCell_becomesEmptyStringKey() throws IOException {
            // This test documents an important subtlety: ExcelReadHandler's cell() callback
            // backfills missing cells with CellData(i, null), which the record's compact
            // constructor then coerces to formattedValue="". The extracted header is
            // therefore "" (not null), and the mapMapper's "if (header == null) continue"
            // branch is effectively dead code for Excel. The deleted ExcelMapReader had the
            // same behavior — the apparent "null filter" in its code was a no-op for the
            // same reason. This test pins the behavior so future CellData changes don't
            // silently alter it.
            byte[] data = writeHandCraftedWorkbook(sheet -> {
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Name");
                // Column 1 header intentionally not written — gets backfilled as empty
                header.createCell(2).setCellValue("City");
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue("Alice");
                dataRow.createCell(1).setCellValue("30");
                dataRow.createCell(2).setCellValue("Seoul");
            });

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("Seoul", results.get(0).get("City"));
            assertTrue(results.get(0).containsKey(""),
                    "blank header cell becomes an empty-string map key, matching the "
                            + "deleted ExcelMapReader's behavior");
            assertEquals("30", results.get(0).get(""));
            assertEquals(3, results.get(0).size());
            assertFalse(results.get(0).containsKey(null));
        }

        @Test
        @DisplayName("empty cell value (present but blank) maps to empty string, not null")
        void excelMapReader_presentButEmptyCell_becomesEmptyString() throws IOException {
            byte[] data = writeHandCraftedWorkbook(sheet -> {
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Age");
                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue("Alice");
                dataRow.createCell(1).setCellValue("");  // explicit empty string
            });

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .build(new ByteArrayInputStream(data))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertTrue(results.get(0).containsKey("Age"),
                    "present cell should still put the header key even when empty");
            // Excel treats empty string cells specially — the cell may either be written as
            // an empty string or collapsed to no value. Either way, the map value must be
            // a non-null String per the forMap() contract (the deleted ExcelMapReader
            // inherits CellData's null→"" coercion).
            assertNotNull(results.get(0).get("Age"),
                    "CellData's compact constructor coerces null to empty string, so values "
                            + "in forMap() maps are never null");
            assertEquals("", results.get(0).get("Age"));
        }
    }

    @Nested
    @DisplayName("Read path error surfacing (sheetIndex out of range)")
    class ReadPathErrors {

        @Test
        @DisplayName("read() on a non-existent sheet throws ExcelReadException")
        void read_nonExistentSheet_throws() throws IOException {
            byte[] data = writeSampleExcel();
            assertThrows(ExcelReadException.class, () ->
                    ExcelReader.forMap()
                            .sheetIndex(99)
                            .build(new ByteArrayInputStream(data))
                            .read(r -> {}));
        }
    }

    @Nested
    @DisplayName("ExcelWorkbook multi-sheet + sheetIndex selection")
    class MultiSheetSelection {

        @Test
        @DisplayName("sheetIndex(1) reads the second sheet, not the first")
        void sheetIndex_selectsSecondSheet() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<Map<String, Object>>sheet("First")
                        .column("Name", m -> m.get("Name"))
                        .write(Stream.<Map<String, Object>>of(Map.of("Name", "first-sheet-row")));
                wb.<Map<String, Object>>sheet("Second")
                        .column("Name", m -> m.get("Name"))
                        .column("Tag", m -> m.get("Tag"))
                        .write(Stream.<Map<String, Object>>of(
                                Map.of("Name", "second-a", "Tag", "x"),
                                Map.of("Name", "second-b", "Tag", "y")));
                wb.finish().write(out);
            }

            List<Map<String, String>> results = new ArrayList<>();
            ExcelReader.forMap()
                    .sheetIndex(1)
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
            assertEquals("second-a", results.get(0).get("Name"));
            assertEquals("x", results.get(0).get("Tag"));
            assertEquals("second-b", results.get(1).get("Name"));
            // Verify we didn't accidentally read the first sheet
            assertFalse(results.stream().anyMatch(r -> "first-sheet-row".equals(r.get("Name"))),
                    "sheetIndex(1) must not leak rows from sheet 0");
        }
    }
}
