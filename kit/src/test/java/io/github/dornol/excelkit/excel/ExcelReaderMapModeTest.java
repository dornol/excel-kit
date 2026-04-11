package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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
}
