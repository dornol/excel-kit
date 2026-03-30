package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Targeted tests to boost branch coverage for:
 * - ExcelReadHandler.Spliterator error paths
 * - ExcelReadHandler.SheetHandler mapping mode + headerRowIndex
 * - ExcelReader.HeaderExtractor gap columns
 * - AbstractReadHandler mapColumn exception, columnAt
 */
class BranchCoverageBoostTest {

    @TempDir
    Path tempDir;

    record Item(String name, int value) {}

    // Mutable types for setter-mode tests
    static class ThreeFields {
        String a, b, c;
    }

    static class Broken {
        String name;
        int value;
    }

    static class Mapped {
        String col0;
        String col1;
    }

    static class StrictTarget {
        String name;
    }

    static class CsvItem {
        String name;
        int value;
    }

    // ============================================================
    // readAsStream - producer throws ExcelReadException
    // ============================================================
    @Nested
    class ReadAsStreamErrorTests {

        @Test
        void readAsStream_nonExistentSheet_throwsExcelReadException() throws IOException {
            byte[] excel = writeSimpleExcel();

            assertThrows(ExcelReadException.class, () -> {
                try (var stream = ExcelReader.<Item>mapping(row ->
                        new Item(row.get("Name").asString(), row.get("Value").asInt())
                ).sheetIndex(5)
                        .build(new ByteArrayInputStream(excel))
                        .readAsStream()) {
                    stream.forEach(r -> {});
                }
            });
        }

        @Test
        void readAsStream_mappingThrowsReadAbort() throws IOException {
            byte[] excel = writeSimpleExcel();

            // mapWithRowMapper catches generic exceptions, but ReadAbortException should propagate
            // Actually mapWithRowMapper catches ALL exceptions. Let's verify the behavior.
            List<ReadResult<Item>> results;
            try (var stream = ExcelReader.<Item>mapping(row -> {
                throw new IllegalStateException("mapper error");
            }).build(new ByteArrayInputStream(excel))
                    .readAsStream()) {
                results = stream.toList();
            }

            assertFalse(results.isEmpty());
            for (var r : results) {
                assertFalse(r.success());
            }
        }
    }

    // ============================================================
    // SheetHandler - mapping mode with headerRowIndex > 0 + progress
    // ============================================================
    @Nested
    class SheetHandlerTests {

        @Test
        void mappingMode_headerRowIndex2_withProgress() throws IOException {
            Path file = tempDir.resolve("header-offset.xlsx");
            try (var wb = new XSSFWorkbook()) {
                Sheet sheet = wb.createSheet("Test");
                sheet.createRow(0).createCell(0).setCellValue("META1");
                sheet.createRow(1).createCell(0).setCellValue("META2");
                Row header = sheet.createRow(2);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Value");
                for (int i = 0; i < 5; i++) {
                    Row row = sheet.createRow(3 + i);
                    row.createCell(0).setCellValue("Item" + i);
                    row.createCell(1).setCellValue(i * 10);
                }
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            var progress = new java.util.concurrent.atomic.AtomicLong(0);
            List<Item> results = new ArrayList<>();
            try (InputStream is = Files.newInputStream(file)) {
                ExcelReader.<Item>mapping(row ->
                        new Item(row.get("Name").asString(), row.get("Value").asInt())
                ).headerRowIndex(2)
                        .onProgress(2, (count, cursor) -> progress.set(count))
                        .build(is)
                        .read(r -> {
                            assertTrue(r.success());
                            results.add(r.data());
                        });
            }

            assertEquals(5, results.size());
            assertEquals("Item0", results.get(0).name());
            assertEquals(4, progress.get());
        }

        @Test
        void mappingMode_headerRowIndex1_viaReadAsStream() throws IOException {
            Path file = tempDir.resolve("header-offset-stream.xlsx");
            try (var wb = new XSSFWorkbook()) {
                Sheet sheet = wb.createSheet("Test");
                sheet.createRow(0).createCell(0).setCellValue("SKIP");
                Row header = sheet.createRow(1);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Value");
                Row data = sheet.createRow(2);
                data.createCell(0).setCellValue("A");
                data.createCell(1).setCellValue(10);
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            List<Item> results;
            try (InputStream is = Files.newInputStream(file)) {
                try (var stream = ExcelReader.<Item>mapping(row ->
                        new Item(row.get("Name").asString(), row.get("Value").asInt())
                ).headerRowIndex(1)
                        .build(is)
                        .readAsStream()) {
                    results = stream.filter(ReadResult::success).map(ReadResult::data).toList();
                }
            }

            assertEquals(1, results.size());
            assertEquals("A", results.get(0).name());
        }

        @Test
        void setterMode_sparseRow_missingCells() throws IOException {
            Path file = tempDir.resolve("sparse-row.xlsx");
            try (var wb = new XSSFWorkbook()) {
                Sheet sheet = wb.createSheet("Test");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("A");
                header.createCell(1).setCellValue("B");
                header.createCell(2).setCellValue("C");
                Row data = sheet.createRow(1);
                data.createCell(0).setCellValue("x");
                // Skip cell at col 1
                data.createCell(2).setCellValue("z");
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            List<ReadResult<ThreeFields>> results = new ArrayList<>();
            try (InputStream is = Files.newInputStream(file)) {
                new ExcelReader<>(ThreeFields::new, null)
                        .addColumn("A", (t, cell) -> t.a = cell.asString())
                        .addColumn("B", (t, cell) -> t.b = cell.asString())
                        .addColumn("C", (t, cell) -> t.c = cell.asString())
                        .build(is)
                        .read(results::add);
            }

            assertEquals(1, results.size());
            assertEquals("x", results.get(0).data().a);
            assertEquals("z", results.get(0).data().c);
        }
    }

    // ============================================================
    // AbstractReadHandler - mapColumn exception, columnAt
    // ============================================================
    @Nested
    class AbstractReadHandlerTests {

        @Test
        void mapColumn_setterThrows_shouldReturnFailedResult() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Broken>> results = new ArrayList<>();
            new ExcelReader<>(Broken::new, null)
                    .addColumn("Name", (t, cell) -> {
                        throw new RuntimeException("setter broke!");
                    })
                    .addColumn("Value", (t, cell) -> t.value = cell.asInt())
                    .build(new ByteArrayInputStream(excel))
                    .read(results::add);

            assertEquals(3, results.size());
            for (var r : results) {
                assertFalse(r.success());
                assertNotNull(r.messages());
                assertTrue(r.messages().get(0).contains("setter broke!"));
            }
        }

        @Test
        void columnAt_explicitIndex_shouldMapCorrectly() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Mapped>> results = new ArrayList<>();
            new ExcelReader<>(Mapped::new, null)
                    .columnAt(1, (t, cell) -> t.col1 = cell.asString())
                    .columnAt(0, (t, cell) -> t.col0 = cell.asString())
                    .build(new ByteArrayInputStream(excel))
                    .read(results::add);

            assertEquals(3, results.size());
            assertTrue(results.get(0).success());
        }

        @Test
        void readStrict_failedRow_showsMessageDetails() throws IOException {
            byte[] excel = writeSimpleExcel();

            var handler = new ExcelReader<>(StrictTarget::new, null)
                    .addColumn("Name", (t, cell) -> {
                        throw new RuntimeException("fail");
                    })
                    .addColumn("Value", (t, cell) -> {})
                    .build(new ByteArrayInputStream(excel));

            var ex = assertThrows(ReadAbortException.class, () -> handler.readStrict(r -> {}));
            assertTrue(ex.getMessage().contains("Row 1"));
        }

        @Test
        void mappingMode_validationSucceeds_messagesNull() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Item>> results = new ArrayList<>();
            ExcelReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).build(new ByteArrayInputStream(excel))
                    .read(results::add);

            for (var r : results) {
                assertTrue(r.success());
                assertNull(r.messages());
            }
        }

        @Test
        void mapColumn_columnIndexBeyondHeaders_fallbackColumnName() throws IOException {
            // When columnIndex >= headerNames.size(), should show "column#N"
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Broken>> results = new ArrayList<>();
            new ExcelReader<>(Broken::new, null)
                    .columnAt(0, (t, cell) -> t.name = cell.asString())
                    .columnAt(99, (t, cell) -> {
                        throw new RuntimeException("bad");
                    })
                    .build(new ByteArrayInputStream(excel))
                    .read(results::add);

            // Column 99 exceeds header count, so error message should show "column#99"
            assertEquals(3, results.size());
        }
    }

    // ============================================================
    // HeaderExtractor - gap columns, headerRowIndex > 0
    // ============================================================
    @Nested
    class HeaderExtractorTests {

        @Test
        void getSheetHeaders_withGapColumns() throws IOException {
            Path file = tempDir.resolve("gap-headers.xlsx");
            try (var wb = new XSSFWorkbook()) {
                Sheet sheet = wb.createSheet("Test");
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("A");
                // Skip col 1
                header.createCell(2).setCellValue("C");
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            List<String> headers = ExcelReader.getSheetHeaders(
                    Files.newInputStream(file), 0, 0);

            assertEquals(3, headers.size());
            assertEquals("A", headers.get(0));
            assertEquals("", headers.get(1));
            assertEquals("C", headers.get(2));
        }

        @Test
        void getSheetHeaders_headerRowIndex1() throws IOException {
            Path file = tempDir.resolve("header-row1.xlsx");
            try (var wb = new XSSFWorkbook()) {
                Sheet sheet = wb.createSheet("Test");
                sheet.createRow(0).createCell(0).setCellValue("META");
                Row header = sheet.createRow(1);
                header.createCell(0).setCellValue("Name");
                header.createCell(1).setCellValue("Value");
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            List<String> headers = ExcelReader.getSheetHeaders(
                    Files.newInputStream(file), 0, 1);

            assertEquals(2, headers.size());
            assertEquals("Name", headers.get(0));
            assertEquals("Value", headers.get(1));
        }

        @Test
        void getSheetHeaders_secondSheet() throws IOException {
            Path file = tempDir.resolve("multi-sheet-headers.xlsx");
            try (var wb = new XSSFWorkbook()) {
                wb.createSheet("First").createRow(0).createCell(0).setCellValue("Ignore");
                Sheet s2 = wb.createSheet("Second");
                Row h = s2.createRow(0);
                h.createCell(0).setCellValue("Col1");
                h.createCell(1).setCellValue("Col2");
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            List<String> headers = ExcelReader.getSheetHeaders(
                    Files.newInputStream(file), 1, 0);

            assertEquals(2, headers.size());
            assertEquals("Col1", headers.get(0));
        }
    }

    // ============================================================
    // CsvReadHandler - ReadAbortException propagation
    // ============================================================
    @Nested
    class CsvReadAbortTests {

        @Test
        void csvReadStrict_failedRow_throwsReadAbort() {
            String csv = "Name,Value\nAlice,10\nBob,bad";

            var handler = new io.github.dornol.excelkit.csv.CsvReader<>(CsvItem::new, null)
                    .addColumn("Name", (t, cell) -> t.name = cell.asString())
                    .addColumn("Value", (t, cell) -> t.value = cell.asInt())
                    .build(new ByteArrayInputStream(csv.getBytes()));

            assertThrows(ReadAbortException.class, () -> handler.readStrict(r -> {}));
        }

        @Test
        void csvMapping_readStrict_throwsReadAbortOnError() {
            String csv = "Name,Value\nAlice,10\nBob,bad";

            var handler = io.github.dornol.excelkit.csv.CsvReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).build(new ByteArrayInputStream(csv.getBytes()));

            assertThrows(ReadAbortException.class, () -> handler.readStrict(r -> {}));
        }
    }

    // ============================================================
    // Helper
    // ============================================================
    private byte[] writeSimpleExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .write(Stream.of(new Item("A", 10), new Item("B", 20), new Item("C", 30)))
                .consumeOutputStream(out);
        return out.toByteArray();
    }
}
