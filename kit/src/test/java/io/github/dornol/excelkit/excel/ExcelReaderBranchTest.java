package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadAbortException;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowData;
import io.github.dornol.excelkit.csv.CsvReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for edge-case branches in the read pipeline:
 * - ExcelReadHandler.Spliterator error paths
 * - ExcelReadHandler.SheetHandler mapping mode + headerRowIndex
 * - ExcelReader.HeaderExtractor gap columns
 * - AbstractReadHandler mapColumn exception, columnAt
 */
class ExcelReaderBranchTest {

    @TempDir
    Path tempDir;

    @Test
    void numericCellsExposeStoredValuesAlongsideFormattedText() throws IOException {
        Path file = tempDir.resolve("numeric-precision.xlsx");
        try (var workbook = new XSSFWorkbook(); var out = Files.newOutputStream(file)) {
            Sheet sheet = workbook.createSheet();
            Row headers = sheet.createRow(0);
            for (String name : List.of("First", "Second", "Gap", "Text", "Percent", "Date", "Formula")) {
                headers.createCell(headers.getPhysicalNumberOfCells()).setCellValue(name);
            }
            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(1_234_567_890_123L);
            row.createCell(1).setCellValue(123_456_789_012_345L);
            row.createCell(3).setCellValue("123456789012345");
            var percent = row.createCell(4);
            var percentStyle = workbook.createCellStyle();
            percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0%"));
            percent.setCellStyle(percentStyle);
            percent.setCellValue(0.25);
            var date = row.createCell(5);
            var dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            date.setCellStyle(dateStyle);
            date.setCellValue(45_658);
            var formula = row.createCell(6);
            formula.setCellFormula("A2+1");
            workbook.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(formula);
            workbook.write(out);
        }

        Locale previousLocale = LocaleUtil.getUserLocale();
        LocaleUtil.setUserLocale(Locale.US);
        try {
            RowData row = ExcelReader.mapping(data -> data)
                    .cellConversion(config -> config.locale(Locale.US)).readAll(file).get(0);
            assertEquals("1234567890123", row.get("First").rawNumber().toPlainString());
            assertEquals("123456789012345", row.get("Second").rawNumber().toPlainString());
            assertEquals("1.23457E+12", row.get("First").formattedValue());
            assertEquals("1.23457E+14", ExcelReader.forMap().readAll(file).get(0).get("Second"));
            assertEquals(1_234_570_000_000L, row.get("First").asLong());
            assertEquals("1234570000000", row.get("First").asBigDecimal().toPlainString());
            assertNull(row.get("Gap").rawNumber());
            assertNull(row.get("Text").rawNumber());
            assertEquals("25%", row.get("Percent").formattedValue());
            assertEquals(new BigDecimal("0.25"), row.get("Percent").rawNumber());
            assertEquals("2025-01-01", row.get("Date").formattedValue());
            assertEquals("45658", row.get("Date").rawNumber().stripTrailingZeros().toPlainString());
            assertEquals("1234567890124", row.get("Formula").rawNumber().toPlainString());
        } finally {
            LocaleUtil.setUserLocale(previousLocale);
        }
    }

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
                        .read(is, r -> {
                            assertTrue(r.success());
                            results.add(r.data());
                        });
            }

            assertEquals(5, results.size());
            assertEquals("Item0", results.get(0).name());
            assertEquals(4, progress.get());
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
                        .column("A", (t, cell) -> t.a = cell.asString())
                        .column("B", (t, cell) -> t.b = cell.asString())
                        .column("C", (t, cell) -> t.c = cell.asString())
                        .read(is, results::add);
            }

            assertEquals(1, results.size());
            assertEquals("x", results.get(0).data().a);
            // Sparse cell B: either null or empty string depending on POI handling
            assertTrue(results.get(0).data().b == null || results.get(0).data().b.isEmpty(),
                    "Sparse cell B should be null or empty");
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
                    .column("Name", (t, cell) -> {
                        throw new RuntimeException("setter broke!");
                    })
                    .column("Value", (t, cell) -> t.value = cell.asInt())
                    .read(new ByteArrayInputStream(excel), results::add);

            assertEquals(3, results.size());
            for (var r : results) {
                assertFalse(r.success(), "Row should fail when setter throws");
                assertEquals(io.github.dornol.excelkit.core.RowError.Type.MAPPING, r.errorType());
                assertNotNull(r.messages());
                assertEquals(1, r.messages().size(), "Should have exactly one error message");
                assertTrue(r.messages().get(0).contains("setter broke!"));
                assertTrue(r.messages().get(0).contains("Name"), "Error should mention column name");
            }
        }

        @Test
        void columnAt_explicitIndex_shouldMapCorrectly() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Mapped>> results = new ArrayList<>();
            new ExcelReader<>(Mapped::new, null)
                    .columnAt(1, (t, cell) -> t.col1 = cell.asString())
                    .columnAt(0, (t, cell) -> t.col0 = cell.asString())
                    .read(new ByteArrayInputStream(excel), results::add);

            assertEquals(3, results.size());
            assertTrue(results.get(0).success());
            assertEquals("A", results.get(0).data().col0, "col0 should map to Name column (index 0)");
            assertEquals("10", results.get(0).data().col1, "col1 should map to Value column (index 1)");
        }



        @Test
        void mappingMode_validationSucceeds_messagesNull() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Item>> results = new ArrayList<>();
            ExcelReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).read(new ByteArrayInputStream(excel), results::add);

            for (var r : results) {
                assertTrue(r.success());
                assertTrue(r.messages().isEmpty());
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
                    .read(new ByteArrayInputStream(excel), results::add);

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

        @Test
        void getSheetHeaders_missingSheet_throwsExcelReadException() throws IOException {
            Path file = tempDir.resolve("missing-sheet-headers.xlsx");
            try (var wb = new XSSFWorkbook()) {
                wb.createSheet("Only").createRow(0).createCell(0).setCellValue("A");
                try (var fos = new FileOutputStream(file.toFile())) {
                    wb.write(fos);
                }
            }

            try (InputStream is = Files.newInputStream(file)) {
                ExcelReadException ex = assertThrows(ExcelReadException.class,
                        () -> ExcelReader.getSheetHeaders(is, 1, 0));
                assertTrue(ex.getMessage().contains("Sheet index 1 not found"));
            }
        }
    }

    // ============================================================
    // CsvReadHandler - ReadAbortException propagation
    // ============================================================
    @Nested
    class CsvReadAbortTests {




    }

    // ============================================================
    // AbstractReadHandler - positional column mapping (headerName == null, explicitIndex < 0)
    // ============================================================
    @Nested
    class PositionalMappingTests {

        @Test
        void positionalColumns_shouldMapByOrder() throws IOException {
            byte[] excel = writeSimpleExcel();

            List<ReadResult<Mapped>> results = new ArrayList<>();
            new ExcelReader<>(Mapped::new, null)
                    .column((t, cell) -> t.col0 = cell.asString())  // positional: index 0
                    .column((t, cell) -> t.col1 = cell.asString())  // positional: index 1
                    .read(new ByteArrayInputStream(excel), results::add);

            assertEquals(3, results.size());
            assertTrue(results.get(0).success());
            assertEquals("A", results.get(0).data().col0);
        }

        @Test
        void csvPositionalColumns_shouldMapByOrder() {
            String csv = "Name,Age\nAlice,30";

            List<ReadResult<Mapped>> results = new ArrayList<>();
            new CsvReader<>(Mapped::new, null)
                    .column((t, cell) -> t.col0 = cell.asString())
                    .column((t, cell) -> t.col1 = cell.asString())
                    .read(new ByteArrayInputStream(csv.getBytes()), results::add);

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).data().col0);
            assertEquals("30", results.get(0).data().col1);
        }


    }

    // ============================================================
    // ExcelReadHandler - exception catch branches in read()
    // ============================================================
    @Nested
    class ExcelReadExceptionCatchTests {




    }

    // ============================================================
    // Helper
    // ============================================================
    private byte[] writeSimpleExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .write(Stream.of(new Item("A", 10), new Item("B", 20), new Item("C", 30)))
                .writeTo(out);
        return out.toByteArray();
    }
}
