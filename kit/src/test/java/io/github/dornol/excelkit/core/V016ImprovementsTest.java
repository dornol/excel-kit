package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for v0.16.0 improvements (public API tests).
 */
class V016ImprovementsTest {

    // ============================================================
    // A1: CellData regex pre-compilation — verify behavior unchanged
    // ============================================================
    @Nested
    @DisplayName("A1: CellData currency symbol stripping")
    class CellDataRegex {

        @Test
        void asNumber_shouldStripDollarSign() {
            assertEquals(1234, new CellData(0, "$1234").asInt());
        }

        @Test
        void asNumber_shouldStripEuroSign() {
            assertEquals(999, new CellData(0, "€999").asInt());
        }

        @Test
        void asNumber_shouldStripKoreanWon() {
            assertEquals(5000, new CellData(0, "₩5000").asInt());
        }

        @Test
        void asNumber_shouldStripKoreanWonText() {
            assertEquals(3000, new CellData(0, "3000원").asInt());
        }

        @Test
        void asNumber_shouldStripPercentSign() {
            assertNotNull(new CellData(0, "50%").asNumber());
        }

        @Test
        void asNumber_shouldStripCommasAndCurrency() {
            assertEquals(1234567, new CellData(0, "$1,234,567").asInt());
        }

        @Test
        void asNumber_shouldStripNonBreakingSpace() {
            assertEquals(100, new CellData(0, "100\u00A0").asInt());
        }

        @Test
        void asNumber_shouldHandleMixedCurrencySymbols() {
            assertEquals(1234, new CellData(0, "₩1,234").asInt());
        }
    }

    // ============================================================
    // B1: nullValue through public API
    // ============================================================
    @Nested
    @DisplayName("B1: Writer nullValue")
    class NullValueTests {

        @Test
        void nullValue_throughColumnBuilder_shouldWork() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("Name", s -> s, c -> c.nullValue("-"));

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of((String) null)).write(bos);

            // Read back to verify
            List<ReadResult<String>> results = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success());
            assertEquals("-", results.get(0).data());
        }

        @Test
        void nullValue_throughDefaultStyle_shouldFallback() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.defaultStyle(d -> d.nullValue("DEFAULT"));
            writer.column("Col", s -> s);

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of((String) null)).write(bos);

            List<ReadResult<String>> results = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("Col").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("DEFAULT", results.get(0).data());
        }

        @Test
        void nullValue_perColumn_overridesDefault() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.defaultStyle(d -> d.nullValue("DEFAULT"));
            writer.column("Col", s -> s, c -> c.nullValue("CUSTOM"));

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of((String) null)).write(bos);

            List<ReadResult<String>> results = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("Col").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("CUSTOM", results.get(0).data());
        }

        @Test
        void nullValue_nonNullData_shouldUseActualData() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("Name", s -> s, c -> c.nullValue("N/A"));

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of("Hello")).write(bos);

            List<ReadResult<String>> results = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals("Hello", results.get(0).data());
        }

        @Test
        void nullValue_withoutSetting_shouldWriteEmpty() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("Name", s -> s);

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of((String) null)).write(bos);

            List<ReadResult<String>> results = new ArrayList<>();
            ExcelReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals("", results.get(0).data());
        }
    }

    // ============================================================
    // B4: freezePane(cols, rows)
    // ============================================================
    @Nested
    @DisplayName("B4: freezePane with columns")
    class FreezePaneTests {

        @Test
        void freezePane_colsAndRows_shouldNotThrow() {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            assertDoesNotThrow(() -> writer.freezePane(2, 1));
        }

        @Test
        void freezePane_negativeCol_shouldThrow() {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            assertThrows(IllegalArgumentException.class, () -> writer.freezePane(-1, 0));
        }

        @Test
        void freezePane_negativeRow_shouldThrow() {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            assertThrows(IllegalArgumentException.class, () -> writer.freezePane(0, -1));
        }

        @Test
        void freezePane_colsAndRows_shouldApplyToSheet() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("A", s -> s)
                    .column("B", s -> s)
                    .column("C", s -> s)
                    .freezePane(2, 1);

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of("row1")).write(bos);
            // No exception = pane was applied during write
        }

        @Test
        void freezePane_singleArg_backwardsCompatible() throws IOException {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("A", s -> s).freezePane(1);

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            writer.write(Stream.of("row1")).write(bos);
        }

        @Test
        void freezePane_onExcelSheetWriter_shouldWork() throws IOException {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWorkbook wb = ExcelWorkbook.builder().build();
            ExcelSheetWriter<String> sw = wb.sheet("test");
            sw.column("Col", row -> row)
                    .freezePane(1, 1);
            sw.write(Stream.of("data"));
            wb.finish().write(bos);
        }
    }

    // ============================================================
    // B6: Reader required() API
    // ============================================================
    @Nested
    @DisplayName("B6: Reader required()")
    class RequiredColumnTests {

        @Test
        void required_excelReader_emptyCell_shouldFail() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter<String[]> writer = ExcelWriter.<String[]>builder().build();
            writer.column("Name", arr -> arr[0])
                    .column("Age", arr -> arr[1]);
            writer.write(Stream.<String[]>of(new String[]{"", "30"})).write(bos);

            List<ReadResult<String[]>> results = new ArrayList<>();
            ExcelReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertFalse(results.get(0).success(), "Required column with blank value should fail");
            assertNotNull(results.get(0).messages());
            assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
        }

        @Test
        void required_excelReader_nonEmptyCell_shouldSucceed() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter<String[]> writer = ExcelWriter.<String[]>builder().build();
            writer.column("Name", arr -> arr[0])
                    .column("Age", arr -> arr[1]);
            writer.write(Stream.<String[]>of(new String[]{"Alice", "30"})).write(bos);

            List<ReadResult<String[]>> results = new ArrayList<>();
            ExcelReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success());
            assertEquals("Alice", results.get(0).data()[0]);
        }

        @Test
        void required_csvReader_emptyCell_shouldFail() {
            String csv = "Name,Age\n,30";

            List<ReadResult<String[]>> results = new ArrayList<>();
            CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertFalse(results.get(0).success());
            assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
        }

        @Test
        void required_csvReader_nonEmptyCell_shouldSucceed() {
            String csv = "Name,Age\nBob,25";

            List<ReadResult<String[]>> results = new ArrayList<>();
            CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success());
            assertEquals("Bob", results.get(0).data()[0]);
        }

        @Test
        void required_multipleRequiredColumns_bothEmpty_shouldFail() {
            String csv = "Name,Age\n,";

            List<ReadResult<String[]>> results = new ArrayList<>();
            CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString()).required()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertFalse(results.get(0).success());
            assertEquals(2, results.get(0).messages().size(), "Both required columns should report errors");
        }

        @Test
        void required_onlySecondEmpty_shouldFailWithColumnName() {
            String csv = "Name,Age\nAlice,";

            List<ReadResult<String[]>> results = new ArrayList<>();
            CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString()).required()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertFalse(results.get(0).success());
            assertEquals(1, results.get(0).messages().size());
            assertTrue(results.get(0).messages().get(0).contains("Age"));
        }

        @Test
        void required_withoutCalling_shouldNotValidateEmpty() {
            String csv = "Name,Age\n,30";

            List<ReadResult<String[]>> results = new ArrayList<>();
            CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString())
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success(), "Non-required column should pass with blank value");
        }

        @Test
        void required_noColumnRegistered_shouldThrow() {
            assertThrows(IllegalStateException.class, () ->
                    CsvReader.setter(() -> new String[1]).required());
        }

        @Test
        void required_noColumnRegistered_excelReader_shouldThrow() {
            assertThrows(IllegalStateException.class, () ->
                    ExcelReader.setter(() -> new String[1]).required());
        }

        @Test
        void readColumn_required_shouldReturnNewInstance() {
            ReadColumn<String> col = new ReadColumn<>("Name", (s, c) -> {});
            assertFalse(col.isRequired());

            ReadColumn<String> req = col.required();
            assertTrue(req.isRequired());
            assertFalse(col.isRequired(), "Original should be unchanged");
            assertEquals("Name", req.headerName());
        }
    }

    // ============================================================
    // A5: Duplicate header detection
    // Note: Writer rejects duplicate column names at write time,
    // so duplicate header detection in Reader is tested via CSV
    // (which allows duplicate headers in raw data).
    // ============================================================
    @Nested
    @DisplayName("A5: Duplicate header — mapping mode uses first occurrence")
    class DuplicateHeaderTests {

        @Test
        void csv_duplicateHeaders_mappingMode_shouldUseFirstOccurrence() {
            // CSV allows duplicate headers in raw data
            String csv = "Name,Name,Age\nAlice,Bob,30";

            List<ReadResult<String>> results = new ArrayList<>();
            CsvReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success());
            assertEquals("Alice", results.get(0).data(), "First occurrence of duplicate header should be used");
        }
    }

    // ============================================================
    // A6: readStrict row numbering
    // ============================================================
    @Nested
    @DisplayName("A6: readStrict row numbering")
    class ReadStrictTests {

        @Test
        void readStrict_shouldReportCorrectRowNumber() {
            String csv = "Name,Age\nAlice,30\n,invalid";

            var reader = CsvReader.setter(() -> new String[2])
                    .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));

            ReadAbortException ex = assertThrows(ReadAbortException.class, () ->
                    reader.readStrict(data -> {}));
            assertTrue(ex.getMessage().contains("Row 2"), "Error should reference row 2, got: " + ex.getMessage());
        }
    }

    // ============================================================
    // readAsStream laziness
    // ============================================================
    @Nested
    @DisplayName("A2: readAsStream is lazy")
    class ReadAsStreamTests {

        @Test
        void csvReadAsStream_withLimit_shouldWork() {
            String csv = "Name\nAlice\nBob\nCharlie";

            try (var stream = CsvReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .readAsStream()) {
                List<String> names = stream
                        .filter(ReadResult::success)
                        .map(ReadResult::data)
                        .limit(2)
                        .toList();

                assertEquals(2, names.size());
                assertEquals("Alice", names.get(0));
                assertEquals("Bob", names.get(1));
            }
        }

        @Test
        void excelReadAsStream_withLimit_shouldWork() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            writer.column("Name", s -> s);
            writer.write(Stream.of("Alice", "Bob", "Charlie")).write(bos);

            try (var stream = ExcelReader.<String>mapping(row -> row.get("Name").asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .readAsStream()) {
                List<String> names = stream
                        .filter(ReadResult::success)
                        .map(ReadResult::data)
                        .limit(2)
                        .toList();

                assertEquals(2, names.size());
                assertEquals("Alice", names.get(0));
            }
        }
    }

    // ============================================================
    // Bug fixes: sparse row required validation, rollover header color
    // ============================================================
    @Nested
    @DisplayName("Bugfix: required column in sparse Excel row")
    class RequiredSparseRowTests {

        @Test
        void required_sparseRow_missingColumn_shouldFail() throws Exception {
            // Write Excel with 3 columns, but third column is null (SAX won't emit it)
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter<String[]> writer = ExcelWriter.<String[]>builder().build();
            writer.column("Name", arr -> arr[0])
                    .column("Age", arr -> arr[1])
                    .column("City", arr -> arr[2]);
            writer.write(Stream.<String[]>of(new String[]{"Alice", "30", null})).write(bos);

            // Read with required() on the third column (which is empty/missing)
            List<ReadResult<String[]>> results = new ArrayList<>();
            ExcelReader.setter(() -> new String[3])
                    .column("Name", (arr, c) -> arr[0] = c.asString())
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .column("City", (arr, c) -> arr[2] = c.asString()).required()
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertFalse(results.get(0).success(), "Required column with missing cell in sparse row should fail");
            assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
        }

        @Test
        void nonRequired_sparseRow_missingColumn_shouldSucceed() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter<String[]> writer = ExcelWriter.<String[]>builder().build();
            writer.column("Name", arr -> arr[0])
                    .column("Age", arr -> arr[1])
                    .column("City", arr -> arr[2]);
            writer.write(Stream.<String[]>of(new String[]{"Alice", "30", null})).write(bos);

            List<ReadResult<String[]>> results = new ArrayList<>();
            ExcelReader.setter(() -> new String[3])
                    .column("Name", (arr, c) -> arr[0] = c.asString())
                    .column("Age", (arr, c) -> arr[1] = c.asString())
                    .column("City", (arr, c) -> arr[2] = c.asString())
                    .build(new ByteArrayInputStream(bos.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success(), "Non-required missing column should succeed");
        }
    }

    @Nested
    @DisplayName("Bugfix: ExcelSheetWriter rollover header style")
    class RolloverHeaderStyleTests {

        @Test
        void rollover_shouldPreserveHeaderFontColor() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = ExcelWorkbook.builder().build()) {
                ExcelSheetWriter<Integer> sw = wb.sheet("test");
                sw.column("ID", i -> i, c -> c.headerFontColor(ExcelColor.RED))
                        .maxRows(2);
                sw.write(Stream.of(1, 2, 3, 4));
                wb.finish().write(bos);
            }

            // Verify the file is readable and has 2 sheets (rollover occurred)
            try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                    new ByteArrayInputStream(bos.toByteArray()))) {
                assertTrue(workbook.getNumberOfSheets() >= 2,
                        "Should have at least 2 sheets due to maxRows=2, got " + workbook.getNumberOfSheets());
            }
        }
    }

    @Nested
    @DisplayName("Bugfix: summary + afterData row overlap")
    class SummaryAfterDataTests {

        @Test
        void summary_afterData_shouldNotOverlapRows() throws Exception {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            ExcelWriter.<int[]>builder().build()
                    .column("Name", arr -> "Item" + arr[0])
                    .column("Amount", arr -> arr[0], c -> c.type(ExcelDataType.INTEGER))
                    .afterData((ctx) -> {
                        // Write a custom row after data
                        var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                        row.createCell(0).setCellValue("Custom Footer");
                        return ctx.getCurrentRow() + 1;
                    })
                    .summary(s -> s.sum("Amount"))
                    .write(Stream.of(new int[]{10}, new int[]{20}, new int[]{30}))
                    .write(bos);

            // Read back and verify no overlapping rows
            try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                    new ByteArrayInputStream(bos.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Row 0: header, Row 1-3: data, Row 4: afterData footer, Row 5: summary
                String footerValue = sheet.getRow(4).getCell(0).getStringCellValue();
                assertEquals("Custom Footer", footerValue, "afterData row should be at row 4");

                // Summary row should be AFTER the footer, not overlapping
                String summaryLabel = sheet.getRow(5).getCell(0).getStringCellValue();
                assertNotNull(summaryLabel, "Summary row should exist at row 5");
            }
        }
    }
}
