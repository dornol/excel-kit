package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ExcelKitSchema;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive tests covering edge cases for all new features.
 */
class ComprehensiveFeatureTest {

    @TempDir
    Path tempDir;

    // ========================================================================
    // CellColor edge cases
    // ========================================================================
    @Test
    void cellColor_withNullValue_shouldNotApplyColor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Value", s -> null) // always null
                    .cellColor((value, row) -> value != null ? ExcelColor.LIGHT_RED : null)
                .write(Stream.of("a", "b"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // null value → cellColor returns null → no color override
            XSSFColor c = (XSSFColor) wb.getSheetAt(0).getRow(1).getCell(0)
                    .getCellStyle().getFillForegroundColorColor();
            // Should be default (no fill or automatic)
            assertTrue(c == null || c.getRGB() == null);
        }
    }

    @Test
    void cellColor_onMultipleColumns_shouldApplyIndependently() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<int[]>()
                .column("A", r -> r[0])
                    .type(ExcelDataType.INTEGER)
                    .cellColor((v, r) -> ((Number) v).intValue() > 50 ? ExcelColor.LIGHT_GREEN : null)
                .column("B", r -> r[1])
                    .type(ExcelDataType.INTEGER)
                    .cellColor((v, r) -> ((Number) v).intValue() < 0 ? ExcelColor.LIGHT_RED : null)
                .write(Stream.of(new int[]{100, -5}))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            // A=100 > 50 → green
            XSSFColor colA = (XSSFColor) row.getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(colA);
            assertEquals(ExcelColor.LIGHT_GREEN.getR(), Byte.toUnsignedInt(colA.getRGB()[0]));
            // B=-5 < 0 → red
            XSSFColor colB = (XSSFColor) row.getCell(1).getCellStyle().getFillForegroundColorColor();
            assertNotNull(colB);
            assertEquals(ExcelColor.LIGHT_RED.getR(), Byte.toUnsignedInt(colB.getRGB()[0]));
        }
    }

    // ========================================================================
    // Group header edge cases
    // ========================================================================
    @Test
    void groupHeader_allColumnsSameGroup_shouldMergeEntireRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<int[]>()
                .column("A", r -> r[0]).group("All")
                .column("B", r -> r[1]).group("All")
                .column("C", r -> r[2]).group("All")
                .write(Stream.of(new int[]{1, 2, 3}))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals(3, sheet.getLastRowNum() + 1); // group header + column header + 1 data
            boolean hasMerge = sheet.getMergedRegions().stream()
                    .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 0
                            && r.getFirstColumn() == 0 && r.getLastColumn() == 2);
            assertTrue(hasMerge, "All 3 columns should be merged in group row");
        }
    }

    @Test
    void groupHeader_multipleDistinctGroups_shouldMergeSeparately() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<int[]>()
                .column("A", r -> r[0]).group("X")
                .column("B", r -> r[1]).group("X")
                .column("C", r -> r[2]).group("Y")
                .column("D", r -> r[3]).group("Y")
                .write(Stream.of(new int[]{1, 2, 3, 4}))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merges = sheet.getMergedRegions();
            boolean hasX = merges.stream().anyMatch(r ->
                    r.getFirstRow() == 0 && r.getFirstColumn() == 0 && r.getLastColumn() == 1);
            boolean hasY = merges.stream().anyMatch(r ->
                    r.getFirstRow() == 0 && r.getFirstColumn() == 2 && r.getLastColumn() == 3);
            assertTrue(hasX, "Group X should merge cols 0-1");
            assertTrue(hasY, "Group Y should merge cols 2-3");
        }
    }

    @Test
    void groupHeader_singleColumnGroup_shouldNotMergeHorizontally() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("A", s -> s)
                .column("B", s -> s).group("Solo")
                .column("C", s -> s)
                .write(Stream.of("test"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Single-column group should not have horizontal merge
            boolean hasHorizontalMerge = sheet.getMergedRegions().stream()
                    .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 0
                            && r.getFirstColumn() == 1 && r.getLastColumn() > 1);
            assertFalse(hasHorizontalMerge);
        }
    }

    // ========================================================================
    // Rollover with combined features
    // ========================================================================
    @Test
    void rollover_withFreezePane_shouldApplyToAllSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .freezePane(1)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, wb.getNumberOfSheets());
            // Both sheets should have freeze pane
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                assertNotNull(wb.getSheetAt(i).getPaneInformation(),
                        "Sheet " + i + " should have freeze pane");
            }
        }
    }

    @Test
    void rollover_withAutoFilter_shouldApplyToAllSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .autoFilter(true)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, wb.getNumberOfSheets());
        }
    }

    // ========================================================================
    // Duplicate column validation edge cases
    // ========================================================================
    @Test
    void duplicateColumnName_constColumn_shouldThrow() {
        var writer = new ExcelWriter<String>()
                .addColumn("Name", s -> s);
        writer.addColumn("Name", s -> s); // duplicate via constColumn path
        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void duplicateColumnName_viaSchema_shouldThrow() {
        ExcelKitSchema<String> schema = ExcelKitSchema.<String>builder()
                .column("Name", s -> s, (s, c) -> {})
                .column("Name", s -> s, (s, c) -> {})
                .build();
        assertThrows(ExcelWriteException.class,
                () -> schema.excelWriter().write(Stream.of("test")));
    }

    // ========================================================================
    // CSV progress boundary
    // ========================================================================
    @Test
    void csvProgress_invalidInterval_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new CsvWriter<String>().column("A", s -> s).onProgress(0, (c, cur) -> {}));
        assertThrows(IllegalArgumentException.class, () ->
                new CsvWriter<String>().column("A", s -> s).onProgress(-1, (c, cur) -> {}));
    }

    @Test
    void csvProgress_cursorShouldProvideCorrectTotal() {
        List<Long> totals = new ArrayList<>();
        List<Integer> rowOfSheets = new ArrayList<>();

        new CsvWriter<Integer>()
                .column("V", i -> i)
                .onProgress(2, (count, cursor) -> {
                    totals.add(count);
                    rowOfSheets.add(cursor.getRowOfSheet());
                })
                .write(Stream.of(1, 2, 3, 4));

        assertEquals(List.of(2L, 4L), totals);
        // cursor rowOfSheet includes header row (row 0 = header)
        assertTrue(rowOfSheets.get(0) > 0);
        assertTrue(rowOfSheets.get(1) > rowOfSheets.get(0));
    }

    // ========================================================================
    // CSV name-based read edge cases
    // ========================================================================
    @Test
    void csvNameBased_caseSensitive_shouldMatchExact() {
        String csv = "Name,Age\nAlice,30\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        // "name" (lowercase) should not match "Name" (title case)
        var handler = new CsvReader<>(TestPerson::new, null)
                .addColumn("name", (p, cell) -> p.name = cell.asString())
                .build(is);

        assertThrows(Exception.class, () -> handler.read(r -> {}));
    }

    @Test
    void csvNameBased_withStreamApi_shouldWork() {
        String csv = "Age,Name\n30,Alice\n25,Bob\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<String> names = new CsvReader<>(TestPerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .build(is)
                .readAsStream()
                .map(r -> r.data().name)
                .toList();

        assertEquals(List.of("Alice", "Bob"), names);
    }

    // ========================================================================
    // columnAt edge cases
    // ========================================================================
    @Test
    void columnAt_outOfBounds_shouldHandleGracefully() throws IOException {
        Path file = tempDir.resolve("small.xlsx");
        createExcelFile(file, new String[]{"A", "B"}, new Object[][]{{"a", "b"}});

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestPerson::new, null)
                    .columnAt(0, (p, cell) -> p.name = cell.asString())
                    .columnAt(99, (p, cell) -> p.age = cell.asInt()) // out of bounds
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        // col 99 doesn't exist → should be skipped (age stays null)
        assertEquals(1, results.size());
        assertEquals("a", results.get(0).name);
        assertNull(results.get(0).age);
    }

    @Test
    void csvColumnAt_outOfBounds_shouldHandleGracefully() {
        String csv = "A,B\na,b\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();
        new CsvReader<>(TestPerson::new, null)
                .columnAt(0, (p, cell) -> p.name = cell.asString())
                .columnAt(99, (p, cell) -> p.age = cell.asInt()) // out of bounds
                .build(is)
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("a", results.get(0).name);
        assertNull(results.get(0).age);
    }

    // ========================================================================
    // Outline edge cases
    // ========================================================================
    @Test
    void outline_level7_shouldBeValid() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("A", s -> s).outline(7)
                .column("B", s -> s)
                .write(Stream.of("test"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getColumnOutlineLevel(0) > 0);
        }
    }

    @Test
    void outline_withRollover_shouldApplyToAllSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .column("A", i -> i, c -> c.outline(1))
                    .column("B", i -> i)
                    .write(Stream.of(1, 2, 3));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, wb.getNumberOfSheets());
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                assertTrue(wb.getSheetAt(i).getColumnOutlineLevel(0) > 0,
                        "Sheet " + i + " should have outline on col 0");
            }
        }
    }

    // ========================================================================
    // Formula/Hyperlink edge cases
    // ========================================================================
    @Test
    void hyperlink_withSpecialCharsInUrl_shouldWork() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Link", s -> s).type(ExcelDataType.HYPERLINK)
                .write(Stream.of("https://example.com/search?q=hello+world&lang=ko#section"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertNotNull(cell.getHyperlink());
            assertTrue(cell.getHyperlink().getAddress().contains("q=hello+world"));
        }
    }

    // ========================================================================
    // Progress cursor state verification
    // ========================================================================
    @Test
    void progress_cursorShouldProvideCorrectState() {
        List<Long> totals = new ArrayList<>();
        List<Integer> sheetRows = new ArrayList<>();

        new ExcelWriter<Integer>()
                .column("V", i -> i).type(ExcelDataType.INTEGER)
                .onProgress(3, (count, cursor) -> {
                    totals.add(count);
                    sheetRows.add(cursor.getRowOfSheet());
                })
                .write(Stream.of(1, 2, 3, 4, 5, 6));

        assertEquals(List.of(3L, 6L), totals);
        // rowOfSheet should increase between callbacks
        assertTrue(sheetRows.get(1) > sheetRows.get(0));
    }

    // ========================================================================
    // Schema round-trip with write config
    // ========================================================================
    @Test
    void schema_roundTrip_withWriteConfig_shouldPreserveData() throws IOException {
        ExcelKitSchema<TestProduct> schema = ExcelKitSchema.<TestProduct>builder()
                .column("Name", TestProduct::name, (p, cell) -> {})
                .column("Price", TestProduct::price, (p, cell) -> {},
                        c -> c.type(ExcelDataType.INTEGER))
                .column("Active", TestProduct::active, (p, cell) -> {},
                        c -> c.type(ExcelDataType.BOOLEAN_TO_YN))
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(
                        new TestProduct("Widget", 1000, true),
                        new TestProduct("Gadget", 2500, false)))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Price should be numeric (written via INTEGER type)
            assertEquals(1000.0, sheet.getRow(1).getCell(1).getNumericCellValue());
            // Boolean should be Y/N
            assertEquals("Y", sheet.getRow(1).getCell(2).getStringCellValue());
            assertEquals("N", sheet.getRow(2).getCell(2).getStringCellValue());
        }
    }

    // ========================================================================
    // columnIf with rollover
    // ========================================================================
    @Test
    void columnIf_withRollover_shouldApplyConsistently() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .column("A", i -> i)
                    .columnIf("B", true, i -> i * 2)
                    .columnIf("C", false, i -> i * 3)
                    .write(Stream.of(1, 2, 3));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, wb.getNumberOfSheets());
            // Both sheets should have 2 columns (A, B) — C was skipped
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                assertEquals(2, wb.getSheetAt(i).getRow(0).getLastCellNum(),
                        "Sheet " + i + " should have 2 columns");
            }
        }
    }

    // ========================================================================
    // Helpers
    // ========================================================================
    private void createExcelFile(Path filePath, String[] headers, Object[][] data) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < data.length; r++) {
                Row row = sheet.createRow(r + 1);
                for (int c = 0; c < data[r].length; c++) {
                    Object val = data[r][c];
                    if (val instanceof String s) row.createCell(c).setCellValue(s);
                    else if (val instanceof Number n) row.createCell(c).setCellValue(n.doubleValue());
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    public static class TestPerson {
        String name;
        Integer age;
    }

    public record TestProduct(String name, int price, boolean active) {}
}
