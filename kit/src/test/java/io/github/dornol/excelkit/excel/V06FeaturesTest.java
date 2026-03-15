package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvMapWriter;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive tests for all v0.6 features:
 * 1. Border Style
 * 2. Multi-Sheet Reading
 * 3. Map-based I/O
 * 5. Cell Comment
 * 7. Conditional Formatting
 * 9. Sheet/Cell Protection
 * 10. Image
 * 12. Chart
 */
class V06FeaturesTest {

    // ============================================================
    // Feature 1: Border Style - additional edge cases
    // ============================================================
    @Nested
    class BorderTests {
        @Test
        void border_allStyles_shouldCompileAndWrite() throws IOException {
            for (ExcelBorderStyle style : ExcelBorderStyle.values()) {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                new ExcelWriter<String>()
                        .addColumn("Col", s -> s, c -> c.border(style))
                        .write(Stream.of("data"))
                        .consumeOutputStream(out);

                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    var cellStyle = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                    assertEquals(style.toPoiBorderStyle(), cellStyle.getBorderTop(),
                            "Border style " + style + " should be applied");
                }
            }
        }

        @Test
        void border_mixedStyles_perColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Thin", s -> s, c -> c.border(ExcelBorderStyle.THIN))
                    .addColumn("Thick", s -> s, c -> c.border(ExcelBorderStyle.THICK))
                    .addColumn("None", s -> s, c -> c.border(ExcelBorderStyle.NONE))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertEquals(BorderStyle.THIN, row.getCell(0).getCellStyle().getBorderTop());
                assertEquals(BorderStyle.THICK, row.getCell(1).getCellStyle().getBorderTop());
                assertEquals(BorderStyle.NONE, row.getCell(2).getCellStyle().getBorderTop());
            }
        }

        @Test
        void border_withOtherStyling_shouldCombine() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s, c -> c
                            .border(ExcelBorderStyle.MEDIUM)
                            .bold(true)
                            .fontSize(14)
                            .backgroundColor(ExcelColor.LIGHT_BLUE))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.MEDIUM, style.getBorderTop());
            }
        }
    }

    // ============================================================
    // Feature 2: Multi-Sheet Reading - additional edge cases
    // ============================================================
    @Nested
    class MultiSheetReadTests {
        @Test
        void getSheetNames_rolloverSheets() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>(3)
                    .sheetName("Data")
                    .addColumn("Name", s -> s)
                    .write(Stream.of("A", "B", "C", "D", "E", "F", "G", "H", "I"))
                    .consumeOutputStream(out);

            var sheets = ExcelReader.getSheetNames(new ByteArrayInputStream(out.toByteArray()));
            assertTrue(sheets.size() >= 3, "Should have multiple sheets");
            assertEquals("Data", sheets.get(0).name());
        }

        @Test
        void getSheetHeaders_emptySheet() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s)
                    .addColumn("B", s -> s)
                    .write(Stream.empty())
                    .consumeOutputStream(out);

            var headers = ExcelReader.getSheetHeaders(
                    new ByteArrayInputStream(out.toByteArray()), 0, 0);
            assertEquals(List.of("A", "B"), headers);
        }

        @Test
        void getSheetNames_exception_invalidStream() {
            assertThrows(ExcelReadException.class, () ->
                    ExcelReader.getSheetNames(new ByteArrayInputStream(new byte[]{1, 2, 3})));
        }

        @Test
        void getSheetHeaders_exception_invalidStream() {
            assertThrows(ExcelReadException.class, () ->
                    ExcelReader.getSheetHeaders(new ByteArrayInputStream(new byte[]{1, 2, 3}), 0, 0));
        }

        @Test
        void readAllSheets_iterateByName() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Users")
                        .column("Name", s -> s)
                        .write(Stream.of("Alice", "Bob"));
                wb.<String>sheet("Items")
                        .column("Item", s -> s)
                        .write(Stream.of("Widget"));
                wb.finish().consumeOutputStream(out);
            }

            byte[] data = out.toByteArray();
            var sheets = ExcelReader.getSheetNames(new ByteArrayInputStream(data));

            // Read each sheet by name
            Map<String, List<String>> allData = new LinkedHashMap<>();
            for (ExcelSheetInfo info : sheets) {
                List<String> values = new ArrayList<>();
                new ExcelReader<>(Holder::new, null)
                        .sheetIndex(info.index())
                        .addColumn((h, c) -> h.value = c.asString())
                        .build(new ByteArrayInputStream(data))
                        .read(r -> values.add(r.data().value));
                allData.put(info.name(), values);
            }

            assertEquals(List.of("Alice", "Bob"), allData.get("Users"));
            assertEquals(List.of("Widget"), allData.get("Items"));
        }
    }

    // ============================================================
    // Feature 3: Map-based I/O - comprehensive tests
    // ============================================================
    @Nested
    class MapIOTests {
        @Test
        void excelMapWriter_withConfigurers() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            @SuppressWarnings("unchecked")
            var writer = new ExcelMapWriter(
                    new ExcelWriter<>(),
                    new String[]{"Name", "Price"},
                    c -> c.bold(true),
                    c -> c.type(ExcelDataType.INTEGER)
            );
            writer.write(Stream.of(
                    Map.of("Name", "Item", "Price", 100)
            )).consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals("Item", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void excelMapWriter_emptyStream() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelMapWriter("A", "B")
                    .write(Stream.empty())
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals("A", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
                assertNull(wb.getSheetAt(0).getRow(1)); // no data rows
            }
        }

        @Test
        void excelMapWriter_nullValuesInMap() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            Map<String, Object> row = new HashMap<>();
            row.put("Name", null);
            row.put("Age", null);

            new ExcelMapWriter("Name", "Age")
                    .write(Stream.of(row))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // null values should result in empty cells
                assertEquals("", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void excelMapWriter_extraKeysIgnored() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelMapWriter("Name")
                    .write(Stream.of(Map.of("Name", "Alice", "Extra", "ignored")))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(1, wb.getSheetAt(0).getRow(0).getLastCellNum()); // only 1 column
            }
        }

        @Test
        void excelMapReader_emptyDataRows() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelMapWriter("Name", "Age")
                    .write(Stream.empty())
                    .consumeOutputStream(out);

            List<Map<String, String>> results = new ArrayList<>();
            new ExcelMapReader()
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> results.add(r.data()));

            assertTrue(results.isEmpty());
        }

        @Test
        void excelMapReader_sparseData() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            Map<String, Object> fullRow = new LinkedHashMap<>();
            fullRow.put("A", "val1");
            fullRow.put("B", "val2");
            fullRow.put("C", "val3");

            Map<String, Object> sparseRow = new LinkedHashMap<>();
            sparseRow.put("A", "only-a");
            sparseRow.put("B", null);
            sparseRow.put("C", null);

            new ExcelMapWriter("A", "B", "C")
                    .write(Stream.of(fullRow, sparseRow))
                    .consumeOutputStream(out);

            List<Map<String, String>> results = new ArrayList<>();
            new ExcelMapReader()
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
            assertEquals("val1", results.get(0).get("A"));
            assertEquals("val2", results.get(0).get("B"));
            assertEquals("only-a", results.get(1).get("A"));
        }

        @Test
        void excelMapReader_headerRowIndex() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .beforeHeader(ctx -> {
                        ctx.getSheet().createRow(0).createCell(0).setCellValue("Title");
                        return 1;
                    })
                    .addColumn("Name", s -> s)
                    .write(Stream.of("Alice"))
                    .consumeOutputStream(out);

            List<Map<String, String>> results = new ArrayList<>();
            new ExcelMapReader()
                    .headerRowIndex(1)
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
        }

        @Test
        void csvMapWriter_specialCharacters() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvMapWriter("Name", "Desc")
                    .write(Stream.of(Map.of("Name", "Alice, Jr.", "Desc", "She said \"hi\"")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8);
            assertTrue(csv.contains("Alice, Jr.") || csv.contains("\"Alice, Jr.\""));
        }

        @Test
        void csvMapWriter_withCustomWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            var csvWriter = new CsvWriter<Map<String, Object>>();
            csvWriter.delimiter('\t');
            new CsvMapWriter(csvWriter, "A", "B")
                    .write(Stream.of(Map.of("A", "1", "B", "2")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8);
            assertTrue(csv.contains("A\tB"));
        }

        @Test
        void csvMapWriter_nullValues() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            Map<String, Object> row = new HashMap<>();
            row.put("Name", "Alice");
            row.put("Value", null);

            new CsvMapWriter("Name", "Value")
                    .write(Stream.of(row))
                    .consumeOutputStream(out);

            assertFalse(out.toByteArray().length == 0);
        }
    }

    // ============================================================
    // Feature 5: Cell Comment - additional tests
    // ============================================================
    @Nested
    class CommentTests {
        @Test
        void comment_multipleColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s, c -> c.comment(s -> "Comment for A: " + s))
                    .addColumn("B", s -> s.toUpperCase(), c -> c.comment(s -> "Comment for B: " + s))
                    .write(Stream.of("hello"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertNotNull(row.getCell(0).getCellComment());
                assertNotNull(row.getCell(1).getCellComment());
                assertEquals("Comment for A: hello",
                        row.getCell(0).getCellComment().getString().getString());
                assertEquals("Comment for B: hello",
                        row.getCell(1).getCellComment().getString().getString());
            }
        }

        @Test
        void comment_emptyString() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s, c -> c.comment(s -> ""))
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // Empty string comment should still be added
                assertNotNull(wb.getSheetAt(0).getRow(1).getCell(0).getCellComment());
            }
        }

        @Test
        void comment_multipleRows() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Name", s -> s, c -> c.comment(s -> "Hi " + s))
                    .write(Stream.of("Alice", "Bob", "Charlie"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                for (int i = 1; i <= 3; i++) {
                    assertNotNull(wb.getSheetAt(0).getRow(i).getCell(0).getCellComment());
                }
            }
        }
    }

    // ============================================================
    // Feature 7: Conditional Formatting - edge cases
    // ============================================================
    @Nested
    class ConditionalFormattingTests {
        @Test
        void conditionalFormatting_noRules_shouldNotFail() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Name", s -> s)
                    .conditionalFormatting(cf -> {})  // no rules added
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(0, wb.getSheetAt(0).getSheetConditionalFormatting()
                        .getNumConditionalFormattings());
            }
        }

        @Test
        void conditionalFormatting_allColumnsByDefault() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s)
                    .addColumn("B", s -> s)
                    .conditionalFormatting(cf -> cf
                            .equalTo("\"test\"", ExcelColor.LIGHT_RED))
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // Rules applied to all columns
                assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                        .getNumConditionalFormattings() > 0);
            }
        }

        @Test
        void conditionalFormatting_multipleRulesSets() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Val", s -> s)
                    .conditionalFormatting(cf -> cf
                            .columns(0)
                            .greaterThan("100", ExcelColor.RED))
                    .conditionalFormatting(cf -> cf
                            .columns(0)
                            .lessThan("0", ExcelColor.BLUE))
                    .write(Stream.of("50"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                        .getNumConditionalFormattings() >= 2);
            }
        }
    }

    // ============================================================
    // Feature 9: Sheet/Cell Protection - edge cases
    // ============================================================
    @Nested
    class ProtectionTests {
        @Test
        void protectSheet_allColumnsLockedByDefault() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("A", s -> s)
                    .addColumn("B", s -> s)
                    .protectSheet("pass")
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertNotNull(wb.getSheetAt(0).getCTWorksheet().getSheetProtection());
            }
        }

        @Test
        void protectSheet_withRollover() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>(3)
                    .addColumn("Name", s -> s)
                    .protectSheet("pass")
                    .write(Stream.of("A", "B", "C", "D", "E", "F", "G"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    assertNotNull(wb.getSheetAt(i).getCTWorksheet().getSheetProtection());
                }
            }
        }

        @Test
        void locked_withDifferentStyles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Editable", s -> s, c -> c.locked(false).backgroundColor(ExcelColor.LIGHT_GREEN))
                    .addColumn("ReadOnly", s -> s, c -> c.locked(true).bold(true))
                    .protectSheet("pass")
                    .write(Stream.of("data"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertFalse(row.getCell(0).getCellStyle().getLocked());
                assertTrue(row.getCell(1).getCellStyle().getLocked());
            }
        }
    }

    // ============================================================
    // Feature 10: Image - edge cases
    // ============================================================
    @Nested
    class ImageTests {
        // Minimal valid 1x1 PNG
        private static final byte[] TINY_PNG = {
                (byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                0x08, 0x02, 0x00, 0x00, 0x00, (byte) 0x90, 0x77, 0x53,
                (byte) 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                0x54, 0x08, (byte) 0xD7, 0x63, (byte) 0xF8, (byte) 0xCF,
                (byte) 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
                (byte) 0xE2, 0x21, (byte) 0xBC, 0x33, 0x00, 0x00, 0x00,
                0x00, 0x49, 0x45, 0x4E, 0x44, (byte) 0xAE, 0x42, 0x60,
                (byte) 0x82
        };

        @Test
        void image_multipleImages() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Name", s -> s)
                    .addColumn("Pic", s -> ExcelImage.png(TINY_PNG), c -> c.type(ExcelDataType.IMAGE))
                    .write(Stream.of("A", "B", "C"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(3, wb.getAllPictures().size());
            }
        }

        @Test
        void image_withTextColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("ID", s -> s)
                    .addColumn("Name", s -> "Item-" + s)
                    .addColumn("Photo", s -> ExcelImage.png(TINY_PNG), c -> c.type(ExcelDataType.IMAGE))
                    .addColumn("Notes", s -> "Note-" + s)
                    .write(Stream.of("1"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals("1", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("Item-1", wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
                assertEquals("Note-1", wb.getSheetAt(0).getRow(1).getCell(3).getStringCellValue());
                assertEquals(1, wb.getAllPictures().size());
            }
        }

        @Test
        void excelImage_record_equality() {
            byte[] data = {1, 2, 3};
            ExcelImage img1 = ExcelImage.png(data);
            ExcelImage img2 = new ExcelImage(data, Workbook.PICTURE_TYPE_PNG);
            assertEquals(img1.imageType(), img2.imageType());
        }
    }

    // ============================================================
    // Feature 12: Chart - edge cases
    // ============================================================
    @Nested
    class ChartTests {
        record Data(String name, int value, int extra) {}

        @Test
        void chart_withNoTitle() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Data>()
                    .addColumn("Name", Data::name)
                    .addColumn("Value", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                    .chart(ch -> ch
                            .type(ExcelChartConfig.ChartType.BAR)
                            .categoryColumn(0)
                            .valueColumn(1, "Values"))
                    .write(Stream.of(new Data("A", 10, 0)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
            }
        }

        @Test
        void chart_defaultPosition() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Data>()
                    .addColumn("Name", Data::name)
                    .addColumn("Value", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                    .chart(ch -> ch
                            .type(ExcelChartConfig.ChartType.LINE)
                            .title("Auto Position")
                            .categoryColumn(0)
                            .valueColumn(1, "Values"))
                    .write(Stream.of(new Data("A", 10, 0), new Data("B", 20, 0)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
            }
        }

        @Test
        void chart_nonZeroCategoryColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Data>()
                    .addColumn("ID", d -> d.value, c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Name", Data::name)
                    .addColumn("Value", d -> d.extra, c -> c.type(ExcelDataType.INTEGER))
                    .chart(ch -> ch
                            .type(ExcelChartConfig.ChartType.PIE)
                            .categoryColumn(1)  // use Name as category
                            .valueColumn(2, "Values"))
                    .write(Stream.of(new Data("A", 1, 60), new Data("B", 2, 40)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertFalse(wb.getSheetAt(0).getDrawingPatriarch().getCharts().isEmpty());
            }
        }

        @Test
        void chartType_enum_coverage() {
            assertEquals(3, ExcelChartConfig.ChartType.values().length);
            assertNotNull(ExcelChartConfig.ChartType.valueOf("BAR"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("LINE"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("PIE"));
        }
    }

    // ============================================================
    // Feature combination tests
    // ============================================================
    @Nested
    class CombinationTests {
        @Test
        void allFeatures_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Name", s -> s, c -> c
                            .border(ExcelBorderStyle.MEDIUM)
                            .comment(s -> "Note: " + s)
                            .locked(false))
                    .addColumn("Value", s -> "100", c -> c
                            .type(ExcelDataType.INTEGER)
                            .border(ExcelBorderStyle.THICK))
                    .protectSheet("pass")
                    .conditionalFormatting(cf -> cf
                            .columns(1)
                            .greaterThan("50", ExcelColor.LIGHT_GREEN))
                    .write(Stream.of("Alice", "Bob"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Protection
                assertNotNull(sheet.getCTWorksheet().getSheetProtection());
                // Comments
                assertNotNull(sheet.getRow(1).getCell(0).getCellComment());
                // Border
                assertEquals(BorderStyle.MEDIUM,
                        sheet.getRow(1).getCell(0).getCellStyle().getBorderTop());
                // Conditional formatting
                assertTrue(sheet.getSheetConditionalFormatting()
                        .getNumConditionalFormattings() > 0);
                // Locked
                assertFalse(sheet.getRow(1).getCell(0).getCellStyle().getLocked());
            }
        }

        @Test
        void mapWriter_withProtection() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            var writer = new ExcelMapWriter("Name", "Age");
            writer.writer().protectSheet("secret");
            writer.write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertNotNull(wb.getSheetAt(0).getCTWorksheet().getSheetProtection());
            }
        }

        @Test
        void sheetWriter_allNewFeatures() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Protected")
                        .column("Name", s -> s, c -> c
                                .border(ExcelBorderStyle.DASHED)
                                .comment(s -> "Note")
                                .locked(false))
                        .protectSheet("pass")
                        .conditionalFormatting(cf -> cf
                                .equalTo("\"Alice\"", ExcelColor.LIGHT_GREEN))
                        .write(Stream.of("Alice"));
                wb.finish().consumeOutputStream(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertNotNull(sheet.getCTWorksheet().getSheetProtection());
                assertNotNull(sheet.getRow(1).getCell(0).getCellComment());
                assertEquals(BorderStyle.DASHED,
                        sheet.getRow(1).getCell(0).getCellStyle().getBorderTop());
            }
        }
    }

    // ============================================================
    // ExcelSheetInfo record tests
    // ============================================================
    @Test
    void excelSheetInfo_record() {
        var info = new ExcelSheetInfo(2, "Sheet3");
        assertEquals(2, info.index());
        assertEquals("Sheet3", info.name());
    }

    public static class Holder {
        String value;
    }
}
