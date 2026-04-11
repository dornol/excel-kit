package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Dedicated tests for {@link ExcelWriter.Builder} (v0.11.0).
 *
 * <p>Covers:
 * <ul>
 *   <li>Validation of each builder field (null color, non-positive maxRows, non-positive rowAccessWindowSize)</li>
 *   <li>Default values applied when nothing is set</li>
 *   <li>Each setting actually reaches the built writer (effect-level checks, not just "no throw")</li>
 *   <li>{@code builder()} returning fresh instances</li>
 * </ul>
 */
class ExcelWriterBuilderTest {

    @Nested
    @DisplayName("Builder validation")
    class Validation {

        @Test
        void color_null_throws() {
            var b = ExcelWriter.<String>builder();
            assertThrows(IllegalArgumentException.class, () -> b.color(null));
        }

        @Test
        void maxRows_zero_throws() {
            var b = ExcelWriter.<String>builder();
            assertThrows(IllegalArgumentException.class, () -> b.maxRows(0));
        }

        @Test
        void maxRows_negative_throws() {
            var b = ExcelWriter.<String>builder();
            assertThrows(IllegalArgumentException.class, () -> b.maxRows(-1));
        }

        @Test
        void rowAccessWindowSize_zero_throws() {
            var b = ExcelWriter.<String>builder();
            assertThrows(IllegalArgumentException.class, () -> b.rowAccessWindowSize(0));
        }

        @Test
        void rowAccessWindowSize_negative_throws() {
            var b = ExcelWriter.<String>builder();
            assertThrows(IllegalArgumentException.class, () -> b.rowAccessWindowSize(-1));
        }
    }

    @Nested
    @DisplayName("Builder defaults")
    class Defaults {

        @Test
        void noSettings_buildsWorkingWriter() throws IOException {
            var writer = ExcelWriter.<String>builder().build();
            var out = new ByteArrayOutputStream();
            writer.column("A", s -> s)
                    .write(Stream.of("x", "y", "z"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals("A", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
                assertEquals("x", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("z", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
            }
        }

        @Test
        void defaultMaxRows_isOneMillion() {
            // Can't practically write a million rows in a test, but we verify that the default
            // does NOT cause a premature rollover: 2 rows with a default writer should produce
            // exactly one sheet (not two).
            var writer = ExcelWriter.<Integer>builder().build();
            var out = new ByteArrayOutputStream();
            try {
                writer.column("n", i -> i)
                        .write(Stream.of(1, 2))
                        .write(out);
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    assertEquals(1, wb.getNumberOfSheets(),
                            "default maxRows should not trigger rollover for tiny datasets");
                }
            } catch (IOException e) {
                fail(e);
            }
        }
    }

    @Nested
    @DisplayName("Builder settings propagate to the built writer")
    class SettingsPropagate {

        @Test
        void maxRows_triggersSheetRollover() throws IOException {
            // Rollover formula: total >= maxRows && total % maxRows == 1
            // maxRows=2, 3 rows: row 3 triggers rollover → 2 sheets (sheet1: rows 1-2, sheet2: row 3).
            var writer = ExcelWriter.<String>builder().maxRows(2).build();
            var out = new ByteArrayOutputStream();
            writer.column("n", s -> s)
                    .write(Stream.of("row-1", "row-2", "row-3"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(2, wb.getNumberOfSheets(),
                        "maxRows=2 with 3 rows should produce 2 sheets (2+1 split)");
                // Sheet 1: header + rows 1, 2
                assertEquals("row-1", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("row-2", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
                // Sheet 2: header + row 3
                assertEquals("row-3", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void color_producesColoredHeader() throws IOException {
            var writer = ExcelWriter.<String>builder().color(ExcelColor.STEEL_BLUE).build();
            var out = new ByteArrayOutputStream();
            writer.column("A", s -> s)
                    .write(Stream.of("x"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                var headerStyle = sheet.getRow(0).getCell(0).getCellStyle();
                assertNotNull(headerStyle.getFillForegroundColorColor(),
                        "header cell should have a fill color applied via builder.color()");
            }
        }

        @Test
        void rowAccessWindowSize_doesNotCorruptOutput() throws IOException {
            // rowAccessWindowSize = 2 flushes rows aggressively. Verify the output is valid
            // and contains all rows when read back.
            var writer = ExcelWriter.<Integer>builder()
                    .rowAccessWindowSize(2)
                    .build();
            var out = new ByteArrayOutputStream();
            writer.column("n", i -> i)
                    .write(IntStream.rangeClosed(1, 10).boxed())
                    .write(out);

            assertTrue(out.size() > 0, "writer with small row window should still produce output");
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertEquals(10, sheet.getLastRowNum(),
                        "all 10 data rows should survive the SXSSF flush (rows 1..10, header at 0)");
            }
        }

        @Test
        void allThreeSettingsTogether() throws IOException {
            // maxRows=3, 4 rows: rollover at row 4 (total=4, 4>=3, 4%3=1)
            // → sheet1 has rows 1-3, sheet2 has row 4.
            var writer = ExcelWriter.<String>builder()
                    .color(ExcelColor.CORAL)
                    .maxRows(3)
                    .rowAccessWindowSize(5)
                    .build();
            var out = new ByteArrayOutputStream();
            writer.column("A", s -> s)
                    .write(Stream.of("a", "b", "c", "d"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(2, wb.getNumberOfSheets());
                // Sheet 1: header + a, b, c
                assertEquals("a", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("b", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
                assertEquals("c", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
                // Sheet 2: header + d
                assertEquals("d", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
            }
        }
    }

    @Nested
    @DisplayName("Builder lifecycle")
    class Lifecycle {

        @Test
        void builder_returnsFreshInstanceEachCall() {
            var a = ExcelWriter.<String>builder();
            var b = ExcelWriter.<String>builder();
            assertNotSame(a, b, "each builder() call should return a new Builder");
        }

        @Test
        void chaining_returnsSameBuilder() {
            var b = ExcelWriter.<String>builder();
            assertSame(b, b.color(ExcelColor.WHITE));
            assertSame(b, b.maxRows(100));
            assertSame(b, b.rowAccessWindowSize(50));
        }

        @Test
        void twoBuildsFromSameBuilder_produceSeparateWriters() throws IOException {
            var b = ExcelWriter.<String>builder().color(ExcelColor.GOLD).maxRows(1000);
            var w1 = b.build();
            var w2 = b.build();
            assertNotSame(w1, w2, "each build() call should produce a new ExcelWriter instance");

            // And both should be independently usable.
            var out1 = new ByteArrayOutputStream();
            var out2 = new ByteArrayOutputStream();
            w1.column("A", s -> s).write(Stream.of("x")).write(out1);
            w2.column("A", s -> s).write(Stream.of("y")).write(out2);

            try (var wb1 = new XSSFWorkbook(new ByteArrayInputStream(out1.toByteArray()));
                 var wb2 = new XSSFWorkbook(new ByteArrayInputStream(out2.toByteArray()))) {
                assertEquals("x", wb1.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("y", wb2.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }
    }
}
