package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
 * Tests for {@link ExcelWriter#create()} / {@link ExcelWriter#create(java.util.function.Consumer)}
 * and the fluent configuration setters that surround them
 * ({@link ExcelWriter#headerColor}, {@link ExcelWriter#maxRows}, and
 * {@link ExcelWriter.InitOptions#rowAccessWindowSize}).
 *
 * <p>Covers:
 * <ul>
 *   <li>Validation for each option (null/zero/negative)</li>
 *   <li>Defaults — verified at the boundary, not just "works"</li>
 *   <li>Effect-level propagation (actual RGB / row counts / rollover positions)</li>
 *   <li>Setter composition (headerColor after font settings, headerColor overriding itself)</li>
 *   <li>Return-this contracts for fluent chaining</li>
 * </ul>
 */
class ExcelWriterBuilderTest {

    // ------- helpers --------------------------------------------------------

    private static byte[] writeSingleString(ExcelWriter<String> writer) throws IOException {
        var out = new ByteArrayOutputStream();
        writer.column("A", s -> s).write(Stream.of("x")).writeTo(out);
        return out.toByteArray();
    }

    private static XSSFColor readHeaderFillColor(byte[] xlsx) throws IOException {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            XSSFCellStyle style = (XSSFCellStyle) wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
            return style.getFillForegroundColorColor();
        }
    }

    private static void assertRgb(XSSFColor color, int r, int g, int b) {
        assertNotNull(color, "expected a fill color to be present on header cell");
        byte[] rgb = color.getRGB();
        assertNotNull(rgb, "XSSFColor must expose RGB bytes");
        assertEquals(r & 0xFF, rgb[0] & 0xFF, "R mismatch");
        assertEquals(g & 0xFF, rgb[1] & 0xFF, "G mismatch");
        assertEquals(b & 0xFF, rgb[2] & 0xFF, "B mismatch");
    }

    // ------- validation -----------------------------------------------------

    @Nested
    @DisplayName("Validation")
    class Validation {

        @Test
        void headerColor_null_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create().headerColor(null));
        }

        @Test
        void maxRows_zero_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create().maxRows(0));
        }

        @Test
        void maxRows_negative_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create().maxRows(-1));
        }

        @Test
        void rowAccessWindowSize_zero_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(0)));
        }

        @Test
        void rowAccessWindowSize_negative_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(-1)));
        }

        @Test
        void maxRows_one_isAccepted() {
            // Boundary: 1 is the minimum legal positive value.
            assertDoesNotThrow(() -> ExcelWriter.<String>create().maxRows(1));
        }

        @Test
        void rowAccessWindowSize_one_isAccepted() {
            assertDoesNotThrow(() -> ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(1)));
        }
    }

    // ------- defaults -------------------------------------------------------

    @Nested
    @DisplayName("Defaults")
    class Defaults {

        @Test
        void noSettings_producesWhiteHeader() throws IOException {
            // Default header color is ExcelColor.WHITE (255, 255, 255).
            byte[] xlsx = writeSingleString(ExcelWriter.<String>create());
            assertRgb(readHeaderFillColor(xlsx), 255, 255, 255);
        }

        @Test
        void noSettings_writesAllData() throws IOException {
            var writer = ExcelWriter.<String>create();
            var out = new ByteArrayOutputStream();
            writer.column("A", s -> s)
                    .write(Stream.of("x", "y", "z"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("x", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("y", sheet.getRow(2).getCell(0).getStringCellValue());
                assertEquals("z", sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals(1, wb.getNumberOfSheets(),
                        "default maxRows should not trigger rollover for tiny datasets");
            }
        }
    }

    // ------- propagation ----------------------------------------------------

    @Nested
    @DisplayName("Settings propagate to output")
    class SettingsPropagate {

        @Test
        void maxRows_rollsOverAtExactBoundary() throws IOException {
            // Rollover formula: total >= maxRows && total % maxRows == 1.
            // maxRows=2, 3 rows → row 3 triggers rollover (3>=2, 3%2==1).
            // Result: sheet 1 has rows 1-2, sheet 2 has row 3.
            var writer = ExcelWriter.<String>create().maxRows(2);
            var out = new ByteArrayOutputStream();
            writer.column("n", s -> s)
                    .write(Stream.of("row-1", "row-2", "row-3"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(2, wb.getNumberOfSheets(),
                        "maxRows=2 with 3 rows should produce 2 sheets");
                assertEquals("row-1", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("row-2", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
                assertNull(wb.getSheetAt(0).getRow(3),
                        "sheet 1 should have exactly 2 data rows (no row-3)");
                assertEquals("row-3", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        void maxRows_noRolloverWhenBelowLimit() throws IOException {
            // maxRows=5, only 3 rows → single sheet.
            var writer = ExcelWriter.<String>create().maxRows(5);
            var out = new ByteArrayOutputStream();
            writer.column("n", s -> s)
                    .write(Stream.of("a", "b", "c"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(1, wb.getNumberOfSheets());
                assertEquals(3, wb.getSheetAt(0).getLastRowNum(), "header + 3 rows = last index 3");
            }
        }

        @Test
        void headerColor_producesExactRgb_steelBlue() throws IOException {
            byte[] xlsx = writeSingleString(
                    ExcelWriter.<String>create().headerColor(ExcelColor.STEEL_BLUE));
            assertRgb(readHeaderFillColor(xlsx),
                    ExcelColor.STEEL_BLUE.getR(),
                    ExcelColor.STEEL_BLUE.getG(),
                    ExcelColor.STEEL_BLUE.getB());
        }

        @Test
        void headerColor_producesExactRgb_custom() throws IOException {
            byte[] xlsx = writeSingleString(
                    ExcelWriter.<String>create().headerColor(ExcelColor.of(10, 20, 30)));
            assertRgb(readHeaderFillColor(xlsx), 10, 20, 30);
        }

        @Test
        void rowAccessWindowSize_preservesAllRowsAndValues() throws IOException {
            // rowAccessWindowSize=2 forces aggressive flushing. All rows must survive AND
            // retain their original values (not just "some bytes produced").
            var writer = ExcelWriter.<Integer>create(opts -> opts.rowAccessWindowSize(2));
            var out = new ByteArrayOutputStream();
            writer.column("n", i -> i, c -> c.type(ExcelDataType.INTEGER))
                    .write(IntStream.rangeClosed(1, 10).boxed())
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertEquals(10, sheet.getLastRowNum(), "all 10 data rows should survive flush");
                for (int i = 1; i <= 10; i++) {
                    assertEquals(i, (int) sheet.getRow(i).getCell(0).getNumericCellValue(),
                            "row " + i + " value must match original after SXSSF flush");
                }
            }
        }

        @Test
        void allSettingsTogether_composeCorrectly() throws IOException {
            // rowAccessWindowSize(5) + headerColor(CORAL) + maxRows(3), 4 rows.
            // Rollover at row 4 (4>=3, 4%3==1). Sheet 1: a,b,c. Sheet 2: d.
            // Both sheets must have CORAL header.
            var writer = ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(5))
                    .headerColor(ExcelColor.CORAL)
                    .maxRows(3);
            var out = new ByteArrayOutputStream();
            writer.column("A", s -> s)
                    .write(Stream.of("a", "b", "c", "d"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals("a", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("b", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
                assertEquals("c", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
                assertEquals("d", wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());

                // Every rollover sheet should inherit the chosen header color.
                for (int i = 0; i < 2; i++) {
                    XSSFCellStyle style = (XSSFCellStyle)
                            wb.getSheetAt(i).getRow(0).getCell(0).getCellStyle();
                    XSSFColor c = style.getFillForegroundColorColor();
                    byte[] rgb = c.getRGB();
                    assertEquals(ExcelColor.CORAL.getR(), rgb[0] & 0xFF, "sheet " + i + " R");
                    assertEquals(ExcelColor.CORAL.getG(), rgb[1] & 0xFF, "sheet " + i + " G");
                    assertEquals(ExcelColor.CORAL.getB(), rgb[2] & 0xFF, "sheet " + i + " B");
                }
            }
        }
    }

    // ------- setter composition --------------------------------------------

    @Nested
    @DisplayName("headerColor setter composition")
    class HeaderColorComposition {

        @Test
        void lastHeaderColorWins() throws IOException {
            // Calling headerColor twice: the last call must be the one that ends up in the output.
            byte[] xlsx = writeSingleString(
                    ExcelWriter.<String>create()
                            .headerColor(ExcelColor.STEEL_BLUE)
                            .headerColor(ExcelColor.CORAL));
            assertRgb(readHeaderFillColor(xlsx),
                    ExcelColor.CORAL.getR(),
                    ExcelColor.CORAL.getG(),
                    ExcelColor.CORAL.getB());
        }

        @Test
        void headerColor_afterFontSettings_preservesFont() throws IOException {
            // Regression guard: headerColor() rebuilds headerStyle internally.
            // Font settings set BEFORE headerColor must survive that rebuild.
            var writer = ExcelWriter.<String>create()
                    .headerFontName("Courier New")
                    .headerFontSize(14)
                    .headerColor(ExcelColor.CORAL);

            byte[] xlsx = writeSingleString(writer);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
                XSSFCellStyle style = (XSSFCellStyle)
                        wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
                XSSFFont font = style.getFont();
                assertEquals("Courier New", font.getFontName(),
                        "font name set before headerColor must survive headerStyle rebuild");
                assertEquals(14, font.getFontHeightInPoints(),
                        "font size set before headerColor must survive headerStyle rebuild");

                // Color must also be correct (primary effect of headerColor).
                byte[] rgb = style.getFillForegroundColorColor().getRGB();
                assertEquals(ExcelColor.CORAL.getR(), rgb[0] & 0xFF);
                assertEquals(ExcelColor.CORAL.getG(), rgb[1] & 0xFF);
                assertEquals(ExcelColor.CORAL.getB(), rgb[2] & 0xFF);
            }
        }

        @Test
        void headerColor_beforeFontSettings_fontStillApplied() throws IOException {
            // write()'s lazy rebuild (when fonts are non-null) must NOT clobber
            // the color that was already set earlier.
            var writer = ExcelWriter.<String>create()
                    .headerColor(ExcelColor.CORAL)
                    .headerFontName("Courier New");

            byte[] xlsx = writeSingleString(writer);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
                XSSFCellStyle style = (XSSFCellStyle)
                        wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
                assertEquals("Courier New", style.getFont().getFontName());
                byte[] rgb = style.getFillForegroundColorColor().getRGB();
                assertEquals(ExcelColor.CORAL.getR(), rgb[0] & 0xFF,
                        "headerColor set BEFORE font must still be visible after write() font rebuild");
                assertEquals(ExcelColor.CORAL.getG(), rgb[1] & 0xFF);
                assertEquals(ExcelColor.CORAL.getB(), rgb[2] & 0xFF);
            }
        }
    }

    // ------- fluent lifecycle ----------------------------------------------

    @Nested
    @DisplayName("create() / fluent lifecycle")
    class Lifecycle {

        @Test
        void create_noArgs_returnsFreshInstanceEachCall() {
            var a = ExcelWriter.<String>create();
            var b = ExcelWriter.<String>create();
            assertNotSame(a, b);
        }

        @Test
        void create_withConsumer_returnsFreshInstanceEachCall() {
            var a = ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(10));
            var b = ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(10));
            assertNotSame(a, b);
        }

        @Test
        void initOptions_rowAccessWindowSize_returnsSameOptions() {
            ExcelWriter.<String>create(opts ->
                    assertSame(opts, opts.rowAccessWindowSize(50)));
        }

        @Test
        void fluentSetters_returnSameWriter() {
            var w = ExcelWriter.<String>create();
            assertSame(w, w.headerColor(ExcelColor.WHITE));
            assertSame(w, w.maxRows(100));
            assertSame(w, w.headerFontName("Arial"));
            assertSame(w, w.headerFontSize(12));
        }

        @Test
        void create_emptyConsumer_equivalentToNoArgsCreate() throws IOException {
            // Both forms must produce the same default output.
            byte[] a = writeSingleString(ExcelWriter.<String>create());
            byte[] b = writeSingleString(ExcelWriter.<String>create(opts -> {}));
            // Byte-level equality is too strict (timestamps etc). Compare default color effect.
            assertRgb(readHeaderFillColor(a), 255, 255, 255);
            assertRgb(readHeaderFillColor(b), 255, 255, 255);
        }
    }
}
