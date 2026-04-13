package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelColor}
 */
class ExcelColorTest {

    @Test
    void presetColors_shouldHaveValidRgbRange() {
        ExcelColor[] presets = {
                ExcelColor.WHITE, ExcelColor.BLACK,
                ExcelColor.LIGHT_GRAY, ExcelColor.GRAY, ExcelColor.DARK_GRAY,
                ExcelColor.RED, ExcelColor.GREEN, ExcelColor.BLUE, ExcelColor.YELLOW, ExcelColor.ORANGE,
                ExcelColor.LIGHT_RED, ExcelColor.LIGHT_GREEN, ExcelColor.LIGHT_BLUE,
                ExcelColor.LIGHT_YELLOW, ExcelColor.LIGHT_ORANGE, ExcelColor.LIGHT_PURPLE,
                ExcelColor.PURPLE, ExcelColor.PINK, ExcelColor.TEAL, ExcelColor.NAVY,
                ExcelColor.CORAL, ExcelColor.STEEL_BLUE, ExcelColor.FOREST_GREEN, ExcelColor.GOLD
        };
        for (ExcelColor color : presets) {
            assertTrue(color.getR() >= 0 && color.getR() <= 255, "R out of range: " + color.getR());
            assertTrue(color.getG() >= 0 && color.getG() <= 255, "G out of range: " + color.getG());
            assertTrue(color.getB() >= 0 && color.getB() <= 255, "B out of range: " + color.getB());
        }
    }

    @Test
    void of_shouldCreateCustomColor() {
        ExcelColor custom = ExcelColor.of(180, 200, 220);
        assertEquals(180, custom.getR());
        assertEquals(200, custom.getG());
        assertEquals(220, custom.getB());
    }

    @Test
    void of_negativeValue_throws() {
        var ex = assertThrows(IllegalArgumentException.class, () -> ExcelColor.of(-1, 0, 0));
        assertTrue(ex.getMessage().contains("-1"));
    }

    @Test
    void of_overMaxValue_throws() {
        var ex = assertThrows(IllegalArgumentException.class, () -> ExcelColor.of(0, 256, 0));
        assertTrue(ex.getMessage().contains("256"));
    }

    @Test
    void of_boundaryValues() {
        ExcelColor min = ExcelColor.of(0, 0, 0);
        assertEquals(0, min.getR());
        assertEquals(0, min.getG());
        assertEquals(0, min.getB());

        ExcelColor max = ExcelColor.of(255, 255, 255);
        assertEquals(255, max.getR());
        assertEquals(255, max.getG());
        assertEquals(255, max.getB());
    }

    @Test
    void toRgb_preset_shouldReturnCorrectArray() {
        int[] rgb = ExcelColor.BLUE.toRgb();
        assertArrayEquals(new int[]{0, 0, 255}, rgb);
    }

    @Test
    void toRgb_custom_shouldReturnCorrectArray() {
        int[] rgb = ExcelColor.of(10, 20, 30).toRgb();
        assertArrayEquals(new int[]{10, 20, 30}, rgb);
    }

    // ---- helpers -----------------------------------------------------------

    private static byte[] writeOne(ExcelWriter<String> writer) throws IOException {
        var out = new ByteArrayOutputStream();
        writer.column("A", (row, c) -> row).write(Stream.of("a")).write(out);
        return out.toByteArray();
    }

    private static void assertHeaderRgb(byte[] xlsx, int r, int g, int b) throws IOException {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            XSSFCellStyle style = (XSSFCellStyle) wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
            XSSFColor color = style.getFillForegroundColorColor();
            assertNotNull(color, "header cell should have a fill color");
            byte[] rgb = color.getRGB();
            assertEquals(r & 0xFF, rgb[0] & 0xFF, "R");
            assertEquals(g & 0xFF, rgb[1] & 0xFF, "G");
            assertEquals(b & 0xFF, rgb[2] & 0xFF, "B");
        }
    }

    // ---- ExcelWriter headerColor integration -------------------------------

    @Test
    void headerColor_steelBluePreset_producesExactRgb() throws IOException {
        byte[] xlsx = writeOne(ExcelWriter.<String>create().headerColor(ExcelColor.STEEL_BLUE));
        assertHeaderRgb(xlsx,
                ExcelColor.STEEL_BLUE.getR(),
                ExcelColor.STEEL_BLUE.getG(),
                ExcelColor.STEEL_BLUE.getB());
    }

    @Test
    void headerColor_customRgb_producesExactRgb() throws IOException {
        byte[] xlsx = writeOne(ExcelWriter.<String>create().headerColor(ExcelColor.of(100, 150, 200)));
        assertHeaderRgb(xlsx, 100, 150, 200);
    }

    @Test
    void headerColor_worksWithRowAccessWindowAndMaxRows() throws IOException {
        byte[] xlsx = writeOne(
                ExcelWriter.<String>create(opts -> opts.rowAccessWindowSize(100))
                        .headerColor(ExcelColor.LIGHT_BLUE)
                        .maxRows(1_000_000));
        assertHeaderRgb(xlsx,
                ExcelColor.LIGHT_BLUE.getR(),
                ExcelColor.LIGHT_BLUE.getG(),
                ExcelColor.LIGHT_BLUE.getB());
    }

    @Test
    void maxRowsOnly_keepsDefaultWhiteHeader() throws IOException {
        // maxRows fluent setter must not affect header color — default remains white.
        byte[] xlsx = writeOne(ExcelWriter.<String>create().maxRows(500_000));
        assertHeaderRgb(xlsx, 255, 255, 255);
    }

    @Test
    void headerColorAndMaxRows_composeIndependently() throws IOException {
        byte[] xlsx = writeOne(
                ExcelWriter.<String>create().headerColor(ExcelColor.CORAL).maxRows(500_000));
        assertHeaderRgb(xlsx,
                ExcelColor.CORAL.getR(),
                ExcelColor.CORAL.getG(),
                ExcelColor.CORAL.getB());
    }

    @Test
    void columnBackgroundColor_appliedPerColumn() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        var out = new ByteArrayOutputStream();
        writer.column("A", (row, c) -> row, cfg -> cfg.backgroundColor(ExcelColor.LIGHT_YELLOW))
                .column("B", (row, c) -> row.length(), cfg -> cfg.backgroundColor(ExcelColor.LIGHT_RED))
                .write(Stream.of("a", "b"))
                .write(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // Data cell (row 1, not header row 0) should carry per-column bg color.
            XSSFCellStyle colA = (XSSFCellStyle) wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            XSSFCellStyle colB = (XSSFCellStyle) wb.getSheetAt(0).getRow(1).getCell(1).getCellStyle();
            byte[] rgbA = colA.getFillForegroundColorColor().getRGB();
            byte[] rgbB = colB.getFillForegroundColorColor().getRGB();
            assertEquals(ExcelColor.LIGHT_YELLOW.getR(), rgbA[0] & 0xFF);
            assertEquals(ExcelColor.LIGHT_YELLOW.getG(), rgbA[1] & 0xFF);
            assertEquals(ExcelColor.LIGHT_YELLOW.getB(), rgbA[2] & 0xFF);
            assertEquals(ExcelColor.LIGHT_RED.getR(), rgbB[0] & 0xFF);
            assertEquals(ExcelColor.LIGHT_RED.getG(), rgbB[1] & 0xFF);
            assertEquals(ExcelColor.LIGHT_RED.getB(), rgbB[2] & 0xFF);
        }
    }
}
