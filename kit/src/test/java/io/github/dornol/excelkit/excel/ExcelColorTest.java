package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

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

    @Test
    void excelWriter_shouldAcceptExcelColor() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.STEEL_BLUE);
        Stream<String> data = Stream.of("a", "b");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_shouldAcceptCustomColor() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.of(100, 150, 200));
        Stream<String> data = Stream.of("a");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_shouldAcceptExcelColorWithMaxRowsAndWindowSize() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.LIGHT_BLUE, 1_000_000, 100);
        Stream<String> data = Stream.of("a");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_maxRowsOnly() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(500_000);
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(Stream.of("a"));
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_colorAndMaxRows() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.CORAL, 500_000);
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(Stream.of("a"));
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void backgroundColor_shouldAcceptExcelColor() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .backgroundColor(ExcelColor.LIGHT_YELLOW)
                .column("B", (row, c) -> row.length())
                .backgroundColor(ExcelColor.LIGHT_RED)
                .write(data);

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }
}
