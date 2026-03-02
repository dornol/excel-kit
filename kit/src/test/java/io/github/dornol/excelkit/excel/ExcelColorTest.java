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
    void allColors_shouldHaveValidRgbRange() {
        for (ExcelColor color : ExcelColor.values()) {
            assertTrue(color.getR() >= 0 && color.getR() <= 255,
                    color.name() + " R out of range: " + color.getR());
            assertTrue(color.getG() >= 0 && color.getG() <= 255,
                    color.name() + " G out of range: " + color.getG());
            assertTrue(color.getB() >= 0 && color.getB() <= 255,
                    color.name() + " B out of range: " + color.getB());
        }
    }

    @Test
    void toRgb_shouldReturnCorrectArray() {
        int[] rgb = ExcelColor.BLUE.toRgb();
        assertEquals(3, rgb.length);
        assertEquals(0, rgb[0]);
        assertEquals(0, rgb[1]);
        assertEquals(255, rgb[2]);
    }

    @Test
    void excelWriter_shouldAcceptExcelColor() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.STEEL_BLUE);
        Stream<String> data = Stream.of("a", "b");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void excelWriter_shouldAcceptExcelColorWithMaxRows() throws IOException {
        ExcelWriter<String> writer = new ExcelWriter<>(ExcelColor.DARK_GRAY, 1_000_000);
        Stream<String> data = Stream.of("a");

        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        assertNotNull(handler);
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

        assertNotNull(handler);
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

        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }
}
