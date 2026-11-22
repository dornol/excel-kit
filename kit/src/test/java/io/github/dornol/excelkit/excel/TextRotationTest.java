package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class TextRotationTest {

    @Test
    void rotation_positive45_appliesCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Rotated", s -> s, c -> c.rotation(45))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(45, style.getRotation(), "Rotation should be 45 degrees");
        }
    }

    @Test
    void rotation_negative45_convertsToPOIFormat() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Rotated", s -> s, c -> c.rotation(-45))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            // POI: -45 degrees → 90 + 45 = 135
            assertEquals(135, style.getRotation(), "Negative rotation should be converted to POI format");
        }
    }

    @Test
    void rotation_zero_noRotation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("NoRotation", s -> s, c -> c.rotation(0))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(0, style.getRotation(), "Rotation should be 0");
        }
    }

    @Test
    void rotation_90_vertical() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Vertical", s -> s, c -> c.rotation(90))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(90, style.getRotation());
        }
    }

    @Test
    void rotation_negative90_vertical() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Vertical", s -> s, c -> c.rotation(-90))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(180, style.getRotation());
        }
    }

    @Test
    void rotation_negative1_conversionBoundary() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("BoundaryNeg", s -> s, c -> c.rotation(-1))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            // -1 → 90 + 1 = 91
            assertEquals(91, style.getRotation());
        }
    }

    @Test
    void rotation_outOfRange_throws() {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () ->
                writer.column("Col", s -> s, c -> c.rotation(91)),
                "Rotation 91 should throw IllegalArgumentException");
        assertThrows(IllegalArgumentException.class, () ->
                writer.column("Col2", s -> s, c -> c.rotation(-91)),
                "Rotation -91 should throw IllegalArgumentException");
    }

    @Test
    void rotation_viaAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Rotated", s -> s, c -> c.rotation(60))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(60, style.getRotation());
        }
    }

    @Test
    void rotation_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Rotated", s -> s, c -> c.rotation(30))
                    .write(Stream.of("test"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(30, style.getRotation());
        }
    }

    @Test
    void rotation_columnConfigOutOfRange_throws() {
        ColumnConfig<String> config = new ColumnConfig<>();
        assertThrows(IllegalArgumentException.class, () -> config.rotation(91));
        assertThrows(IllegalArgumentException.class, () -> config.rotation(-91));
    }

    @Test
    void rotation_withOtherStyling() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("StyledRotated", s -> s, c -> c
                    .rotation(45)
                    .bold(true)
                    .backgroundColor(ExcelColor.LIGHT_YELLOW))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(45, style.getRotation(), "Rotation should be preserved with other styles");
            Font font = wb.getFontAt(style.getFontIndex());
            assertTrue(font.getBold(), "Bold should also be applied");
        }
    }

    @Test
    void rotation_multipleColumns_differentAngles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col1", s -> s, c -> c.rotation(0))
                .column("Col2", s -> s, c -> c.rotation(45))
                .column("Col3", s -> s, c -> c.rotation(-45))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertEquals(0, row.getCell(0).getCellStyle().getRotation());
            assertEquals(45, row.getCell(1).getCellStyle().getRotation());
            assertEquals(135, row.getCell(2).getCellStyle().getRotation());
        }
    }
}
