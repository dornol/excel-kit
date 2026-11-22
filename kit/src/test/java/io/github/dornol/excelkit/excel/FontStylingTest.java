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

class FontStylingTest {

    @Test
    void fontColor_rgb_appliesCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Colored", s -> s, c -> c.fontColor(255, 0, 0))
                .write(Stream.of("red text"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
            XSSFColor color = xssfFont.getXSSFColor();
            assertNotNull(color, "Font color should not be null");
            byte[] rgb = color.getRGB();
            assertNotNull(rgb);
            assertEquals((byte) 255, rgb[0], "Red component");
            assertEquals((byte) 0, rgb[1], "Green component");
            assertEquals((byte) 0, rgb[2], "Blue component");
        }
    }

    @Test
    void fontColor_preset_appliesCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Blue", s -> s, c -> c.fontColor(ExcelColor.BLUE))
                .write(Stream.of("blue text"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
            XSSFColor color = xssfFont.getXSSFColor();
            assertNotNull(color);
            byte[] rgb = color.getRGB();
            assertNotNull(rgb);
            assertEquals((byte) 0, rgb[0]);
            assertEquals((byte) 0, rgb[1]);
            assertEquals((byte) 255, rgb[2]);
        }
    }

    @Test
    void fontColor_green_rgb() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Green", s -> s, c -> c.fontColor(0, 128, 0))
                .write(Stream.of("green"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
            byte[] rgb = xssfFont.getXSSFColor().getRGB();
            assertEquals((byte) 0, rgb[0]);
            assertEquals((byte) 128, rgb[1]);
            assertEquals((byte) 0, rgb[2]);
        }
    }

    @Test
    void strikethrough_enabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Strike", s -> s, c -> c.strikethrough())
                .write(Stream.of("struck out"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertTrue(font.getStrikeout(), "Font should be strikethrough");
        }
    }

    @Test
    void strikethrough_explicitTrue() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Strike", s -> s, c -> c.strikethrough(true))
                .write(Stream.of("struck"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertTrue(font.getStrikeout());
        }
    }

    @Test
    void strikethrough_disabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("NoStrike", s -> s, c -> c.strikethrough(false))
                .write(Stream.of("normal text"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertFalse(font.getStrikeout(), "Font should not be strikethrough");
        }
    }

    @Test
    void underline_enabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Underlined", s -> s, c -> c.underline())
                .write(Stream.of("underlined text"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertEquals(Font.U_SINGLE, font.getUnderline(), "Font should be underlined");
        }
    }

    @Test
    void underline_explicitTrue() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("UL", s -> s, c -> c.underline(true))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertEquals(Font.U_SINGLE, font.getUnderline());
        }
    }

    @Test
    void underline_disabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("NoUL", s -> s, c -> c.underline(false))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertEquals(Font.U_NONE, font.getUnderline());
        }
    }

    @Test
    void combined_allFontStyling() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Styled", s -> s, c -> c
                    .fontColor(255, 0, 0)
                    .bold(true)
                    .fontSize(14)
                    .underline()
                    .strikethrough())
                .write(Stream.of("fully styled"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertTrue(font.getBold(), "Font should be bold");
            assertTrue(font.getStrikeout(), "Font should be strikethrough");
            assertEquals(Font.U_SINGLE, font.getUnderline(), "Font should be underlined");
            assertEquals(14, font.getFontHeightInPoints(), "Font size should be 14");
            XSSFFont xssfFont = (XSSFFont) font;
            assertNotNull(xssfFont.getXSSFColor(), "Font color should be set");
        }
    }

    @Test
    void fontStyling_viaAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Styled", s -> s, c -> c
                        .fontColor(ExcelColor.RED)
                        .strikethrough()
                        .underline())
                .write(Stream.of("styled"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertTrue(font.getStrikeout());
            assertEquals(Font.U_SINGLE, font.getUnderline());
        }
    }

    @Test
    void fontStyling_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Colored", s -> s, c -> c
                            .fontColor(0, 128, 0)
                            .underline()
                            .strikethrough())
                    .write(Stream.of("styled"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Font font = xwb.getFontAt(xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
            assertTrue(font.getStrikeout());
            assertEquals(Font.U_SINGLE, font.getUnderline());
            XSSFFont xf = (XSSFFont) font;
            assertNotNull(xf.getXSSFColor());
        }
    }

    @Test
    void fontColor_withBackgroundColor_bothApply() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Both", s -> s, c -> c
                    .fontColor(ExcelColor.RED)
                    .backgroundColor(ExcelColor.LIGHT_YELLOW))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
            assertNotNull(xssfFont.getXSSFColor(), "Font color should be set");
            assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        }
    }

    @Test
    void multipleColumns_differentFontStyles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Red", s -> s, c -> c.fontColor(255, 0, 0))
                .column("Strike", s -> s, c -> c.strikethrough())
                .column("Under", s -> s, c -> c.underline())
                .column("Plain", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            // Red column has color
            XSSFFont f0 = (XSSFFont) wb.getFontAt(row.getCell(0).getCellStyle().getFontIndex());
            assertNotNull(f0.getXSSFColor());
            // Strike column
            Font f1 = wb.getFontAt(row.getCell(1).getCellStyle().getFontIndex());
            assertTrue(f1.getStrikeout());
            // Underline column
            Font f2 = wb.getFontAt(row.getCell(2).getCellStyle().getFontIndex());
            assertEquals(Font.U_SINGLE, f2.getUnderline());
        }
    }
}
