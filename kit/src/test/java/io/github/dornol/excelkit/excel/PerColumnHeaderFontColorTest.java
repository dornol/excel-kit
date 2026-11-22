package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class PerColumnHeaderFontColorTest {

    @Test
    void headerFontColor_appliedToSpecificColumn_excelWriter() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Normal", s -> s)
                .column("Alert", s -> s, c -> c
                    .headerFontColor(ExcelColor.RED))
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerRow = wb.getSheetAt(0).getRow(0);

            // Normal column: default font color (black or white based on header bg)
            var normalFont = wb.getFontAt(headerRow.getCell(0).getCellStyle().getFontIndex());
            var alertFont = wb.getFontAt(headerRow.getCell(1).getCellStyle().getFontIndex());

            // Alert column should have RED font
            var alertXssfFont = (org.apache.poi.xssf.usermodel.XSSFFont) alertFont;
            var alertColor = alertXssfFont.getXSSFColor();
            assertNotNull(alertColor, "Alert header should have custom font color");
            byte[] rgb = alertColor.getRGB();
            assertEquals((byte) 255, rgb[0], "Red component");
            assertEquals((byte) 0, rgb[1], "Green component");
            assertEquals((byte) 0, rgb[2], "Blue component");

            // Normal column should NOT have RED font
            assertNotEquals(alertFont.getColor(), normalFont.getColor(),
                    "Normal and alert headers should have different font colors");
        }
    }

    @Test
    void headerFontColor_appliedToSpecificColumn_excelSheetWriter() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Sheet1")
                    .column("Normal", s -> s)
                    .column("Alert", s -> s, cfg -> cfg.headerFontColor(ExcelColor.RED))
                    .write(Stream.of("data"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerRow = wb.getSheetAt(0).getRow(0);

            var alertFont = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(headerRow.getCell(1).getCellStyle().getFontIndex());
            var alertColor = alertFont.getXSSFColor();
            assertNotNull(alertColor);
            byte[] rgb = alertColor.getRGB();
            assertEquals((byte) 255, rgb[0]);
            assertEquals((byte) 0, rgb[1]);
            assertEquals((byte) 0, rgb[2]);
        }
    }

    @Test
    void headerFontColor_null_usesDefaultHeaderStyle() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Col1", s -> s, c -> c
                    .headerFontColor((ExcelColor) null))
                .column("Col2", s -> s)
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerRow = wb.getSheetAt(0).getRow(0);
            var font1 = wb.getFontAt(headerRow.getCell(0).getCellStyle().getFontIndex());
            var font2 = wb.getFontAt(headerRow.getCell(1).getCellStyle().getFontIndex());
            assertEquals(font1.getColor(), font2.getColor(),
                    "Both headers should have the same default font color");
        }
    }

    @Test
    void headerFontColor_rgb_appliedCorrectly() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Custom", s -> s, c -> c
                    .headerFontColor(0, 128, 255))
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            var color = font.getXSSFColor();
            assertNotNull(color);
            byte[] rgb = color.getRGB();
            assertEquals((byte) 0, rgb[0]);
            assertEquals((byte) 128, rgb[1]);
            assertEquals((byte) 255, rgb[2]);
        }
    }

    @Test
    void headerFontColor_preservesBoldAndSize() throws Exception {
        var handler = ExcelWriter.<String>create()
                .headerFontName("Arial")
                .headerFontSize(14)
                .column("Alert", s -> s, c -> c
                    .headerFontColor(ExcelColor.RED))
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertTrue(font.getBold(), "Header font should still be bold");
            assertEquals("Arial", font.getFontName(), "Header font name should be preserved");
            assertEquals(14, font.getFontHeightInPoints(), "Header font size should be preserved");
        }
    }

    @Test
    void headerFontColor_multipleColumns_mixedStyles() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Normal", s -> s)
                .column("Red", s -> s, c -> c
                    .headerFontColor(ExcelColor.RED))
                .column("Blue", s -> s, c -> c
                    .headerFontColor(ExcelColor.BLUE))
                .column("Default", s -> s)
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerRow = wb.getSheetAt(0).getRow(0);

            // Red column
            var redFont = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(headerRow.getCell(1).getCellStyle().getFontIndex());
            byte[] redRgb = redFont.getXSSFColor().getRGB();
            assertEquals((byte) 255, redRgb[0]);
            assertEquals((byte) 0, redRgb[1]);
            assertEquals((byte) 0, redRgb[2]);

            // Blue column
            var blueFont = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(headerRow.getCell(2).getCellStyle().getFontIndex());
            byte[] blueRgb = blueFont.getXSSFColor().getRGB();
            assertEquals((byte) 0, blueRgb[0]);
            assertEquals((byte) 0, blueRgb[1]);
            assertEquals((byte) 255, blueRgb[2]);

            // Normal and Default columns should have same style
            var normalStyle = headerRow.getCell(0).getCellStyle();
            var defaultStyle = headerRow.getCell(3).getCellStyle();
            assertEquals(normalStyle.getFontIndex(), defaultStyle.getFontIndex(),
                    "Non-overridden columns should share the same header style");
        }
    }

    @Test
    void headerFontColor_withGroupHeaders() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Sheet1")
                    .column("Name", s -> s, cfg -> cfg.group("Info"))
                    .column("Alert", s -> s, cfg -> cfg
                            .group("Status")
                            .headerFontColor(ExcelColor.RED))
                    .write(Stream.of("data"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            // Row 0 = group header, Row 1 = column header
            var columnHeaderRow = wb.getSheetAt(0).getRow(1);
            var alertFont = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(columnHeaderRow.getCell(1).getCellStyle().getFontIndex());
            var alertColor = alertFont.getXSSFColor();
            assertNotNull(alertColor, "Alert column header should have custom font color in grouped layout");
            byte[] rgb = alertColor.getRGB();
            assertEquals((byte) 255, rgb[0]);
        }
    }

    @Test
    void headerFontColor_conditionalUsage() throws Exception {
        boolean hasError = true;

        var handler = ExcelWriter.<String>create()
                .column("Amount", s -> s, c -> c
                    .headerFontColor(hasError ? ExcelColor.RED : null))
                .write(Stream.of("1000"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = (org.apache.poi.xssf.usermodel.XSSFFont)
                    wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            var color = font.getXSSFColor();
            assertNotNull(color, "Conditional header font color should be applied when condition is true");
            assertEquals((byte) 255, color.getRGB()[0]);
        }
    }
}
