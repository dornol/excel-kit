package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for rowStyle, image size, and image URL features.
 */
class RowStyleImageTest {

    record Product(String name, int price) {}

    // ========================================================================
    // rowStyle — conditional row styling
    // ========================================================================
    @Nested
    class RowStyleTests {

        @Test
        void rowStyle_boldOnCondition() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .column("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                    .rowStyle(p -> p.price() > 500, style -> style.bold(true))
                    .write(Stream.of(new Product("Cheap", 100), new Product("Expensive", 1000)))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                // Row 1: Cheap (100) — not bold
                Font cheapFont = wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertFalse(cheapFont.getBold());
                // Row 2: Expensive (1000) — bold
                Font expensiveFont = wb.getFontAt(sheet.getRow(2).getCell(0).getCellStyle().getFontIndex());
                assertTrue(expensiveFont.getBold());
            }
        }

        @Test
        void rowStyle_backgroundColorOnCondition() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .column("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                    .rowStyle(p -> p.price() > 500,
                            style -> style.backgroundColor(ExcelColor.LIGHT_YELLOW))
                    .write(Stream.of(new Product("A", 100), new Product("B", 1000)))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                // Row 1: no background
                CellStyle cheapStyle = sheet.getRow(1).getCell(0).getCellStyle();
                assertNotEquals(FillPatternType.SOLID_FOREGROUND, cheapStyle.getFillPattern());
                // Row 2: yellow background
                CellStyle expStyle = sheet.getRow(2).getCell(0).getCellStyle();
                assertEquals(FillPatternType.SOLID_FOREGROUND, expStyle.getFillPattern());
            }
        }

        @Test
        void rowStyle_firstMatchWins() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .column("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                    .rowStyle(p -> p.price() > 900, style -> style.bold(true).italic(true))
                    .rowStyle(p -> p.price() > 500, style -> style.bold(true))
                    .write(Stream.of(new Product("Expensive", 1000)))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                Font font = wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex());
                // First rule matches: bold + italic
                assertTrue(font.getBold());
                assertTrue(font.getItalic());
            }
        }

        @Test
        void rowStyle_fontColorAndSize() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .rowStyle(p -> true,
                            style -> style.fontSize(14).fontColor(ExcelColor.RED))
                    .write(Stream.of(new Product("Test", 100)))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                Font font = wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(14, font.getFontHeightInPoints());
                if (font instanceof XSSFFont xf) {
                    XSSFColor color = xf.getXSSFColor();
                    assertNotNull(color);
                    byte[] rgb = color.getRGB();
                    assertEquals((byte) ExcelColor.RED.getR(), rgb[0]);
                }
            }
        }

        @Test
        void rowStyle_noMatchLeavesDefaultStyle() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .rowStyle(p -> p.price() > 9999, style -> style.bold(true))
                    .write(Stream.of(new Product("Normal", 100)))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                Font font = wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertFalse(font.getBold());
            }
        }
    }

    // ========================================================================
    // ExcelImage — size specification
    // ========================================================================
    @Nested
    class ImageSizeTests {

        @Test
        void image_withSize_shouldPreserveWidthHeight() {
            byte[] data = new byte[]{(byte) 0x89, 0x50, 0x4E, 0x47}; // PNG header stub
            ExcelImage img = ExcelImage.png(data).size(3, 4);
            assertEquals(3, img.width());
            assertEquals(4, img.height());
        }

        @Test
        void image_defaultSize_shouldBeOneByOne() {
            byte[] data = new byte[]{(byte) 0x89, 0x50};
            ExcelImage img = ExcelImage.png(data);
            assertEquals(1, img.width());
            assertEquals(1, img.height());
        }

        @Test
        void image_withSize_shouldWriteSuccessfully() throws IOException {
            // 1x1 white PNG
            byte[] pngData = createMinimalPng();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<byte[]>create()
                    .column("Image", d -> ExcelImage.png(d).size(2, 3),
                            c -> c.type(ExcelDataType.IMAGE))
                    .write(Stream.of(pngData))
                    .writeTo(out);

            try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertNotNull(sheet.getRow(1));
                // Image was written (no exception), and drawing patriarch exists
                assertNotNull(sheet.getDrawingPatriarch());
            }
        }

        private byte[] createMinimalPng() {
            // Minimal 1x1 white PNG
            return new byte[]{
                    (byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
                    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
                    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
                    0x08, 0x02, 0x00, 0x00, 0x00, (byte) 0x90, 0x77, 0x53,
                    (byte) 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                    0x54, 0x08, (byte) 0xD7, 0x63, (byte) 0xF8, (byte) 0xCF, (byte) 0xC0, 0x00,
                    0x00, 0x00, 0x02, 0x00, 0x01, (byte) 0xE2, 0x21, (byte) 0xBC,
                    0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
                    0x44, (byte) 0xAE, 0x42, 0x60, (byte) 0x82 // IEND
            };
        }
    }
}
