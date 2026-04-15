package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class HeaderCustomizationTest {

    @Nested
    class HeaderBackgroundColor {
        @Test
        void perColumnOverride_appliesToHeaderCell() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create().headerColor(ExcelColor.STEEL_BLUE)
                    .column("Alert", s -> s, c -> c.headerBackgroundColor(ExcelColor.LIGHT_RED))
                    .column("Normal", s -> s)
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                var alertStyle = sheet.getRow(0).getCell(0).getCellStyle();
                var normalStyle = sheet.getRow(0).getCell(1).getCellStyle();
                // Different styles = override took effect
                assertNotEquals(alertStyle.getIndex(), normalStyle.getIndex());

                // Verify RGB explicitly for the overridden cell
                var fill = alertStyle.getFillForegroundColorColor();
                assertNotNull(fill);
                byte[] expected = {
                        (byte) ExcelColor.LIGHT_RED.getR(),
                        (byte) ExcelColor.LIGHT_RED.getG(),
                        (byte) ExcelColor.LIGHT_RED.getB()};
                assertArrayEquals(expected, ((org.apache.poi.xssf.usermodel.XSSFColor) fill).getRGB());
            }
        }

        @Test
        void rgbOverload() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Col", s -> s, c -> c.headerBackgroundColor(120, 40, 200))
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var style = wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
                var rgb = ((org.apache.poi.xssf.usermodel.XSSFColor) style.getFillForegroundColorColor()).getRGB();
                assertEquals(120, rgb[0] & 0xFF);
                assertEquals(40,  rgb[1] & 0xFF);
                assertEquals(200, rgb[2] & 0xFF);
            }
        }
    }

    @Nested
    class HeaderRowHeight {
        @Test
        void singleHeaderRow_appliesHeight() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .headerRowHeight(40f)
                    .column("Col", s -> s)
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(40f, wb.getSheetAt(0).getRow(0).getHeightInPoints(), 0.01);
            }
        }

        @Test
        void multiLevelGroupHeaders_allHeaderRowsGetHeight() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .headerRowHeight(25f)
                    .column("Q1", s -> s, c -> c.group("Financial", "Revenue"))
                    .column("Profit", s -> s, c -> c.group("Financial"))
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals(25f, sheet.getRow(0).getHeightInPoints(), 0.01);  // group row
                assertEquals(25f, sheet.getRow(1).getHeightInPoints(), 0.01);  // subgroup row
                assertEquals(25f, sheet.getRow(2).getHeightInPoints(), 0.01);  // column header row
            }
        }

        @Test
        void zeroMeansDefault() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Col", s -> s)
                    .write(Stream.of("x"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // Default Excel row height is ~15pt. We just verify not overridden to a custom value.
                float h = wb.getSheetAt(0).getRow(0).getHeightInPoints();
                assertTrue(h > 0 && h < 25, "Expected default height, got " + h);
            }
        }

        @Test
        void negativeHeight_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> ExcelWriter.<String>create().headerRowHeight(-1f));
        }
    }

    @Nested
    class RowNumberColumn {
        @Test
        void addsSequentialOneBasedNumbers() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .rowNumberColumn("No.")
                    .column("Name", s -> s)
                    .write(Stream.of("a", "b", "c"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("No.", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals(1, (long) sheet.getRow(1).getCell(0).getNumericCellValue());
                assertEquals(2, (long) sheet.getRow(2).getCell(0).getNumericCellValue());
                assertEquals(3, (long) sheet.getRow(3).getCell(0).getNumericCellValue());
            }
        }
    }
}
