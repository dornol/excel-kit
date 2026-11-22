package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CellColorTest {

    // ========================================================================
    // Feature 5: Conditional cell color
    // ========================================================================
    @Test
    void cellColor_shouldApplyPerCellColor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Integer>create()
                .column("Value", i -> i, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int v = ((Number) value).intValue();
                        if (v < 0) return ExcelColor.LIGHT_RED;
                        if (v > 100) return ExcelColor.LIGHT_GREEN;
                        return null;
                    }))
                .write(Stream.of(-5, 50, 200))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 1 (value=-5) should have red background
            XSSFColor color1 = (XSSFColor) sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color1);
            assertEquals(ExcelColor.LIGHT_RED.getR(), Byte.toUnsignedInt(color1.getRGB()[0]));

            // Row 2 (value=50) should have no fill override (default style)
            // Row 3 (value=200) should have green background
            XSSFColor color3 = (XSSFColor) sheet.getRow(3).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color3);
            assertEquals(ExcelColor.LIGHT_GREEN.getR(), Byte.toUnsignedInt(color3.getRGB()[0]));
        }
    }

    @Test
    void cellColor_shouldOverrideRowColor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Integer>create()
                .rowColor(i -> ExcelColor.LIGHT_YELLOW) // all rows yellow
                .column("Value", i -> i, cfg -> cfg
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int v = ((Number) value).intValue();
                        return v < 0 ? ExcelColor.LIGHT_RED : null; // negative → red, else fall through to row color
                    }))
                .write(Stream.of(-5, 50))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 1 (value=-5): cellColor=RED takes precedence over rowColor=YELLOW
            XSSFColor color1 = (XSSFColor) sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertEquals(ExcelColor.LIGHT_RED.getR(), Byte.toUnsignedInt(color1.getRGB()[0]));

            // Row 2 (value=50): cellColor=null, falls back to rowColor=YELLOW
            XSSFColor color2 = (XSSFColor) sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertEquals(ExcelColor.LIGHT_YELLOW.getR(), Byte.toUnsignedInt(color2.getRGB()[0]));
        }
    }

    @Test
    void cellColor_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Test")
                    .column("Value", i -> i, c -> c
                            .type(ExcelDataType.INTEGER)
                            .cellColor((value, row) -> {
                                int v = ((Number) value).intValue();
                                return v > 100 ? ExcelColor.LIGHT_GREEN : null;
                            }))
                    .write(Stream.of(50, 200));
            wb.finish().writeTo(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 2 (value=200) should have green background
            XSSFColor color = (XSSFColor) sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color);
            assertEquals(ExcelColor.LIGHT_GREEN.getR(), Byte.toUnsignedInt(color.getRGB()[0]));
        }
    }
}
