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

class TabColorTest {

    @Test
    void tabColor_rgb_appliesCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .tabColor(255, 0, 0)
                .column("Data", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
            assertNotNull(tabColor, "Tab color should be set");
            byte[] rgb = tabColor.getRGB();
            assertNotNull(rgb);
            assertEquals((byte) 255, rgb[0], "Red component");
            assertEquals((byte) 0, rgb[1], "Green component");
            assertEquals((byte) 0, rgb[2], "Blue component");
        }
    }

    @Test
    void tabColor_preset_appliesCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .tabColor(ExcelColor.BLUE)
                .column("Data", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
            assertNotNull(tabColor, "Tab color should be set");
            byte[] rgb = tabColor.getRGB();
            assertNotNull(rgb);
            assertEquals((byte) 0, rgb[0]);
            assertEquals((byte) 0, rgb[1]);
            assertEquals((byte) 255, rgb[2]);
        }
    }

    @Test
    void tabColor_noColor_tabColorIsNull() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Data", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
            assertNull(tabColor, "Tab color should be null when not set");
        }
    }

    @Test
    void tabColor_multipleSheets_applied() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().maxRows(2)
                .tabColor(0, 255, 0)
                .column("Data", s -> s)
                .write(Stream.of("a", "b", "c"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // Tab color should be applied to all rollover sheets
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                XSSFColor tabColor = wb.getSheetAt(i).getTabColor();
                assertNotNull(tabColor, "Tab color should be set on sheet " + i);
            }
        }
    }

    @Test
    void tabColor_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("TabTest")
                    .tabColor(ExcelColor.RED)
                    .column("Data", s -> s)
                    .write(Stream.of("test"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFColor tabColor = xwb.getSheetAt(0).getTabColor();
            assertNotNull(tabColor, "Tab color should be set via SheetWriter");
            byte[] rgb = tabColor.getRGB();
            assertEquals((byte) 255, rgb[0]);
            assertEquals((byte) 0, rgb[1]);
            assertEquals((byte) 0, rgb[2]);
        }
    }

    @Test
    void tabColor_differentPerSheet() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Red")
                    .tabColor(ExcelColor.RED)
                    .column("Data", s -> s)
                    .write(Stream.of("a"));
            wb.<String>sheet("Blue")
                    .tabColor(ExcelColor.BLUE)
                    .column("Data", s -> s)
                    .write(Stream.of("b"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            byte[] red = xwb.getSheetAt(0).getTabColor().getRGB();
            byte[] blue = xwb.getSheetAt(1).getTabColor().getRGB();
            assertEquals((byte) 255, red[0]);
            assertEquals((byte) 0, blue[0]);
            assertEquals((byte) 255, blue[2]);
        }
    }
}
