package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class BorderStyleTest {

    @Test
    void border_shouldApplyConfiguredBorderStyle() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.border(ExcelBorderStyle.MEDIUM))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.MEDIUM, style.getBorderTop());
            assertEquals(BorderStyle.MEDIUM, style.getBorderBottom());
            assertEquals(BorderStyle.MEDIUM, style.getBorderLeft());
            assertEquals(BorderStyle.MEDIUM, style.getBorderRight());
        }
    }

    @Test
    void border_none_shouldRemoveBorders() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.border(ExcelBorderStyle.NONE))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.NONE, style.getBorderTop());
        }
    }

    @Test
    void border_thick_shouldApply() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.border(ExcelBorderStyle.THICK))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.THICK, style.getBorderTop());
        }
    }

    @Test
    void border_dashed_shouldApply() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.border(ExcelBorderStyle.DASHED))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.DASHED, style.getBorderTop());
        }
    }

    @Test
    void defaultBorder_shouldBeThin() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s)
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.THIN, style.getBorderTop());
        }
    }

    @Test
    void border_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = new ExcelWorkbook()) {
            workbook.<String>sheet("Sheet1")
                    .column("Name", s -> s, c -> c.border(ExcelBorderStyle.DOUBLE))
                    .write(Stream.of("Alice"));
            workbook.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.DOUBLE, style.getBorderTop());
        }
    }

    @Test
    void excelBorderStyleEnum_shouldCoverAllValues() {
        for (ExcelBorderStyle style : ExcelBorderStyle.values()) {
            assertNotNull(style.toPoiBorderStyle());
        }
    }
}
