package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class HeaderRowStylingTest {

    @Test
    void headerFontName_only_verifiesFontNameOnHeaderCell() throws Exception {
        var handler = ExcelWriter.<String>create()
                .headerFontName("Arial")
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertEquals("Arial", font.getFontName());
        }
    }

    @Test
    void headerFontSize_only_verifiesFontSizeOnHeaderCell() throws Exception {
        var handler = ExcelWriter.<String>create()
                .headerFontSize(18)
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertEquals(18, font.getFontHeightInPoints());
        }
    }

    @Test
    void headerFontNameAndSize_bothApplied() throws Exception {
        var handler = ExcelWriter.<String>create()
                .headerFontName("Courier New")
                .headerFontSize(14)
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertEquals("Courier New", font.getFontName());
            assertEquals(14, font.getFontHeightInPoints());
        }
    }

    @Test
    void excelWorkbook_headerFontNameAndSize_appliedToSheet() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.headerFontName("Times New Roman").headerFontSize(16);
            workbook.<String>sheet("Sheet1")
                    .column("Name", s -> s)
                    .write(Stream.of("Bob"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertEquals("Times New Roman", font.getFontName());
            assertEquals(16, font.getFontHeightInPoints());
        }
    }

    @Test
    void defaultHeaderStyle_isBold11pt() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
            assertTrue(font.getBold(), "Default header font should be bold");
            assertEquals(11, font.getFontHeightInPoints(),
                    "Default header font size should be 11pt");
        }
    }

    @Test
    void headerFontSize_zeroThrowsIllegalArgumentException() {
        var writer = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> writer.headerFontSize(0));
    }

    @Test
    void headerFontSize_negativeThrowsIllegalArgumentException() {
        var writer = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> writer.headerFontSize(-5));
    }

    @Test
    void multipleColumns_allHeadersHaveSameFont() throws Exception {
        var handler = ExcelWriter.<String>create()
                .headerFontName("Georgia")
                .headerFontSize(12)
                .column("Col1", s -> s)
                .column("Col2", s -> s)
                .column("Col3", s -> s)
                .write(Stream.of("data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var headerRow = wb.getSheetAt(0).getRow(0);
            for (int i = 0; i < 3; i++) {
                var font = wb.getFontAt(headerRow.getCell(i).getCellStyle().getFontIndex());
                assertEquals("Georgia", font.getFontName(),
                        "Header cell " + i + " should have font Georgia");
                assertEquals(12, font.getFontHeightInPoints(),
                        "Header cell " + i + " should have size 12");
            }
        }
    }
}
