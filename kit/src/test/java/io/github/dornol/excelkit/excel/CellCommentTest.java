package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CellCommentTest {

    @Test
    void comment_shouldAddCommentToCell() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.comment(s -> "Note: " + s))
                .write(Stream.of("Alice"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Alice", cell.getStringCellValue());
            assertNotNull(cell.getCellComment());
            assertEquals("Note: Alice", cell.getCellComment().getString().getString());
        }
    }

    @Test
    void comment_nullReturn_shouldNotAddComment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.comment(s -> null))
                .write(Stream.of("Alice"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertNull(cell.getCellComment());
        }
    }

    @Test
    void comment_conditionalComment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.comment(s -> s.startsWith("A") ? "Starts with A" : null))
                .write(Stream.of("Alice", "Bob"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertNotNull(sheet.getRow(1).getCell(0).getCellComment());
            assertEquals("Starts with A", sheet.getRow(1).getCell(0).getCellComment().getString().getString());
            assertNull(sheet.getRow(2).getCell(0).getCellComment());
        }
    }

    @Test
    void comment_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Sheet1")
                    .column("Name", s -> s, c -> c.comment(s -> "Hi " + s))
                    .write(Stream.of("Alice"));
            wb.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertNotNull(cell.getCellComment());
            assertEquals("Hi Alice", cell.getCellComment().getString().getString());
        }
    }

    @Test
    void headerComment_shouldAddCommentToHeaderCell() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("생년월일", s -> s, c -> c.headerComment("YYYY-MM-DD 형식으로 입력"))
                .column("이름", s -> s)
                .write(Stream.of("2024-01-01"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            var headerCell = sheet.getRow(0).getCell(0);
            assertEquals("생년월일", headerCell.getStringCellValue());
            assertNotNull(headerCell.getCellComment());
            assertEquals("YYYY-MM-DD 형식으로 입력",
                    headerCell.getCellComment().getString().getString());

            // second header has no comment
            assertNull(sheet.getRow(0).getCell(1).getCellComment());
            // data row unaffected
            assertNull(sheet.getRow(1).getCell(0).getCellComment());
        }
    }

    @Test
    void headerComment_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Sheet1")
                    .column("Amount", s -> s, c -> c.headerComment("원 단위"))
                    .write(Stream.of("1000"));
            wb.finish().writeTo(out);
        }
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            assertNotNull(headerCell.getCellComment());
            assertEquals("원 단위", headerCell.getCellComment().getString().getString());
        }
    }

    @Test
    void headerComment_withGroupHeader_attachesToColumnHeaderRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Price", s -> s, c -> c.group("Financial").headerComment("원 단위"))
                .column("Qty", s -> s, c -> c.group("Financial"))
                .write(Stream.of("100"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0 = group row, Row 1 = column header row
            var colHeaderCell = sheet.getRow(1).getCell(0);
            assertEquals("Price", colHeaderCell.getStringCellValue());
            assertNotNull(colHeaderCell.getCellComment());
            assertEquals("원 단위", colHeaderCell.getCellComment().getString().getString());

            // Group header row does not get the column comment
            assertNull(sheet.getRow(0).getCell(0).getCellComment());
        }
    }

    @Test
    void headerComment_null_shouldNotAddComment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.headerComment(null))
                .write(Stream.of("Alice"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNull(wb.getSheetAt(0).getRow(0).getCell(0).getCellComment());
        }
    }

    @Test
    void excelCellComment_record() {
        ExcelCellComment c1 = new ExcelCellComment("text");
        assertEquals("text", c1.text());
        assertNull(c1.author());

        ExcelCellComment c2 = new ExcelCellComment("text", "author");
        assertEquals("text", c2.text());
        assertEquals("author", c2.author());
    }
}
