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
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.comment(s -> "Note: " + s))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

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
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.comment(s -> null))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertNull(cell.getCellComment());
        }
    }

    @Test
    void comment_conditionalComment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c.comment(s -> s.startsWith("A") ? "Starts with A" : null))
                .write(Stream.of("Alice", "Bob"))
                .consumeOutputStream(out);

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
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Sheet1")
                    .column("Name", s -> s, c -> c.comment(s -> "Hi " + s))
                    .write(Stream.of("Alice"));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertNotNull(cell.getCellComment());
            assertEquals("Hi Alice", cell.getCellComment().getString().getString());
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
