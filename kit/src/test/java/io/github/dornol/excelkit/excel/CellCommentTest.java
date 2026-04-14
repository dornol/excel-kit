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
                .column("Name", s -> s, c -> c.headerComment((String) null))
                .write(Stream.of("Alice"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNull(wb.getSheetAt(0).getRow(0).getCell(0).getCellComment());
        }
    }

    @Test
    void excelCellComment_record() {
        ExcelCellComment c1 = ExcelCellComment.of("text");
        assertEquals("text", c1.text());
        assertNull(c1.author());
        assertEquals(0, c1.width());
        assertEquals(0, c1.height());

        ExcelCellComment c2 = ExcelCellComment.of("text").author("bob").size(4, 6);
        assertEquals("text", c2.text());
        assertEquals("bob", c2.author());
        assertEquals(4, c2.width());
        assertEquals(6, c2.height());

        // Record should remain immutable — withers return new instances
        assertNotSame(c1, c1.author("x"));
    }

    @Test
    void headerComment_excelCellCommentOverload_appliesSizeAndAuthor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c.headerComment(
                        ExcelCellComment.of("detailed note").author("System").size(4, 5)))
                .write(Stream.of("v"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
            var comment = headerCell.getCellComment();
            assertNotNull(comment);
            assertEquals("detailed note", comment.getString().getString());
            assertEquals("System", comment.getAuthor());
            // Size: col2 = col1 + width, row2 = row1 + height
            var anchor = comment.getClientAnchor();
            assertEquals(0, anchor.getCol1());
            assertEquals(4, anchor.getCol2());
            assertEquals(0, anchor.getRow1());
            assertEquals(5, anchor.getRow2());
        }
    }

    @Test
    void commentSize_columnLevel_appliesToDataAndHeader() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c
                        .headerComment("hdr")
                        .comment(s -> "cell")
                        .commentSize(3, 6))
                .write(Stream.of("v"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            var headerAnchor = sheet.getRow(0).getCell(0).getCellComment().getClientAnchor();
            assertEquals(3, headerAnchor.getCol2() - headerAnchor.getCol1());
            assertEquals(6, headerAnchor.getRow2() - headerAnchor.getRow1());

            var dataAnchor = sheet.getRow(1).getCell(0).getCellComment().getClientAnchor();
            assertEquals(3, dataAnchor.getCol2() - dataAnchor.getCol1());
            assertEquals(6, dataAnchor.getRow2() - dataAnchor.getRow1());
        }
    }

    @Test
    void excelCellComment_sizeWins_overColumnCommentSize() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c
                        .headerComment(ExcelCellComment.of("x").size(7, 8))
                        .commentSize(3, 3))
                .write(Stream.of("v"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var anchor = wb.getSheetAt(0).getRow(0).getCell(0).getCellComment().getClientAnchor();
            assertEquals(7, anchor.getCol2() - anchor.getCol1());
            assertEquals(8, anchor.getRow2() - anchor.getRow1());
        }
    }

    @Test
    void commentSize_negativeOrZero_throws() {
        assertThrows(IllegalArgumentException.class, () ->
                ExcelWriter.<String>create().column("C", s -> s, c -> c.commentSize(0, 3)));
        assertThrows(IllegalArgumentException.class, () ->
                ExcelWriter.<String>create().column("C", s -> s, c -> c.commentSize(3, -1)));
    }

    @Test
    void excelCellComment_nullText_throws() {
        assertThrows(IllegalArgumentException.class, () -> new ExcelCellComment(null, null, 0, 0));
    }

    @Test
    void excelCellComment_negativeSize_throws() {
        assertThrows(IllegalArgumentException.class, () -> ExcelCellComment.of("x").size(-1, 0));
    }
}
