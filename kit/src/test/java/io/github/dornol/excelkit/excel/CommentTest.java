package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CommentTest {
    @Test
    void comment_multipleColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s, c -> c.comment(s -> "Comment for A: " + s))
                .column("B", s -> s.toUpperCase(), c -> c.comment(s -> "Comment for B: " + s))
                .write(Stream.of("hello"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertNotNull(row.getCell(0).getCellComment());
            assertNotNull(row.getCell(1).getCellComment());
            assertEquals("Comment for A: hello",
                    row.getCell(0).getCellComment().getString().getString());
            assertEquals("Comment for B: hello",
                    row.getCell(1).getCellComment().getString().getString());
        }
    }

    @Test
    void comment_emptyString() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s, c -> c.comment(s -> ""))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // Empty string comment should still be added
            assertNotNull(wb.getSheetAt(0).getRow(1).getCell(0).getCellComment());
        }
    }

    @Test
    void comment_multipleRows() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.comment(s -> "Hi " + s))
                .write(Stream.of("Alice", "Bob", "Charlie"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            for (int i = 1; i <= 3; i++) {
                assertNotNull(wb.getSheetAt(0).getRow(i).getCell(0).getCellComment());
            }
        }
    }
}
