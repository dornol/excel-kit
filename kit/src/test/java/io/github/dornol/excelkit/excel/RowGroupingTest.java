package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class RowGroupingTest {

    @Test
    void groupRows_basic() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Data", s -> s)
                .afterData(ctx -> {
                    ctx.groupRows(1, 3);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("a", "b", "c", "d", "e"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be grouped");
            assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be grouped");
            assertTrue(sheet.getRow(3).getOutlineLevel() > 0, "Row 3 should be grouped");
            assertEquals(0, sheet.getRow(4).getOutlineLevel(), "Row 4 should not be grouped");
            assertEquals(0, sheet.getRow(5).getOutlineLevel(), "Row 5 should not be grouped");
        }
    }

    @Test
    void groupRows_collapsed() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Data", s -> s)
                .afterData(ctx -> {
                    ctx.groupRows(1, 3, true);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("a", "b", "c", "d"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be grouped");
            assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be grouped");
            assertTrue(sheet.getRow(3).getOutlineLevel() > 0, "Row 3 should be grouped");
        }
    }

    @Test
    void groupRows_multipleGroups() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Data", s -> s)
                .afterData(ctx -> {
                    ctx.groupRows(1, 2);
                    ctx.groupRows(4, 5);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("a", "b", "c", "d", "e", "f"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be in first group");
            assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be in first group");
            assertEquals(0, sheet.getRow(3).getOutlineLevel(), "Row 3 should not be grouped");
            assertTrue(sheet.getRow(4).getOutlineLevel() > 0, "Row 4 should be in second group");
            assertTrue(sheet.getRow(5).getOutlineLevel() > 0, "Row 5 should be in second group");
        }
    }

    @Test
    void groupRows_chainingReturnsContext() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Data", s -> s)
                .afterData(ctx -> {
                    // Verify chaining works
                    SheetContext result = ctx.groupRows(1, 2).groupRows(3, 4);
                    assertSame(ctx, result, "groupRows should return same context for chaining");
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("a", "b", "c", "d", "e"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNotNull(wb.getSheetAt(0));
        }
    }

    @Test
    void groupRows_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 3);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = xwb.getSheetAt(0);
            assertTrue(sheet.getRow(1).getOutlineLevel() > 0);
            assertTrue(sheet.getRow(3).getOutlineLevel() > 0);
            assertEquals(0, sheet.getRow(4).getOutlineLevel());
        }
    }
}
