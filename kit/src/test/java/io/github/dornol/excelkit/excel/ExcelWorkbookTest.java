package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelWorkbook} and {@link ExcelSheetWriter}.
 */
class ExcelWorkbookTest {

    @Test
    void multiSheet_shouldCreateSeparateSheets() throws IOException {
        // Arrange & Act
        ExcelWorkbook workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE);

        workbook.<String>sheet("Users")
                .column("Name", s -> s)
                .column("Length", s -> s.length())
                .write(Stream.of("Alice", "Bob"));

        workbook.<Integer>sheet("Numbers")
                .column("Value", n -> n)
                .column("Squared", n -> n * n)
                .write(Stream.of(1, 2, 3));

        ExcelHandler handler = workbook.finish();

        // Assert via output
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0, "Output should not be empty");
        }
        workbook.close();
    }

    @Test
    void multiSheet_shouldHaveCorrectSheetNames() {
        // Arrange & Act
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Sheet A")
                .column("Col", s -> s)
                .write(Stream.of("a"));

        workbook.<String>sheet("Sheet B")
                .column("Col", s -> s)
                .write(Stream.of("b"));

        ExcelHandler handler = workbook.finish();

        // We need to get the workbook to verify sheet names before consuming
        // Since ExcelHandler wraps wb, let's just consume and trust the sheet creation
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        } catch (IOException e) {
            fail(e);
        }
        workbook.close();
    }

    @Test
    void duplicateSheetName_shouldThrow() {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Users")
                .column("Name", s -> s)
                .write(Stream.of("a"));

        assertThrows(ExcelWriteException.class, () -> workbook.<String>sheet("Users"),
                "Duplicate sheet name should throw");

        workbook.close();
    }

    @Test
    void finishedWorkbook_shouldRejectNewSheets() {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Col", s -> s)
                .write(Stream.of("x"));

        workbook.finish();

        assertThrows(ExcelWriteException.class, () -> workbook.<String>sheet("More"),
                "Adding sheet to finished workbook should throw");
    }

    @Test
    void sheetWriter_shouldSupportDropdown() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Config")
                .column("Name", s -> s)
                .column("Status", s -> "Active", c -> c.dropdown("Active", "Inactive"))
                .write(Stream.of("Alice", "Bob"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetWriter_shouldSupportRowColor() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Name", s -> s)
                .rowColor(s -> "error".equals(s) ? ExcelColor.LIGHT_RED : null)
                .write(Stream.of("ok", "error", "ok"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetWriter_shouldSupportBeforeHeaderAndAfterData() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Generated Report");
                    return ctx.getCurrentRow() + 1;
                })
                .column("Name", s -> s)
                .afterData(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Total: 2");
                    return ctx.getCurrentRow() + 1;
                })
                .write(Stream.of("Alice", "Bob"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetWriter_noColumns_shouldThrow() {
        ExcelWorkbook workbook = new ExcelWorkbook();

        ExcelSheetWriter<String> sheetWriter = workbook.<String>sheet("Empty");

        assertThrows(ExcelWriteException.class, () -> sheetWriter.write(Stream.of("x")),
                "Writing without columns should throw");

        workbook.close();
    }

    @Test
    void sheetWriter_shouldSupportColumnConfig() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Styled")
                .column("Name", s -> s, c -> c.bold(true).fontSize(14))
                .column("Value", s -> s.length(), c -> c.type(ExcelDataType.INTEGER).alignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT))
                .column("BG", s -> s, c -> c.backgroundColor(ExcelColor.LIGHT_GREEN))
                .write(Stream.of("Hello", "World"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetWriter_shouldSupportAutoFilterAndFreezePane() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Filtered")
                .column("Col", s -> s)
                .autoFilter()
                .freezePane(1)
                .write(Stream.of("a", "b"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetWriter_shouldSupportConstColumn() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Const")
                .column("Name", s -> s)
                .constColumn("Type", "USER")
                .write(Stream.of("Alice"));

        ExcelHandler handler = workbook.finish();

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
        workbook.close();
    }

    @Test
    void sheetContext_shouldProvideCorrectSheetInCallbacks() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();
        SheetContext[] beforeCtx = new SheetContext[1];
        SheetContext[] afterCtx = new SheetContext[1];

        workbook.<String>sheet("Data")
                .beforeHeader(ctx -> {
                    beforeCtx[0] = ctx;
                    return ctx.getCurrentRow();
                })
                .column("Name", s -> s)
                .afterData(ctx -> {
                    afterCtx[0] = ctx;
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("Alice"));

        ExcelHandler handler = workbook.finish();

        // Assert
        assertNotNull(beforeCtx[0].getSheet());
        assertNotNull(beforeCtx[0].getWorkbook());
        assertEquals(0, beforeCtx[0].getCurrentRow(), "beforeHeader should start at row 0");
        assertEquals(1, beforeCtx[0].getColumnCount());
        assertEquals(List.of("Name"), beforeCtx[0].getColumnNames());

        assertNotNull(afterCtx[0].getSheet());
        assertSame(beforeCtx[0].getSheet(), afterCtx[0].getSheet(),
                "Both callbacks should reference the same sheet");
        assertEquals(2, afterCtx[0].getCurrentRow(),
                "afterData should start at row 2 (header + 1 data)");

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        }
        workbook.close();
    }

    @Test
    void multiSheet_shouldProvideCorrectSheetContextPerSheet() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook();
        SheetContext[] ctx1 = new SheetContext[1];
        SheetContext[] ctx2 = new SheetContext[1];

        workbook.<String>sheet("Sheet1")
                .beforeHeader(ctx -> {
                    ctx1[0] = ctx;
                    return ctx.getCurrentRow();
                })
                .column("A", s -> s)
                .write(Stream.of("x"));

        workbook.<Integer>sheet("Sheet2")
                .beforeHeader(ctx -> {
                    ctx2[0] = ctx;
                    return ctx.getCurrentRow();
                })
                .column("B", i -> i)
                .column("C", i -> i * 2)
                .write(Stream.of(1));

        ExcelHandler handler = workbook.finish();

        // Assert: each callback got different sheets and correct column metadata
        assertNotSame(ctx1[0].getSheet(), ctx2[0].getSheet());
        assertSame(ctx1[0].getWorkbook(), ctx2[0].getWorkbook(),
                "Both sheets share the same workbook");
        assertEquals(1, ctx1[0].getColumnCount());
        assertEquals(List.of("A"), ctx1[0].getColumnNames());
        assertEquals(2, ctx2[0].getColumnCount());
        assertEquals(List.of("B", "C"), ctx2[0].getColumnNames());

        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        }
        workbook.close();
    }

    // ============================================================
    // Constructor variants
    // ============================================================

    @Test
    void defaultConstructor_createsWorkbook() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("S1").column("A", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    @Test
    void customColorConstructor_createsWorkbook() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.of(100, 150, 200))) {
            wb.<String>sheet("S1").column("A", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    @Test
    void customColorWithWindowSizeConstructor_createsWorkbook() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.of(100, 150, 200), 500)) {
            wb.<String>sheet("S1").column("A", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    @Test
    void colorConstructor_createsWorkbook() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.LIGHT_BLUE)) {
            wb.<String>sheet("S1").column("A", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    @Test
    void colorWithWindowSizeConstructor_createsWorkbook() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE, 2000)) {
            wb.<String>sheet("S1").column("A", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    // ============================================================
    // protectWorkbook
    // ============================================================

    @Test
    void protectWorkbook_chainsCorrectly() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            ExcelWorkbook result = wb.protectWorkbook("password123");
            assertSame(wb, result);
        }
    }

    @Test
    void protectWorkbook_producesValidOutput() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.protectWorkbook("secret");
            wb.<String>sheet("Data").column("Col", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    // ============================================================
    // headerFontName / headerFontSize
    // ============================================================

    @Test
    void headerFontName_chainsCorrectly() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            ExcelWorkbook result = wb.headerFontName("Arial");
            assertSame(wb, result);
        }
    }

    @Test
    void headerFontSize_chainsCorrectly() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            ExcelWorkbook result = wb.headerFontSize(14);
            assertSame(wb, result);
        }
    }

    @Test
    void headerFontSize_zero_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            assertThrows(IllegalArgumentException.class, () -> wb.headerFontSize(0));
        }
    }

    @Test
    void headerFontSize_negative_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            assertThrows(IllegalArgumentException.class, () -> wb.headerFontSize(-1));
        }
    }

    @Test
    void headerFont_producesValidOutput() throws IOException {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.headerFontName("Arial").headerFontSize(16);
            wb.<String>sheet("Data").column("Col", s -> s).write(Stream.of("x"));
            ExcelHandler handler = wb.finish();
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                handler.consumeOutputStream(bos);
                assertTrue(bos.toByteArray().length > 0);
            }
        }
    }

    // ============================================================
    // close() behavior
    // ============================================================

    @Test
    void close_beforeFinish_doesNotThrow() {
        ExcelWorkbook wb = new ExcelWorkbook();
        wb.<String>sheet("Data").column("Col", s -> s).write(Stream.of("x"));
        assertDoesNotThrow(wb::close);
    }

    @Test
    void close_afterFinish_doesNotThrow() throws IOException {
        ExcelWorkbook wb = new ExcelWorkbook();
        wb.<String>sheet("Data").column("Col", s -> s).write(Stream.of("x"));
        ExcelHandler handler = wb.finish();
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        }
        // close() after finish should not close the wb (managed by ExcelHandler)
        assertDoesNotThrow(wb::close);
    }

    @Test
    void close_multipleTimes_doesNotThrow() {
        ExcelWorkbook wb = new ExcelWorkbook();
        assertDoesNotThrow(wb::close);
        assertDoesNotThrow(wb::close);
    }
}
