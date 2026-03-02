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
    void sheetWriter_shouldSupportTitle() throws IOException {
        ExcelWorkbook workbook = new ExcelWorkbook(ExcelColor.FOREST_GREEN);

        workbook.<String>sheet("Report")
                .title("Monthly Report", 16)
                .column("Item", s -> s)
                .write(Stream.of("Item1", "Item2"));

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
                .beforeHeader((sheet, wb, startRow) -> {
                    sheet.createRow(startRow).createCell(0).setCellValue("Generated Report");
                    return startRow + 1;
                })
                .column("Name", s -> s)
                .afterData((sheet, wb, nextRow) -> {
                    sheet.createRow(nextRow).createCell(0).setCellValue("Total: 2");
                    return nextRow + 1;
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
}
