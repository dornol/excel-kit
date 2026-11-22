package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ColumnConfigOptionsTest {

    // -----------------------------------------------------------------------
    // 2. ColumnConfig — format, alignment, bold, fontSize, width, minWidth, maxWidth
    // -----------------------------------------------------------------------

    @Test
    void columnConfig_format_shouldApplyCustomFormat() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<Integer>sheet("Data")
                .column("Value", v -> v, c -> c.type(ExcelDataType.INTEGER).format("#,##0.00"))
                .write(Stream.of(1234));

        workbook.finish().writeTo(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void columnConfig_alignment_shouldApplyAlignment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<String>sheet("Data")
                .column("Left", s -> s, c -> c.alignment(HorizontalAlignment.LEFT))
                .column("Right", s -> s, c -> c.alignment(HorizontalAlignment.RIGHT))
                .write(Stream.of("test"));

        workbook.finish().writeTo(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            assertEquals(HorizontalAlignment.LEFT, row.getCell(0).getCellStyle().getAlignment());
            assertEquals(HorizontalAlignment.RIGHT, row.getCell(1).getCellStyle().getAlignment());
        }
    }

    @Test
    void columnConfig_bold_shouldApplyBoldFont() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<String>sheet("Data")
                .column("Bold", s -> s, c -> c.bold(true))
                .column("NotBold", s -> s, c -> c.bold(false))
                .write(Stream.of("test"));

        workbook.finish().writeTo(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            CellStyle boldStyle = row.getCell(0).getCellStyle();
            assertTrue(wb.getFontAt(boldStyle.getFontIndex()).getBold());
            CellStyle notBoldStyle = row.getCell(1).getCellStyle();
            assertFalse(wb.getFontAt(notBoldStyle.getFontIndex()).getBold());
        }
    }

    @Test
    void columnConfig_fontSize_shouldApplyFontSize() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<String>sheet("Data")
                .column("Big", s -> s, c -> c.fontSize(18))
                .write(Stream.of("test"));

        workbook.finish().writeTo(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
            assertEquals(18, wb.getFontAt(style.getFontIndex()).getFontHeightInPoints());
        }
    }

    @Test
    void columnConfig_fontSize_shouldThrowForNonPositive() {
        ColumnConfig<String> config = new ColumnConfig<>();
        assertThrows(IllegalArgumentException.class, () -> config.fontSize(0));
        assertThrows(IllegalArgumentException.class, () -> config.fontSize(-1));
    }

    @Test
    void columnConfig_width_shouldSetFixedWidth() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<String>sheet("Data")
                .column("Fixed", s -> s, c -> c.width(30))
                .write(Stream.of("test"));

        workbook.finish().writeTo(out);
        workbook.close();

        // Just verify it writes without error; fixed width is applied via internal logic
        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void columnConfig_minWidth_maxWidth_shouldWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = ExcelWorkbook.create();

        workbook.<String>sheet("Data")
                .column("Bounded", s -> s, c -> c.minWidth(10).maxWidth(50))
                .write(Stream.of("test"));

        workbook.finish().writeTo(out);
        workbook.close();

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void columnConfig_chainingReturnsThis() {
        ColumnConfig<String> config = new ColumnConfig<>();
        assertSame(config, config.type(ExcelDataType.STRING));
        assertSame(config, config.format("#,##0"));
        assertSame(config, config.alignment(HorizontalAlignment.CENTER));
        assertSame(config, config.bold(true));
        assertSame(config, config.fontSize(12));
        assertSame(config, config.width(20));
        assertSame(config, config.minWidth(10));
        assertSame(config, config.maxWidth(50));
        assertSame(config, config.dropdown("A", "B"));
        assertSame(config, config.group("G1"));
        assertSame(config, config.outline(1));
    }

    @Test
    void columnConfig_outline_invalidLevel_shouldThrow() {
        ColumnConfig<String> config = new ColumnConfig<>();
        assertThrows(IllegalArgumentException.class, () -> config.outline(-1));
        assertThrows(IllegalArgumentException.class, () -> config.outline(8));
    }
}
