package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelSheetWriter} — verifies actual Excel content.
 */
class ExcelSheetWriterEdgeCaseTest {

    record Item(String name, int value) {}

    // ============================================================
    // columnIf false condition — should produce fewer columns
    // ============================================================
    @Test
    void columnIf_falseCondition_shouldNotAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", false, i -> i.value)
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var header = wb.getSheetAt(0).getRow(0);
            assertEquals("Name", header.getCell(0).getStringCellValue());
            assertNull(header.getCell(1), "Value column should not exist when condition=false");
        }
    }

    @Test
    void columnIf_trueCondition_shouldAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", true, i -> i.value)
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var header = wb.getSheetAt(0).getRow(0);
            assertEquals("Name", header.getCell(0).getStringCellValue());
            assertEquals("Value", header.getCell(1).getStringCellValue(), "Value column should exist");
        }
    }

    @Test
    void columnIf_withConfig_falseCondition_shouldNotAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", false, i -> i.value,
                            c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var header = wb.getSheetAt(0).getRow(0);
            assertNull(header.getCell(1), "Configured column should not exist when condition=false");
        }
    }

    @Test
    void columnIf_withConfig_trueCondition_shouldAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", true, i -> i.value,
                            c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var header = wb.getSheetAt(0).getRow(0);
            assertEquals("Value", header.getCell(1).getStringCellValue());
            // Data row should have numeric value
            var data = wb.getSheetAt(0).getRow(1);
            assertEquals(1.0, data.getCell(1).getNumericCellValue(), 0.001);
        }
    }

    // ============================================================
    // onProgress invalid interval
    // ============================================================
    @Test
    void onProgress_zeroInterval_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<Item>sheet("Test").column("Name", Item::name);
            assertThrows(IllegalArgumentException.class,
                    () -> sheet.onProgress(0, (count, cursor) -> {}));
        }
    }

    @Test
    void onProgress_negativeInterval_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<Item>sheet("Test").column("Name", Item::name);
            assertThrows(IllegalArgumentException.class,
                    () -> sheet.onProgress(-1, (count, cursor) -> {}));
        }
    }

    // ============================================================
    // defaultStyle — verify bold/fontSize applied via POI
    // ============================================================
    @Test
    void defaultStyle_shouldApplyBoldAndFontSize() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .defaultStyle(d -> d.bold(true).fontSize(14))
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var dataRow = wb.getSheetAt(0).getRow(1);
            var font0 = wb.getFontAt(dataRow.getCell(0).getCellStyle().getFontIndex());
            assertTrue(font0.getBold(), "Name should be bold from defaultStyle");
            assertEquals(14, font0.getFontHeightInPoints(), "Font size should be 14");

            var font1 = wb.getFontAt(dataRow.getCell(1).getCellStyle().getFontIndex());
            assertTrue(font1.getBold(), "Value should also be bold from defaultStyle");
        }
    }

    // ============================================================
    // maxRows validation
    // ============================================================
    @Test
    void maxRows_zero_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<Item>sheet("Test").column("Name", Item::name);
            assertThrows(IllegalArgumentException.class, () -> sheet.maxRows(0));
        }
    }

    @Test
    void maxRows_negative_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<Item>sheet("Test").column("Name", Item::name);
            assertThrows(IllegalArgumentException.class, () -> sheet.maxRows(-1));
        }
    }

    // ============================================================
    // autoWidthSampleRows validation
    // ============================================================
    @Test
    void autoWidthSampleRows_negative_throws() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<Item>sheet("Test").column("Name", Item::name);
            assertThrows(IllegalArgumentException.class, () -> sheet.autoWidthSampleRows(-1));
        }
    }

    @Test
    void autoWidthSampleRows_zero_accepted() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .autoWidthSampleRows(0)
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("A", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue(),
                    "Data should be written even with autoWidthSampleRows=0");
        }
    }
}
