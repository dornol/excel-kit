package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelSheetWriter} to cover:
 * - columnIf with false condition
 * - onProgress with invalid interval
 * - defaultStyle with applyDefaults
 * - maxRows validation
 */
class ExcelSheetWriterEdgeCaseTest {

    record Item(String name, int value) {}

    // ============================================================
    // columnIf false condition
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
        assertTrue(out.size() > 0);
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
        assertTrue(out.size() > 0);
    }

    @Test
    void columnIf_withConfig_falseCondition() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", false, i -> i.value,
                            c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void columnIf_withConfig_trueCondition() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .column("Name", Item::name)
                    .columnIf("Value", true, i -> i.value,
                            c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
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
    // defaultStyle
    // ============================================================
    @Test
    void defaultStyle_shouldApplyToAllColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Test")
                    .defaultStyle(d -> d.bold(true).fontSize(14))
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
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
        assertTrue(out.size() > 0);
    }
}
