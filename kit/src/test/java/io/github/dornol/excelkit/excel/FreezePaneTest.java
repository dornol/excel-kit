package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("B4: freezePane / freezeRows / freezeCols")
class FreezePaneTest {

    // ── Negative-input guards ────────────────────────────────
    @Test
    void freezePane_negativeCol_shouldThrow() {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> writer.freezePane(-1, 0));
    }

    @Test
    void freezePane_negativeRow_shouldThrow() {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> writer.freezePane(0, -1));
    }

    @Test
    void freezeRows_negative_shouldThrow() {
        ExcelWriter<String> w = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> w.freezeRows(-1));
    }

    @Test
    void freezeCols_negative_shouldThrow() {
        ExcelWriter<String> w = ExcelWriter.<String>create();
        assertThrows(IllegalArgumentException.class, () -> w.freezeCols(-1));
    }

    // ── Positive behavior verified via XSSFWorkbook pane inspection ──
    @Test
    void freezeRows_shouldSetHorizontalSplitOnly() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .freezeRows(1)
                .write(Stream.of("row1"))
                .writeTo(bos);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = wb.getSheetAt(0).getPaneInformation();
            assertNotNull(pane, "freezeRows should create a pane");
            assertEquals(1, pane.getHorizontalSplitPosition(),
                    "freezeRows(1) should split at row 1");
            assertEquals(0, pane.getVerticalSplitPosition(),
                    "freezeRows alone should leave vertical split at 0");
        }
    }

    @Test
    void freezeCols_shouldSetVerticalSplitOnly() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s).column("B", s -> s).column("C", s -> s)
                .freezeCols(2)
                .write(Stream.of("row1"))
                .writeTo(bos);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = wb.getSheetAt(0).getPaneInformation();
            assertNotNull(pane, "freezeCols should create a pane");
            assertEquals(2, pane.getVerticalSplitPosition(),
                    "freezeCols(2) should split at column 2");
            assertEquals(0, pane.getHorizontalSplitPosition(),
                    "freezeCols alone should leave horizontal split at 0");
        }
    }

    @Test
    void freezePane_bothAxes_shouldSetBothSplits() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s).column("B", s -> s).column("C", s -> s)
                .freezePane(2, 1)
                .write(Stream.of("row1"))
                .writeTo(bos);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = wb.getSheetAt(0).getPaneInformation();
            assertNotNull(pane, "freezePane should create a pane");
            assertEquals(2, pane.getVerticalSplitPosition(),
                    "freezePane(2, 1) should split vertically at column 2");
            assertEquals(1, pane.getHorizontalSplitPosition(),
                    "freezePane(2, 1) should split horizontally at row 1");
        }
    }

    @Test
    void noFreeze_shouldHaveNoPane() throws IOException {
        // Guard: absence of freeze* calls must not produce a pane.
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("row1"))
                .writeTo(bos);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = wb.getSheetAt(0).getPaneInformation();
            assertNull(pane, "No freeze* call → sheet must have no pane information");
        }
    }

    // ── ExcelSheetWriter delegates to same config ─────────────
    @Test
    void freezePane_onExcelSheetWriter_appliesToOutput() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            ExcelSheetWriter<String> sw = wb.sheet("test");
            sw.column("Col", row -> row).freezePane(3, 2).write(Stream.of("data"));
            wb.finish().writeTo(bos);
        }
        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = xwb.getSheet("test").getPaneInformation();
            assertNotNull(pane);
            assertEquals(3, pane.getVerticalSplitPosition());
            assertEquals(2, pane.getHorizontalSplitPosition());
        }
    }

    @Test
    void freezeRows_onExcelSheetWriter_appliesToOutput() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("s").column("Col", r -> r).freezeRows(1).write(Stream.of("x"));
            wb.finish().writeTo(bos);
        }
        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = xwb.getSheet("s").getPaneInformation();
            assertNotNull(pane);
            assertEquals(1, pane.getHorizontalSplitPosition());
            assertEquals(0, pane.getVerticalSplitPosition());
        }
    }

    @Test
    void freezeCols_onExcelSheetWriter_appliesToOutput() throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("s").column("A", r -> r).column("B", r -> r)
                    .freezeCols(1).write(Stream.of("x"));
            wb.finish().writeTo(bos);
        }
        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new java.io.ByteArrayInputStream(bos.toByteArray()))) {
            var pane = xwb.getSheet("s").getPaneInformation();
            assertNotNull(pane);
            assertEquals(1, pane.getVerticalSplitPosition());
            assertEquals(0, pane.getHorizontalSplitPosition());
        }
    }

    @Test
    void freezeSheetWriter_negatives_shouldThrow() {
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            ExcelSheetWriter<String> sw = wb.sheet("s");
            assertThrows(IllegalArgumentException.class, () -> sw.freezeRows(-1));
            assertThrows(IllegalArgumentException.class, () -> sw.freezeCols(-1));
        }
    }
}
