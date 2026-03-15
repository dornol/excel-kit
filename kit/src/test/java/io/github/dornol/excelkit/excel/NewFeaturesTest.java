package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for Features 5 (cellColor), 6 (group header), 7 (auto-rollover), 8 (progress).
 */
class NewFeaturesTest {

    // ========================================================================
    // Feature 8: Progress callback
    // ========================================================================
    @Test
    void progress_shouldFireAtCorrectIntervals() {
        List<Long> progressCounts = new ArrayList<>();

        new ExcelWriter<Integer>()
                .column("Value", i -> i).type(ExcelDataType.INTEGER)
                .onProgress(3, (count, cursor) -> progressCounts.add(count))
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10));

        // Should fire at 3, 6, 9
        assertEquals(List.of(3L, 6L, 9L), progressCounts);
    }

    @Test
    void progress_shouldNotFireWhenIntervalNotReached() {
        List<Long> progressCounts = new ArrayList<>();

        new ExcelWriter<Integer>()
                .column("Value", i -> i).type(ExcelDataType.INTEGER)
                .onProgress(100, (count, cursor) -> progressCounts.add(count))
                .write(Stream.of(1, 2, 3));

        assertTrue(progressCounts.isEmpty());
    }

    @Test
    void progress_shouldWorkInExcelSheetWriter() {
        List<Long> progressCounts = new ArrayList<>();

        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Test")
                    .column("Value", i -> i)
                    .onProgress(2, (count, cursor) -> progressCounts.add(count))
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish();
        }

        assertEquals(List.of(2L, 4L), progressCounts);
    }

    @Test
    void progress_shouldWorkWithColumnBuilderChain() {
        AtomicInteger callCount = new AtomicInteger();

        new ExcelWriter<Integer>()
                .column("A", i -> i)
                .column("B", i -> i * 2)
                .onProgress(5, (count, cursor) -> callCount.incrementAndGet())
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10));

        assertEquals(2, callCount.get()); // fires at 5 and 10
    }

    // ========================================================================
    // Feature 5: Conditional cell color
    // ========================================================================
    @Test
    void cellColor_shouldApplyPerCellColor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Integer>()
                .column("Value", i -> i)
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int v = ((Number) value).intValue();
                        if (v < 0) return ExcelColor.LIGHT_RED;
                        if (v > 100) return ExcelColor.LIGHT_GREEN;
                        return null;
                    })
                .write(Stream.of(-5, 50, 200))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 1 (value=-5) should have red background
            XSSFColor color1 = (XSSFColor) sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color1);
            assertEquals(ExcelColor.LIGHT_RED.getR(), Byte.toUnsignedInt(color1.getRGB()[0]));

            // Row 2 (value=50) should have no fill override (default style)
            // Row 3 (value=200) should have green background
            XSSFColor color3 = (XSSFColor) sheet.getRow(3).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color3);
            assertEquals(ExcelColor.LIGHT_GREEN.getR(), Byte.toUnsignedInt(color3.getRGB()[0]));
        }
    }

    @Test
    void cellColor_shouldOverrideRowColor() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Integer>()
                .rowColor(i -> ExcelColor.LIGHT_YELLOW) // all rows yellow
                .column("Value", i -> i)
                    .type(ExcelDataType.INTEGER)
                    .cellColor((value, row) -> {
                        int v = ((Number) value).intValue();
                        return v < 0 ? ExcelColor.LIGHT_RED : null; // negative → red, else fall through to row color
                    })
                .write(Stream.of(-5, 50))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 1 (value=-5): cellColor=RED takes precedence over rowColor=YELLOW
            XSSFColor color1 = (XSSFColor) sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertEquals(ExcelColor.LIGHT_RED.getR(), Byte.toUnsignedInt(color1.getRGB()[0]));

            // Row 2 (value=50): cellColor=null, falls back to rowColor=YELLOW
            XSSFColor color2 = (XSSFColor) sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertEquals(ExcelColor.LIGHT_YELLOW.getR(), Byte.toUnsignedInt(color2.getRGB()[0]));
        }
    }

    @Test
    void cellColor_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Test")
                    .column("Value", i -> i, c -> c
                            .type(ExcelDataType.INTEGER)
                            .cellColor((value, row) -> {
                                int v = ((Number) value).intValue();
                                return v > 100 ? ExcelColor.LIGHT_GREEN : null;
                            }))
                    .write(Stream.of(50, 200));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 2 (value=200) should have green background
            XSSFColor color = (XSSFColor) sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor();
            assertNotNull(color);
            assertEquals(ExcelColor.LIGHT_GREEN.getR(), Byte.toUnsignedInt(color.getRGB()[0]));
        }
    }

    // ========================================================================
    // Feature 7: ExcelSheetWriter auto-rollover
    // ========================================================================
    @Test
    void rollover_shouldCreateMultipleSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(3)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5, 6, 7));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getNumberOfSheets()); // 3+3+1 = 7 rows
            assertEquals("Data", wb.getSheetName(0));
            assertEquals("Data (2)", wb.getSheetName(1));
            assertEquals("Data (3)", wb.getSheetName(2));

            // Each sheet should have header + data rows
            assertEquals(4, wb.getSheetAt(0).getLastRowNum() + 1); // header + 3 data
            assertEquals(4, wb.getSheetAt(1).getLastRowNum() + 1); // header + 3 data
            assertEquals(2, wb.getSheetAt(2).getLastRowNum() + 1); // header + 1 data
        }
    }

    @Test
    void rollover_shouldUseCustomSheetNameFunction() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Orders")
                    .maxRows(2)
                    .sheetName(idx -> "Orders-Page" + (idx + 1))
                    .column("ID", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getNumberOfSheets());
            assertEquals("Orders", wb.getSheetName(0)); // original name kept
            assertEquals("Orders-Page2", wb.getSheetName(1));
            assertEquals("Orders-Page3", wb.getSheetName(2));
        }
    }

    @Test
    void rollover_withoutMaxRows_shouldWriteToSingleSheet() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Single")
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, wb.getNumberOfSheets());
            assertEquals(6, wb.getSheetAt(0).getLastRowNum() + 1); // header + 5 data
        }
    }

    @Test
    void rollover_shouldNotConflictWithOtherSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("A")
                    .maxRows(2)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3));

            wb.<String>sheet("B")
                    .column("Name", s -> s)
                    .write(Stream.of("x", "y"));

            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getNumberOfSheets()); // A, A (2), B
            assertEquals("A", wb.getSheetName(0));
            assertEquals("A (2)", wb.getSheetName(1));
            assertEquals("B", wb.getSheetName(2));
        }
    }

    // ========================================================================
    // Feature 6: Group header
    // ========================================================================
    @Test
    void groupHeader_shouldCreateMergedGroupRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<int[]>()
                .column("Name", r -> "Item")
                .column("Price", r -> r[0]).type(ExcelDataType.INTEGER).group("Financial")
                .column("Qty", r -> r[1]).type(ExcelDataType.INTEGER).group("Financial")
                .column("Total", r -> r[0] * r[1]).type(ExcelDataType.INTEGER).group("Financial")
                .column("Notes", r -> "note")
                .write(Stream.of(new int[]{100, 5}))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Row 0: group header, Row 1: column header, Row 2: data
            assertEquals(3, sheet.getLastRowNum() + 1);

            // "Name" (col 0) should be vertically merged (rows 0-1)
            // "Financial" (cols 1-3) should be horizontally merged in row 0
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            assertTrue(mergedRegions.size() >= 2);

            // Check "Financial" horizontal merge
            boolean hasFinancialMerge = mergedRegions.stream()
                    .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 0
                            && r.getFirstColumn() == 1 && r.getLastColumn() == 3);
            assertTrue(hasFinancialMerge, "Financial group should be merged across cols 1-3");

            // Check "Name" vertical merge
            boolean hasNameMerge = mergedRegions.stream()
                    .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 1
                            && r.getFirstColumn() == 0 && r.getLastColumn() == 0);
            assertTrue(hasNameMerge, "Name should be vertically merged across rows 0-1");

            // Check "Notes" vertical merge
            boolean hasNotesMerge = mergedRegions.stream()
                    .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 1
                            && r.getFirstColumn() == 4 && r.getLastColumn() == 4);
            assertTrue(hasNotesMerge, "Notes should be vertically merged across rows 0-1");

            // Data row should be at row 2
            assertEquals(100.0, sheet.getRow(2).getCell(1).getNumericCellValue());
        }
    }

    @Test
    void groupHeader_withNoGroups_shouldCreateSingleHeaderRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("A", s -> s)
                .column("B", s -> s)
                .write(Stream.of("test"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals(2, sheet.getLastRowNum() + 1); // header + 1 data
            assertTrue(sheet.getMergedRegions().isEmpty());
        }
    }

    @Test
    void groupHeader_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<int[]>sheet("Test")
                    .column("Name", r -> "Item")
                    .column("Price", r -> r[0], c -> c.type(ExcelDataType.INTEGER).group("Financial"))
                    .column("Qty", r -> r[1], c -> c.type(ExcelDataType.INTEGER).group("Financial"))
                    .write(Stream.of(new int[]{100, 5}));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals(3, sheet.getLastRowNum() + 1); // group header + column header + 1 data
            assertFalse(sheet.getMergedRegions().isEmpty());
        }
    }
}
