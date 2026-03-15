package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for API improvements: duplicate column validation, columnIf, autoFilter(boolean),
 * CSV onProgress, empty stream + callbacks, rollover edge cases.
 */
class ImprovementsTest {

    // ========================================================================
    // Duplicate column name validation
    // ========================================================================
    @Test
    void duplicateColumnName_shouldThrowInExcelWriter() {
        var writer = new ExcelWriter<String>()
                .addColumn("Name", s -> s)
                .addColumn("Name", s -> s);
        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void duplicateColumnName_shouldThrowInExcelSheetWriter() {
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            var sheet = wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .column("Name", s -> s);
            assertThrows(ExcelWriteException.class, () -> sheet.write(Stream.of("test")));
        }
    }

    @Test
    void uniqueColumnNames_shouldNotThrow() {
        assertDoesNotThrow(() ->
                new ExcelWriter<String>()
                        .addColumn("Name", s -> s)
                        .addColumn("Age", s -> s)
                        .write(Stream.of("test")));
    }

    // ========================================================================
    // ExcelSheetWriter autoFilter(boolean)
    // ========================================================================
    @Test
    void autoFilterBoolean_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Test")
                    .autoFilter(true)
                    .column("Name", s -> s)
                    .write(Stream.of("A", "B"));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoFilterFalse_shouldNotApplyFilter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Test")
                    .autoFilter(false)
                    .column("Name", s -> s)
                    .write(Stream.of("A"));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }

    // ========================================================================
    // ExcelSheetWriter columnIf()
    // ========================================================================
    @Test
    void columnIf_shouldAddColumnWhenTrue() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .columnIf("Age", true, s -> "30")
                    .write(Stream.of("Alice"));
            wb.finish().consumeOutputStream(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, xwb.getSheetAt(0).getRow(0).getLastCellNum()); // 2 columns
        }
    }

    @Test
    void columnIf_shouldSkipColumnWhenFalse() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .columnIf("Age", false, s -> "30")
                    .write(Stream.of("Alice"));
            wb.finish().consumeOutputStream(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, xwb.getSheetAt(0).getRow(0).getLastCellNum()); // 1 column
        }
    }

    // ========================================================================
    // CSV onProgress
    // ========================================================================
    @Test
    void csvProgress_shouldFireAtCorrectIntervals() {
        List<Long> counts = new ArrayList<>();
        new CsvWriter<Integer>()
                .column("Value", i -> i)
                .onProgress(3, (count, cursor) -> counts.add(count))
                .write(Stream.of(1, 2, 3, 4, 5, 6, 7, 8, 9));

        assertEquals(List.of(3L, 6L, 9L), counts);
    }

    @Test
    void csvProgress_shouldNotFireWhenIntervalNotReached() {
        List<Long> counts = new ArrayList<>();
        new CsvWriter<Integer>()
                .column("Value", i -> i)
                .onProgress(100, (count, cursor) -> counts.add(count))
                .write(Stream.of(1, 2, 3));

        assertTrue(counts.isEmpty());
    }

    // ========================================================================
    // Empty stream + callbacks
    // ========================================================================
    @Test
    void emptyStream_withAfterData_shouldStillCallback() throws IOException {
        boolean[] called = {false};
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Name", s -> s)
                .afterData(ctx -> {
                    called[0] = true;
                    return ctx.getCurrentRow();
                })
                .write(Stream.empty())
                .consumeOutputStream(out);

        assertTrue(called[0], "afterData should be called even with empty stream");
    }

    @Test
    void emptyStream_withBeforeHeader_shouldStillCallback() throws IOException {
        boolean[] called = {false};
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .column("Name", s -> s)
                .beforeHeader(ctx -> {
                    called[0] = true;
                    return ctx.getCurrentRow();
                })
                .write(Stream.empty())
                .consumeOutputStream(out);

        assertTrue(called[0], "beforeHeader should be called even with empty stream");
    }

    // ========================================================================
    // Rollover edge cases
    // ========================================================================
    @Test
    void rollover_maxRowsTwo_shouldCreateMultipleSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().consumeOutputStream(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, xwb.getNumberOfSheets()); // 2+2+1
            assertEquals(3, xwb.getSheetAt(0).getLastRowNum() + 1); // header + 2 data
            assertEquals(3, xwb.getSheetAt(1).getLastRowNum() + 1); // header + 2 data
            assertEquals(2, xwb.getSheetAt(2).getLastRowNum() + 1); // header + 1 data
        }
    }

    @Test
    void rollover_withCallbacks_shouldCallPerSheet() throws IOException {
        List<String> callbackLog = new ArrayList<>();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .beforeHeader(ctx -> {
                        callbackLog.add("beforeHeader");
                        return ctx.getCurrentRow();
                    })
                    .afterData(ctx -> {
                        callbackLog.add("afterData");
                        return ctx.getCurrentRow();
                    })
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().consumeOutputStream(out);
        }

        // 3 sheets: beforeHeader on each, afterData on each (including rollover + final)
        long beforeCount = callbackLog.stream().filter("beforeHeader"::equals).count();
        long afterCount = callbackLog.stream().filter("afterData"::equals).count();
        assertEquals(3, beforeCount, "beforeHeader should be called for each sheet");
        assertEquals(3, afterCount, "afterData should be called for each sheet");
    }

    @Test
    void rollover_emptyStream_shouldNotCreateExtraSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(5)
                    .column("Value", i -> i)
                    .write(Stream.empty());
            wb.finish().consumeOutputStream(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, xwb.getNumberOfSheets());
        }
    }

    @Test
    void rollover_progressShouldContinueAcrossSheets() {
        List<Long> counts = new ArrayList<>();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Integer>sheet("Data")
                    .maxRows(3)
                    .column("Value", i -> i)
                    .onProgress(2, (count, cursor) -> counts.add(count))
                    .write(Stream.of(1, 2, 3, 4, 5, 6, 7));
            wb.finish();
        }

        // Should fire at 2, 4, 6 (total count, not per-sheet)
        assertEquals(List.of(2L, 4L, 6L), counts);
    }
}
