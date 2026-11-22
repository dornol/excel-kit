package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class SheetRolloverTest {

    // ========================================================================
    // Rollover edge cases
    // ========================================================================
    @Test
    void rollover_maxRowsTwo_shouldCreateMultipleSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Data")
                    .maxRows(2)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().writeTo(out);
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
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
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
            wb.finish().writeTo(out);
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
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Data")
                    .maxRows(5)
                    .column("Value", i -> i)
                    .write(Stream.empty());
            wb.finish().writeTo(out);
        }

        try (var xwb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, xwb.getNumberOfSheets());
        }
    }

    @Test
    void rollover_progressShouldContinueAcrossSheets() {
        List<Long> counts = new ArrayList<>();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Data")
                    .maxRows(3)
                    .column("Value", i -> i)
                    .onProgress(2, (count, cursor) -> counts.add(count))
                    .write(Stream.of(1, 2, 3, 4, 5, 6, 7));
            wb.finish().writeTo(java.io.OutputStream.nullOutputStream());
        }

        // Should fire at 2, 4, 6 (total count, not per-sheet)
        assertEquals(List.of(2L, 4L, 6L), counts);
    }

    // ========================================================================
    // Feature 7: ExcelSheetWriter auto-rollover
    // ========================================================================
    @Test
    void rollover_shouldCreateMultipleSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Data")
                    .maxRows(3)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5, 6, 7));
            wb.finish().writeTo(out);
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
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Orders")
                    .maxRows(2)
                    .sheetName(idx -> "Orders-Page" + (idx + 1))
                    .column("ID", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().writeTo(out);
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
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("Single")
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3, 4, 5));
            wb.finish().writeTo(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, wb.getNumberOfSheets());
            assertEquals(6, wb.getSheetAt(0).getLastRowNum() + 1); // header + 5 data
        }
    }

    @Test
    void rollover_shouldNotConflictWithOtherSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Integer>sheet("A")
                    .maxRows(2)
                    .column("Value", i -> i)
                    .write(Stream.of(1, 2, 3));

            wb.<String>sheet("B")
                    .column("Name", s -> s)
                    .write(Stream.of("x", "y"));

            wb.finish().writeTo(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getNumberOfSheets()); // A, A (2), B
            assertEquals("A", wb.getSheetName(0));
            assertEquals("A (2)", wb.getSheetName(1));
            assertEquals("B", wb.getSheetName(2));
        }
    }
}
