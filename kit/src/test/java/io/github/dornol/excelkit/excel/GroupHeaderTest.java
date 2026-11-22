package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class GroupHeaderTest {

    // ========================================================================
    // Feature 6: Group header
    // ========================================================================
    @Test
    void groupHeader_shouldCreateMergedGroupRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<int[]>create()
                .column("Name", r -> "Item")
                .column("Price", r -> r[0], cfg -> cfg.type(ExcelDataType.INTEGER).group("Financial"))
                .column("Qty", r -> r[1], cfg -> cfg.type(ExcelDataType.INTEGER).group("Financial"))
                .column("Total", r -> r[0] * r[1], cfg -> cfg.type(ExcelDataType.INTEGER).group("Financial"))
                .column("Notes", r -> "note")
                .write(Stream.of(new int[]{100, 5}))
                .writeTo(out);

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
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s)
                .write(Stream.of("test"))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals(2, sheet.getLastRowNum() + 1); // header + 1 data
            assertTrue(sheet.getMergedRegions().isEmpty());
        }
    }

    @Test
    void groupHeader_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<int[]>sheet("Test")
                    .column("Name", r -> "Item")
                    .column("Price", r -> r[0], c -> c.type(ExcelDataType.INTEGER).group("Financial"))
                    .column("Qty", r -> r[1], c -> c.type(ExcelDataType.INTEGER).group("Financial"))
                    .write(Stream.of(new int[]{100, 5}));
            wb.finish().writeTo(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals(3, sheet.getLastRowNum() + 1); // group header + column header + 1 data
            assertFalse(sheet.getMergedRegions().isEmpty());
        }
    }
}
