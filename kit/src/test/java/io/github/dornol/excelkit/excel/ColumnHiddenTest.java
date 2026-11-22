package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ColumnHiddenTest {

    @Test
    void hidden_noArgs_columnIsHidden() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Visible", s -> s)
                .column("Secret", s -> "hidden-value", c -> c.hidden())
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertFalse(sheet.isColumnHidden(0), "First column should not be hidden");
            assertTrue(sheet.isColumnHidden(1), "Second column should be hidden when hidden() is called");
        }
    }

    @Test
    void hidden_true_columnIsHidden() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c.hidden(true))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).isColumnHidden(0),
                    "Column should be hidden when hidden(true) is called");
        }
    }

    @Test
    void hidden_false_columnIsNotHidden() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c.hidden(false))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(wb.getSheetAt(0).isColumnHidden(0),
                    "Column should not be hidden when hidden(false) is called");
        }
    }

    @Test
    void hidden_mixedColumns_someHiddenSomeNot() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s, c -> c.hidden(true))
                .column("C", s -> s)
                .column("D", s -> s, c -> c.hidden())
                .column("E", s -> s, c -> c.hidden(false))
                .write(Stream.of("row1"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertFalse(sheet.isColumnHidden(0), "Column A should be visible");
            assertTrue(sheet.isColumnHidden(1), "Column B should be hidden");
            assertFalse(sheet.isColumnHidden(2), "Column C should be visible");
            assertTrue(sheet.isColumnHidden(3), "Column D should be hidden");
            assertFalse(sheet.isColumnHidden(4), "Column E should be visible");
        }
    }

    @Test
    void hidden_withExcelSheetWriter_columnConfigHidden() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Sheet1")
                    .column("Visible", s -> s)
                    .column("Hidden", s -> "secret", c -> c.hidden())
                    .write(Stream.of("data"));
            workbook.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertFalse(sheet.isColumnHidden(0), "First column should be visible in ExcelSheetWriter");
            assertTrue(sheet.isColumnHidden(1), "Hidden column via ColumnConfig.hidden() should be hidden");
        }
    }

    @Test
    void hidden_columnStillHasData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .column("Secret", s -> "hidden-" + s, c -> c.hidden())
                .write(Stream.of("Alice"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.isColumnHidden(1), "Column should be hidden");
            assertEquals("hidden-Alice", sheet.getRow(1).getCell(1).getStringCellValue(),
                    "Hidden column should still contain data");
        }
    }

    @Test
    void hidden_withSheetRollover_allSheetsHaveHiddenColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Data")
                    .column("Visible", s -> s)
                    .column("Hidden", s -> "h", c -> c.hidden())
                    .maxRows(3)
                    .write(Stream.of("A", "B", "C", "D", "E", "F", "G"));
            workbook.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getNumberOfSheets() >= 2,
                    "Should have multiple sheets from rollover");
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                var sheet = wb.getSheetAt(i);
                assertFalse(sheet.isColumnHidden(0),
                        "Visible column should not be hidden on sheet " + i);
                assertTrue(sheet.isColumnHidden(1),
                        "Hidden column should be hidden on all rollover sheets, sheet " + i);
            }
        }
    }

    @Test
    void hidden_combinedWithOtherStyling() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Styled Hidden", s -> s, c -> c
                        .hidden()
                        .bold(true)
                        .backgroundColor(ExcelColor.LIGHT_BLUE))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.isColumnHidden(0),
                    "Column should be hidden even with other styling applied");
            // Data should still be present
            assertEquals("data", sheet.getRow(1).getCell(0).getStringCellValue());
        }
    }
}
