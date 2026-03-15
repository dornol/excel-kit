package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class SheetProtectionTest {

    private boolean isProtected(XSSFSheet sheet) {
        // XSSFSheet.isSheetProtectionEnabled() is package-private in some POI versions
        // Use reflection or check if the sheet is locked
        try {
            var method = XSSFSheet.class.getDeclaredMethod("isSheetProtectionEnabled");
            method.setAccessible(true);
            return (boolean) method.invoke(sheet);
        } catch (Exception e) {
            // Fallback: check if sheet has protection via the CT model
            return sheet.getCTWorksheet().getSheetProtection() != null;
        }
    }

    @Test
    void protectSheet_shouldEnableProtection() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s)
                .protectSheet("password123")
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(isProtected(wb.getSheetAt(0)));
        }
    }

    @Test
    void protectSheet_lockedColumn_shouldLockCells() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Locked", s -> s, c -> c.locked(true))
                .addColumn("Unlocked", s -> s, c -> c.locked(false))
                .protectSheet("password")
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(isProtected(sheet));
            assertTrue(sheet.getRow(1).getCell(0).getCellStyle().getLocked());
            assertFalse(sheet.getRow(1).getCell(1).getCellStyle().getLocked());
        }
    }

    @Test
    void protectSheet_withoutProtection_lockedFlagOnStyle() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Col", s -> s, c -> c.locked(false))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertFalse(isProtected(wb.getSheetAt(0)));
            assertFalse(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getLocked());
        }
    }

    @Test
    void protectSheet_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Protected")
                    .column("Name", s -> s)
                    .protectSheet("pass")
                    .write(Stream.of("Alice"));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(isProtected(wb.getSheetAt(0)));
        }
    }

    @Test
    void protectSheet_multipleSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>(5)
                .addColumn("Name", s -> s)
                .protectSheet("password")
                .write(Stream.of("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                assertTrue(isProtected(wb.getSheetAt(i)),
                        "Sheet " + i + " should be protected");
            }
        }
    }
}
