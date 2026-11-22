package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookProtectionTest {

    @Test
    void excelWriter_protectWorkbook_structureIsLocked() throws Exception {
        var handler = ExcelWriter.<String>create()
                .protectWorkbook("secret")
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertTrue(wb.isStructureLocked(),
                    "Workbook structure should be locked after protectWorkbook");
        }
    }

    @Test
    void excelWorkbook_protectWorkbook_structureIsLocked() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            workbook.protectWorkbook("pwd");
            workbook.<String>sheet("Sheet1")
                    .column("Name", s -> s)
                    .write(Stream.of("Bob"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertTrue(wb.isStructureLocked(),
                    "ExcelWorkbook structure should be locked after protectWorkbook");
        }
    }

    @Test
    void withoutProtectWorkbook_structureIsNotLocked() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertFalse(wb.isStructureLocked(),
                    "Workbook structure should NOT be locked when protectWorkbook is not called");
        }
    }

    @Test
    void protectWorkbook_combinedWithProtectSheet_bothApply() throws Exception {
        var handler = ExcelWriter.<String>create()
                .protectWorkbook("workbookPwd")
                .protectSheet("sheetPwd")
                .column("Name", s -> s)
                .write(Stream.of("Data"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertTrue(wb.isStructureLocked(),
                    "Workbook structure should be locked");
            assertTrue(wb.getSheetAt(0).isSheetLocked(),
                    "Sheet should be protected when protectSheet is also called");
        }
    }
}
