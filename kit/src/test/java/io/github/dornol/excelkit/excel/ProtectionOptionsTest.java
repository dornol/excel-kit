package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ProtectionOptionsTest {
    @Test
    void protectSheet_allColumnsLockedByDefault() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s)
                .protectSheet("pass")
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNotNull(wb.getSheetAt(0).getCTWorksheet().getSheetProtection());
        }
    }

    @Test
    void protectSheet_withRollover() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().maxRows(3)
                .column("Name", s -> s)
                .protectSheet("pass")
                .write(Stream.of("A", "B", "C", "D", "E", "F", "G"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                assertNotNull(wb.getSheetAt(i).getCTWorksheet().getSheetProtection());
            }
        }
    }

    @Test
    void locked_withDifferentStyles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Editable", s -> s, c -> c.locked(false).backgroundColor(ExcelColor.LIGHT_GREEN))
                .column("ReadOnly", s -> s, c -> c.locked(true).bold(true))
                .protectSheet("pass")
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertFalse(row.getCell(0).getCellStyle().getLocked());
            assertTrue(row.getCell(1).getCellStyle().getLocked());
        }
    }
}
