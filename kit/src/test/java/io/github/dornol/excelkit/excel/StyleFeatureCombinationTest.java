package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class StyleFeatureCombinationTest {
    @Test
    void allFeatures_combined() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c
                        .border(ExcelBorderStyle.MEDIUM)
                        .comment(s -> "Note: " + s)
                        .locked(false))
                .column("Value", s -> 100, c -> c
                        .type(ExcelDataType.INTEGER)
                        .border(ExcelBorderStyle.THICK))
                .protectSheet("pass")
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("50", ExcelColor.LIGHT_GREEN))
                .write(Stream.of("Alice", "Bob"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Protection
            assertNotNull(sheet.getCTWorksheet().getSheetProtection());
            // Comments
            assertNotNull(sheet.getRow(1).getCell(0).getCellComment());
            // Border
            assertEquals(BorderStyle.MEDIUM,
                    sheet.getRow(1).getCell(0).getCellStyle().getBorderTop());
            // Conditional formatting
            assertTrue(sheet.getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
            // Locked
            assertFalse(sheet.getRow(1).getCell(0).getCellStyle().getLocked());
        }
    }

    @Test
    void mapWriter_withProtection() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        var writer = ExcelWriter.forMap("Name", "Age")
                .protectSheet("secret");
        writer.write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertNotNull(wb.getSheetAt(0).getCTWorksheet().getSheetProtection());
        }
    }

    @Test
    void sheetWriter_allNewFeatures() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Protected")
                    .column("Name", s -> s, c -> c
                            .border(ExcelBorderStyle.DASHED)
                            .comment(s -> "Note")
                            .locked(false))
                    .protectSheet("pass")
                    .conditionalFormatting(cf -> cf
                            .equalTo("\"Alice\"", ExcelColor.LIGHT_GREEN))
                    .write(Stream.of("Alice"));
            wb.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertNotNull(sheet.getCTWorksheet().getSheetProtection());
            assertNotNull(sheet.getRow(1).getCell(0).getCellComment());
            assertEquals(BorderStyle.DASHED,
                    sheet.getRow(1).getCell(0).getCellStyle().getBorderTop());
        }
    }
}
