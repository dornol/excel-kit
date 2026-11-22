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

class BorderTest {
    @Test
    void border_allStyles_shouldCompileAndWrite() throws IOException {
        for (ExcelBorderStyle style : ExcelBorderStyle.values()) {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Col", s -> s, c -> c.border(style))
                    .write(Stream.of("data"))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var cellStyle = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(ExcelStyleSupporter.toPoiBorderStyle(style), cellStyle.getBorderTop(),
                        "Border style " + style + " should be applied");
            }
        }
    }

    @Test
    void border_mixedStyles_perColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Thin", s -> s, c -> c.border(ExcelBorderStyle.THIN))
                .column("Thick", s -> s, c -> c.border(ExcelBorderStyle.THICK))
                .column("None", s -> s, c -> c.border(ExcelBorderStyle.NONE))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertEquals(BorderStyle.THIN, row.getCell(0).getCellStyle().getBorderTop());
            assertEquals(BorderStyle.THICK, row.getCell(1).getCellStyle().getBorderTop());
            assertEquals(BorderStyle.NONE, row.getCell(2).getCellStyle().getBorderTop());
        }
    }

    @Test
    void border_withOtherStyling_shouldCombine() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c
                        .border(ExcelBorderStyle.MEDIUM)
                        .bold(true)
                        .fontSize(14)
                        .backgroundColor(ExcelColor.LIGHT_BLUE))
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.MEDIUM, style.getBorderTop());
        }
    }
}
