package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class IndividualBorderTest {

    @Test
    void perSideBorders_allDifferentStyles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Borders", s -> s, c -> c
                    .borderTop(ExcelBorderStyle.THICK)
                    .borderBottom(ExcelBorderStyle.THIN)
                    .borderLeft(ExcelBorderStyle.DASHED)
                    .borderRight(ExcelBorderStyle.DOTTED))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.THICK, style.getBorderTop());
            assertEquals(BorderStyle.THIN, style.getBorderBottom());
            assertEquals(BorderStyle.DASHED, style.getBorderLeft());
            assertEquals(BorderStyle.DOTTED, style.getBorderRight());
        }
    }

    @Test
    void perSideBorders_partialOverride_uniformFallback() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Borders", s -> s, c -> c
                    .border(ExcelBorderStyle.MEDIUM)
                    .borderTop(ExcelBorderStyle.THICK))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.THICK, style.getBorderTop(), "Top should use per-side override");
            assertEquals(BorderStyle.MEDIUM, style.getBorderBottom(), "Bottom should use uniform border");
            assertEquals(BorderStyle.MEDIUM, style.getBorderLeft(), "Left should use uniform border");
            assertEquals(BorderStyle.MEDIUM, style.getBorderRight(), "Right should use uniform border");
        }
    }

    @Test
    void perSideBorders_noUniform_defaultThin() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Borders", s -> s, c -> c
                    .borderTop(ExcelBorderStyle.DOUBLE))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.DOUBLE, style.getBorderTop(), "Top should use per-side");
            assertEquals(BorderStyle.THIN, style.getBorderBottom(), "Bottom should default to THIN");
            assertEquals(BorderStyle.THIN, style.getBorderLeft(), "Left should default to THIN");
            assertEquals(BorderStyle.THIN, style.getBorderRight(), "Right should default to THIN");
        }
    }

    @Test
    void perSideBorders_noneOnSomeSides() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Borders", s -> s, c -> c
                    .borderTop(ExcelBorderStyle.NONE)
                    .borderBottom(ExcelBorderStyle.NONE)
                    .borderLeft(ExcelBorderStyle.THICK)
                    .borderRight(ExcelBorderStyle.THICK))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.NONE, style.getBorderTop());
            assertEquals(BorderStyle.NONE, style.getBorderBottom());
            assertEquals(BorderStyle.THICK, style.getBorderLeft());
            assertEquals(BorderStyle.THICK, style.getBorderRight());
        }
    }

    @Test
    void perSideBorders_viaAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Col", s -> s, c -> c
                        .borderTop(ExcelBorderStyle.MEDIUM)
                        .borderBottom(ExcelBorderStyle.DASHED))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.MEDIUM, style.getBorderTop());
            assertEquals(BorderStyle.DASHED, style.getBorderBottom());
        }
    }

    @Test
    void perSideBorders_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Col", s -> s, c -> c
                            .borderTop(ExcelBorderStyle.DOUBLE)
                            .borderBottom(ExcelBorderStyle.HAIR))
                    .write(Stream.of("test"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            CellStyle style = xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
            assertEquals(BorderStyle.DOUBLE, style.getBorderTop());
            assertEquals(BorderStyle.HAIR, style.getBorderBottom());
            assertEquals(BorderStyle.THIN, style.getBorderLeft(), "Default THIN for unset sides");
            assertEquals(BorderStyle.THIN, style.getBorderRight(), "Default THIN for unset sides");
        }
    }

    @Test
    void perSideBorders_multipleColumns_mixedStyles() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s, c -> c
                    .borderTop(ExcelBorderStyle.THICK))
                .column("B", s -> s, c -> c
                    .border(ExcelBorderStyle.MEDIUM))
                .column("C", s -> s, c -> c
                    .borderLeft(ExcelBorderStyle.DOTTED)
                    .borderRight(ExcelBorderStyle.DOTTED))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertEquals(BorderStyle.THICK, row.getCell(0).getCellStyle().getBorderTop());
            assertEquals(BorderStyle.MEDIUM, row.getCell(1).getCellStyle().getBorderTop());
            assertEquals(BorderStyle.DOTTED, row.getCell(2).getCellStyle().getBorderLeft());
            assertEquals(BorderStyle.DOTTED, row.getCell(2).getCellStyle().getBorderRight());
            assertEquals(BorderStyle.THIN, row.getCell(2).getCellStyle().getBorderTop());
        }
    }
}
