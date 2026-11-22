package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class SheetWriterFeatureCombinationTest {

    @Test
    void allNewFeatures_viaSheetWriter_combined() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
            wb.<String>sheet("Styled")
                    .tabColor(ExcelColor.RED)
                    .column("Rotated", s -> s, c -> c
                            .rotation(45)
                            .fontColor(ExcelColor.RED)
                            .strikethrough()
                            .underline()
                            .bold(true)
                            .fontSize(14)
                            .borderTop(ExcelBorderStyle.THICK)
                            .borderBottom(ExcelBorderStyle.THIN)
                            .borderLeft(ExcelBorderStyle.DASHED)
                            .borderRight(ExcelBorderStyle.DOTTED)
                            .backgroundColor(ExcelColor.LIGHT_YELLOW))
                    .column("Validated", s -> Integer.parseInt(s), c -> c
                            .validation(ExcelValidation.integerBetween(1, 100))
                            .type(ExcelDataType.INTEGER))
                    .column("Hidden", s -> s, c -> c.hidden())
                    .column("Dropdown", s -> s, c -> c.dropdown("A", "B", "C"))
                    .afterData(ctx -> {
                        ctx.groupRows(1, 2);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("10", "20", "30"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = xwb.getSheetAt(0);

            // Tab color
            assertNotNull(sheet.getTabColor());
            assertEquals((byte) 255, sheet.getTabColor().getRGB()[0]);

            // Rotation + font styling + borders on first column
            CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
            assertEquals(45, style.getRotation());
            assertEquals(BorderStyle.THICK, style.getBorderTop());
            assertEquals(BorderStyle.THIN, style.getBorderBottom());
            assertEquals(BorderStyle.DASHED, style.getBorderLeft());
            assertEquals(BorderStyle.DOTTED, style.getBorderRight());
            Font font = xwb.getFontAt(style.getFontIndex());
            assertTrue(font.getBold());
            assertTrue(font.getStrikeout());
            assertEquals(Font.U_SINGLE, font.getUnderline());
            assertEquals(14, font.getFontHeightInPoints());

            // Validation + dropdown
            assertTrue(sheet.getDataValidations().size() >= 2);

            // Hidden column
            assertTrue(sheet.isColumnHidden(2));

            // Row grouping
            assertTrue(sheet.getRow(1).getOutlineLevel() > 0);
            assertTrue(sheet.getRow(2).getOutlineLevel() > 0);
        }
    }

    @Test
    void sheetWriter_noConfig_defaultsWork() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Default")
                    .column("Plain", s -> s)
                    .write(Stream.of("a", "b"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = xwb.getSheetAt(0);
            assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("b", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void allNewFeatures_combined() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .tabColor(ExcelColor.STEEL_BLUE)
                .column("Rotated", s -> s, c -> c
                    .rotation(45)
                    .fontColor(ExcelColor.RED)
                    .strikethrough()
                    .underline()
                    .bold(true)
                    .borderTop(ExcelBorderStyle.THICK)
                    .borderBottom(ExcelBorderStyle.THIN))
                .column("Validated", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 100))
                    .backgroundColor(ExcelColor.LIGHT_GREEN))
                .afterData(ctx -> {
                    ctx.groupRows(1, 2);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("10", "20", "30"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);

            // Tab color
            assertNotNull(sheet.getTabColor(), "Tab color should be set");

            // Rotation + font styling on first column
            CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
            assertEquals(45, style.getRotation());
            Font font = wb.getFontAt(style.getFontIndex());
            assertTrue(font.getBold());
            assertTrue(font.getStrikeout());
            assertEquals(Font.U_SINGLE, font.getUnderline());
            assertEquals(BorderStyle.THICK, style.getBorderTop());
            assertEquals(BorderStyle.THIN, style.getBorderBottom());

            // Validation on second column
            assertFalse(sheet.getDataValidations().isEmpty(), "Should have validation rules");

            // Row grouping
            assertEquals(1, sheet.getRow(1).getOutlineLevel(), "Grouped row should have outline level 1");
            assertEquals(1, sheet.getRow(2).getOutlineLevel(), "Grouped row should have outline level 1");
        }
    }
}
