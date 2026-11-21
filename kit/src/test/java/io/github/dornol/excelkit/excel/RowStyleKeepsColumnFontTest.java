package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** A row style changes only what it sets; the rest of the column's font must survive. */
class RowStyleKeepsColumnFontTest {

    private static XSSFFont dataFont(Consumer<RowStyleConfig> rowStyle) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("v", s -> s, c -> c.fontColor(ExcelColor.RED).underline(true)
                        .fontName("Arial").fontSize(14).strikethrough(true))
                .rowStyle(s -> true, rowStyle)
                .write(List.of("x")).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return workbook.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFont();
        }
    }

    @Test
    void boldKeepsRgbColorUnderlineNameSizeAndStrikeout() throws Exception {
        XSSFFont font = dataFont(style -> style.bold(true));

        assertTrue(font.getBold());
        assertArrayEquals(new byte[]{(byte) 255, 0, 0}, font.getXSSFColor().getRGB());
        assertEquals(Font.U_SINGLE, font.getUnderline());
        assertEquals("Arial", font.getFontName());
        assertEquals(14, font.getFontHeightInPoints());
        assertTrue(font.getStrikeout());
    }

    @Test
    void aRowFontColorStillOverridesTheColumnColor() throws Exception {
        XSSFFont font = dataFont(style -> style.fontColor(ExcelColor.BLUE));

        assertArrayEquals(new byte[]{0, 0, (byte) 255}, font.getXSSFColor().getRGB());
        assertEquals(Font.U_SINGLE, font.getUnderline());
    }

    @Test
    void aRowStyleOnlySettingTheBackgroundLeavesTheFontAlone() throws Exception {
        XSSFFont font = dataFont(style -> style.backgroundColor(ExcelColor.LIGHT_YELLOW));

        assertArrayEquals(new byte[]{(byte) 255, 0, 0}, font.getXSSFColor().getRGB());
        assertEquals(Font.U_SINGLE, font.getUnderline());
    }
}
