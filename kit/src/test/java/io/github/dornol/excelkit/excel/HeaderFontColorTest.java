package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The text of a header is black on a light fill and white on a dark one, and the colour is an RGB value. A palette colour
 * (black or white by index) looks the same without a table, but under a table Excel shows the table style's own header
 * text colour, white, instead: found on a real Excel, where every table header was white on the header's white fill.
 */
class HeaderFontColorTest {

    private record Item(String name) {
    }

    private static XSSFColor headerTextColor(byte[] xlsx, int row) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            XSSFCellStyle style = (XSSFCellStyle) wb.getSheetAt(0).getRow(row).getCell(0).getCellStyle();
            return style.getFont().getXSSFColor();
        }
    }

    private static void assertRgb(XSSFColor color, int expected, String what) {
        assertNotNull(color, what);
        assertTrue(color.isRGB(), what + " must be an RGB colour, was " + color);
        assertEquals(expected, Byte.toUnsignedInt(color.getRGB()[0]), what);
        assertEquals(expected, Byte.toUnsignedInt(color.getRGB()[1]), what);
        assertEquals(expected, Byte.toUnsignedInt(color.getRGB()[2]), what);
    }

    private static byte[] write(boolean table, ExcelColor headerColor) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter<Item> writer = ExcelWriter.<Item>create();
        if (headerColor != null) writer.headerColor(headerColor);
        if (table) writer.table("HeaderTable");
        writer.column("Name", Item::name).write(List.of(new Item("a"))).writeTo(out);
        return out.toByteArray();
    }

    @Test
    void theDefaultHeaderHasBlackRgbTextWithAndWithoutATable() throws Exception {
        assertRgb(headerTextColor(write(false, null), 0), 0, "header text");
        assertRgb(headerTextColor(write(true, null), 0), 0, "header text under a table");
    }

    @Test
    void aDarkHeaderHasWhiteRgbText() throws Exception {
        assertRgb(headerTextColor(write(false, ExcelColor.STEEL_BLUE), 0), 255, "header text on a dark fill");
        assertRgb(headerTextColor(write(true, ExcelColor.STEEL_BLUE), 0), 255, "header text on a dark fill under a table");
    }

    @Test
    void theMultiSheetWriterAndTheTemplateListWriteTheSame() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("S").table("SheetTable").column("Name", Item::name).write(List.of(new Item("a")).stream());
            book.finish().writeTo(out);
        }
        assertRgb(headerTextColor(out.toByteArray(), 0), 0, "header text of ExcelSheetWriter");

        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("T").createRow(0).createCell(0).setCellValue("Title");
            wb.write(template);
        }
        ByteArrayOutputStream filled = new ByteArrayOutputStream();
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template.toByteArray()))) {
            writer.<Item>list(2).column("Name", Item::name).writeWithHeaders(List.of(new Item("a")));
            writer.finish().writeTo(filled);
        }
        assertRgb(headerTextColor(filled.toByteArray(), 2), 0, "header text of a template list");
    }
}
