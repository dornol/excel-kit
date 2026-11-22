package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class HeaderStyleDetailsTest {

    private static CellStyle headerStyle(Consumer<ExcelWriter<String>> configure) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter<String> writer = ExcelWriter.create();
        writer.column("v", s -> s);
        configure.accept(writer);
        writer.write(List.of("x")).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
        }
    }

    @Test
    void theHeaderIsCenteredBoldAndBorderedByDefault() throws Exception {
        CellStyle style = headerStyle(w -> { });
        assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
        assertEquals(VerticalAlignment.CENTER, style.getVerticalAlignment());
        assertTrue(((XSSFFont) ((org.apache.poi.xssf.usermodel.XSSFCellStyle) style).getFont()).getBold());
        assertEquals(BorderStyle.THIN, style.getBorderTop());
        assertEquals(BorderStyle.THIN, style.getBorderLeft());
        assertFalse(style.getWrapText());
    }

    @Test
    void headerStyleSettingsReplaceTheDefaults() throws Exception {
        CellStyle style = headerStyle(w -> w.headerStyle(h -> h.alignment(HorizontalAlignment.LEFT)
                .verticalAlignment(VerticalAlignment.TOP).bold(false).border(ExcelBorderStyle.MEDIUM).wrapText(true)));
        assertEquals(HorizontalAlignment.LEFT, style.getAlignment());
        assertEquals(VerticalAlignment.TOP, style.getVerticalAlignment());
        assertFalse(((org.apache.poi.xssf.usermodel.XSSFCellStyle) style).getFont().getBold());
        assertEquals(BorderStyle.MEDIUM, style.getBorderBottom());
        assertEquals(BorderStyle.MEDIUM, style.getBorderRight());
        assertTrue(style.getWrapText());
    }

    @Test
    void settingsLeftUnsetKeepTheirDefaults() throws Exception {
        CellStyle style = headerStyle(w -> w.headerStyle(h -> h.wrapText(true)));
        assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
        assertEquals(VerticalAlignment.CENTER, style.getVerticalAlignment());
        assertTrue(((org.apache.poi.xssf.usermodel.XSSFCellStyle) style).getFont().getBold());
        assertEquals(BorderStyle.THIN, style.getBorderTop());
    }

    @Test
    void theHeaderFontIsWhiteOnADarkHeaderAndBlackOnALightOne() throws Exception {
        for (ExcelColor dark : List.of(ExcelColor.NAVY, ExcelColor.TEAL, ExcelColor.of(10, 10, 10))) {
            assertEquals(WHITE, fontColor(dark), "on " + dark.getR() + "," + dark.getG() + "," + dark.getB());
        }
        for (ExcelColor light : List.of(ExcelColor.LIGHT_YELLOW, ExcelColor.GOLD, ExcelColor.of(250, 250, 250))) {
            assertEquals(BLACK, fontColor(light), "on " + light.getR() + "," + light.getG() + "," + light.getB());
        }
    }

    @Test
    void theLuminanceThresholdIsAroundMidGrey() throws Exception {
        // W3C brightness 0.299R + 0.587G + 0.114B against 128. (Grey 128 itself is not pinned: the three
        // weights add up to 1 only to within floating point error.)
        assertEquals(WHITE, fontColor(ExcelColor.of(127, 127, 127)));
        assertEquals(BLACK, fontColor(ExcelColor.of(129, 129, 129)));
    }

    private static final String WHITE = "FFFFFFFF";
    private static final String BLACK = "FF000000";

    /** The header text colour as ARGB hex; it must be an RGB colour, not a palette index (see {@code HeaderFontColorTest}). */
    private static String fontColor(ExcelColor header) throws Exception {
        CellStyle style = headerStyle(w -> w.headerColor(header));
        var color = ((org.apache.poi.xssf.usermodel.XSSFCellStyle) style).getFont().getXSSFColor();
        assertTrue(color != null && color.isRGB(), "the header text colour must be an RGB colour, was " + color);
        return color.getARGBHex();
    }

    @org.junit.jupiter.api.io.TempDir
    Path tempDir;

}
