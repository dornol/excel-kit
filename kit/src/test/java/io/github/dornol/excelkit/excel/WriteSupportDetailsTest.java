package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.csv.CsvWriter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Details of what the writers produce: row style fonts, header colors, named ranges, sheet names, CSV errors. */
class WriteSupportDetailsTest {

    // ---- row styles change what they set and nothing else

    private static XSSFFont rowFont(Consumer<RowStyleConfig> style) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("v", s -> s).rowStyle(s -> true, style).write(List.of("x")).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return workbook.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFont();
        }
    }

    private static final byte[] RED = {(byte) 255, 0, 0};

    @Test
    void eachFontSettingOfARowStyleWorksOnItsOwn() throws Exception {
        XSSFFont base = rowFont(style -> style.backgroundColor(ExcelColor.LIGHT_YELLOW));
        assertFalse(base.getBold());
        assertFalse(base.getItalic());
        assertFalse(base.getStrikeout());

        XSSFFont bold = rowFont(style -> style.bold(true));
        assertTrue(bold.getBold());
        assertFalse(bold.getItalic());
        assertFalse(bold.getStrikeout());
        assertEquals(base.getFontHeightInPoints(), bold.getFontHeightInPoints());

        XSSFFont italic = rowFont(style -> style.italic(true));
        assertTrue(italic.getItalic());
        assertFalse(italic.getBold());

        XSSFFont struck = rowFont(style -> style.strikethrough(true));
        assertTrue(struck.getStrikeout());
        assertFalse(struck.getBold());

        XSSFFont large = rowFont(style -> style.fontSize(16f));
        assertEquals(16, large.getFontHeightInPoints());
        assertFalse(large.getBold());

        XSSFFont red = rowFont(style -> style.fontColor(ExcelColor.RED));
        assertArrayEquals(RED, red.getXSSFColor().getRGB());
        assertFalse(red.getBold());
        assertEquals(base.getFontHeightInPoints(), red.getFontHeightInPoints());
    }

    @Test
    void aRowStyleThatSetsNoFontAttributeLeavesTheFontAsItWas() throws Exception {
        XSSFFont plain = rowFont(style -> { });
        XSSFFont background = rowFont(style -> style.backgroundColor(ExcelColor.LIGHT_RED));
        assertEquals(plain.getFontName(), background.getFontName());
        assertEquals(plain.getFontHeightInPoints(), background.getFontHeightInPoints());
        assertEquals(plain.getBold(), background.getBold());
    }

    // ---- per-column header colors

    private static XSSFCellStyle header(Consumer<ColumnStyleConfig<String, ?>> column) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().headerColor(ExcelColor.LIGHT_YELLOW).column("v", s -> s, column::accept)
                .write(List.of("x")).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
        }
    }

    private static byte[] rgb(ExcelColor color) {
        return new byte[]{(byte) color.getR(), (byte) color.getG(), (byte) color.getB()};
    }

    private static byte[] fill(CellStyle style) {
        return ((XSSFColor) style.getFillForegroundColorColor()).getRGB();
    }

    @Test
    void aColumnCanHaveItsOwnHeaderColors() throws Exception {
        XSSFCellStyle both = header(c -> c.headerBackgroundColor(255, 0, 0).headerFontColor(0, 0, 255));
        assertEquals(FillPatternType.SOLID_FOREGROUND, both.getFillPattern());
        assertArrayEquals(RED, fill(both));
        assertArrayEquals(new byte[]{0, 0, (byte) 255}, both.getFont().getXSSFColor().getRGB());
        assertTrue(both.getFont().getBold(), "the rest of the header font is kept");
        assertEquals(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, both.getAlignment());

        XSSFCellStyle background = header(c -> c.headerBackgroundColor(255, 0, 0));
        assertEquals(FillPatternType.SOLID_FOREGROUND, background.getFillPattern());
        assertArrayEquals(RED, fill(background));

        XSSFCellStyle font = header(c -> c.headerFontColor(0, 0, 255));
        assertArrayEquals(new byte[]{0, 0, (byte) 255}, font.getFont().getXSSFColor().getRGB());
        assertArrayEquals(rgb(ExcelColor.LIGHT_YELLOW), fill(font), "the workbook's header color stays");

        XSSFCellStyle plain = header(c -> { });
        assertArrayEquals(rgb(ExcelColor.LIGHT_YELLOW), fill(plain));
    }

    // ---- named ranges, sheet names

    @Test
    void aNamedRangeNeedsDataRows() throws Exception {
        ByteArrayOutputStream empty = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("v", s -> s).namedRange("Values", 0).write(List.<String>of()).writeTo(empty);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(empty.toByteArray()))) {
            assertNull(workbook.getName("Values"), "no data rows, no range");
        }

        ByteArrayOutputStream one = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("v", s -> s).namedRange("Values", 0).write(List.of("a")).writeTo(one);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(one.toByteArray()))) {
            assertTrue(workbook.getName("Values").getRefersToFormula().endsWith("$A$2:$A$2"),
                    workbook.getName("Values").getRefersToFormula());
        }
    }

    @Test
    void aRolloverSheetNameNeverSplitsASurrogatePair() {
        String emoji = "😀";
        // the cut falls between the two halves of the emoji at index 26/27: the whole emoji goes
        String base = "a".repeat(26) + emoji + "bbb";
        assertEquals(31, base.length());
        assertEquals("a".repeat(26) + " (2)", ExcelWorkbookSupport.rolloverSheetName(base, 1));
        // the cut falls after the pair: it stays
        String roomy = "a".repeat(25) + emoji + "bbbbb";
        assertEquals("a".repeat(25) + emoji + " (2)", ExcelWorkbookSupport.rolloverSheetName(roomy, 1));
        // the first sheet keeps its name, even a long one
        assertEquals(base, ExcelWorkbookSupport.rolloverSheetName(base, 0));
        // a short name is not cut
        assertEquals("Data (3)", ExcelWorkbookSupport.rolloverSheetName("Data", 2));
        // the longest name still fits with the suffix
        assertEquals(31, ExcelWorkbookSupport.rolloverSheetName("x".repeat(40), 1).length());
        assertEquals(31, ExcelWorkbookSupport.rolloverSheetName("x".repeat(40), 99).length());
    }

    // ---- where a CSV write failed

    @Test
    void aFailureFoundWhileWritingARowSaysWhichRow() {
        CsvWriteException error = assertThrows(CsvWriteException.class, () -> CsvWriter.<String>create()
                .column("v", s -> s).write(List.of("fine", "bad\uD800")));
        assertTrue(error.getMessage().matches("(?s).*detected at row \\d+.*"), error.getMessage());
    }

    @Test
    void aFailureFoundOnlyAtTheEndHasNoRowToNameAndStillSaysWhat() {
        // the header is bad and there are no rows after it to notice
        CsvWriteException error = assertThrows(CsvWriteException.class, () -> CsvWriter.<String>create()
                .column("bad\uD800", s -> s).write(List.<String>of()));
        assertFalse(error.getMessage().contains("detected at row"), error.getMessage());
        assertTrue(error.getMessage().startsWith("Failed to write CSV using UTF-8"), error.getMessage());
    }
}
