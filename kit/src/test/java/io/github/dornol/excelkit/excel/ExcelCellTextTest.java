package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;

/**
 * SpreadsheetML cannot carry every character, and reads {@code _xHHHH_} as an escape. Text is encoded so that
 * what was written is what is read: the characters XML cannot hold become escapes, and text that already looks
 * like an escape gets its underscore escaped.
 */
class ExcelCellTextTest {

    @Test
    void controlCharactersAndNoncharactersAreEscapedByTheirFourHexDigits() {
        assertEquals("_x0001_", ExcelCellText.encode("\u0001"));
        assertEquals("_x001F_", ExcelCellText.encode("\u001F"));
        assertEquals("_x000D_", ExcelCellText.encode("\r"));
        assertEquals("_x0000_", ExcelCellText.encode("\u0000"));
        assertEquals("_xFFFE_", ExcelCellText.encode("￾"));
        assertEquals("_xFFFF_", ExcelCellText.encode("￿"));
        assertEquals("a_x0002_b_x0003_c", ExcelCellText.encode("a\u0002b\u0003c"));
    }

    @Test
    void tabAndLineFeedAndOrdinaryTextAreLeftAlone() {
        for (String text : new String[]{"", "plain", "tab\there", "line1\nline2", " ", "é한글", "~\u007F"}) {
            assertSame(text, ExcelCellText.encode(text), text);
        }
    }

    @Test
    void aSurrogatePairIsLeftAloneButALoneHalfIsEscaped() {
        String pair = "😀";
        assertSame(pair, ExcelCellText.encode(pair));
        assertEquals("a" + pair + "b", ExcelCellText.encode("a" + pair + "b"));

        assertEquals("_xD800_", ExcelCellText.encode("\uD800"));              // a high half at the end
        assertEquals("_xD800_a", ExcelCellText.encode("\uD800a"));            // a high half before something else
        assertEquals("_xD800__xD800_", ExcelCellText.encode("\uD800\uD800")); // two high halves
        assertEquals("_xDC00_", ExcelCellText.encode("\uDC00"));              // a low half at the start
        assertEquals("a_xDC00_", ExcelCellText.encode("a\uDC00"));            // a low half after something else
        assertEquals("_xDC00__xDC00_", ExcelCellText.encode("\uDC00\uDC00")); // two low halves
        assertEquals(pair + "_xDC00_", ExcelCellText.encode(pair + "\uDC00")); // a low half after a whole pair
    }

    @Test
    void textThatLooksLikeAnEscapeHasItsUnderscoreEscaped() {
        assertEquals("_x005F_x0041_", ExcelCellText.encode("_x0041_"));
        assertEquals("_x005F_x00af_", ExcelCellText.encode("_x00af_"), "lower case hex digits count too");
        assertEquals("a_x005F_x0041_b", ExcelCellText.encode("a_x0041_b"));
        // an escape that is not complete or not hex is just text
        for (String text : new String[]{"_x004", "_x0041", "_x00", "_x", "_", "_y0041_", "_xZZZZ_", "_x00G1_", "x0041_",
                "_x0041a", "_x0041 "}) {
            assertEquals(text, ExcelCellText.encode(text), text);
        }
    }

    @Test
    void anEscapeThatTheNextEscapeWouldCompleteIsEscapedToo() {
        // "_x0041" followed by a character that becomes "_x0001_" would read as the escape "_x0041_"
        assertEquals("_x005F_x0041_x0001_", ExcelCellText.encode("_x0041\u0001"));
        assertEquals("_x005F_x0041__x005F_x0041_", ExcelCellText.encode("_x0041__x0041_"));
    }

    @Test
    void theWholeEscapeAtTheEndOfTheTextIsRecognised() {
        // exactly seven characters, the last one the closing underscore
        assertEquals("_x005F_x0041_", ExcelCellText.encode("_x0041_"));
        // six characters of it are not enough
        assertEquals("_x0041", ExcelCellText.encode("_x0041"));
    }

    @Test
    void anEncodedTextIsStoredAndANullTextIsABlankCell() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet();
            var row = sheet.createRow(0);

            ExcelCellText.write(row.createCell(0), null);
            assertEquals(CellType.BLANK, row.getCell(0).getCellType());

            ExcelCellText.write(row.createCell(1), "plain");
            assertEquals("plain", row.getCell(1).getStringCellValue());

            ExcelCellText.write(row.createCell(2), "a\u0001b");
            // stored as an escape, read back as the character
            assertEquals("a\u0001b", row.getCell(2).getRichStringCellValue().getString());
        }
    }
}
