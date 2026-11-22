package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** Auto column width counts an ASCII character as one unit and any other character, an emoji included, as two. */
class LogicalWidthTest {

    private static int widthFor(String text) {
        ExcelColumn<String> column = ExcelColumn.of("", (row, cursor) -> row, null,
                ExcelDataType.STRING.getSetter());
        column.fitColumnWidthByValue(text);
        return column.getColumnWidth();
    }

    @Test
    void fourUnitsWhicheverKindOfCharacterMakesThemUp() {
        int fourUnits = 4 * 250 + 1024;
        assertEquals(fourUnits, widthFor("abcd"));
        assertEquals(fourUnits, widthFor("한글"));
        assertEquals(fourUnits, widthFor("😀😀"), "an emoji is one wide character, not two");
    }

    @Test
    void aLoneSurrogateCountsAsOneWideCharacter() {
        assertEquals(2 * 250 + 1024, widthFor("\uD83D"));
    }
}
