package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReadSupportTest {

    @ParameterizedTest
    @CsvSource({
            "A1, 0",
            "B3, 1",
            "Z1, 25",
            "AA1, 26",
            "AZ1, 51",
            "BA1, 52",
            "XFD1, 16383",
    })
    void getColumnIndex_validReferences(String cellRef, int expected) {
        assertEquals(expected, ExcelReadSupport.getColumnIndex(cellRef));
    }

    @Test
    void getColumnIndex_lowercase_treatedAsUppercase() {
        assertEquals(0, ExcelReadSupport.getColumnIndex("a1"));
        assertEquals(26, ExcelReadSupport.getColumnIndex("aa1"));
    }

    @Test
    void getColumnIndex_exceedsMaxColumn_throws() {
        // XFE1 = 16,385th column, exceeds Excel max (16,384)
        assertThrows(ExcelReadException.class,
                () -> ExcelReadSupport.getColumnIndex("XFE1"));
    }
}
