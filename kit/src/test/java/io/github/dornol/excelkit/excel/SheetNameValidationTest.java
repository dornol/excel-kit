package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** A sheet name Excel cannot hold is refused where it is given, not when the first sheet is created. */
class SheetNameValidationTest {

    @Test
    void refusesNamesExcelCannotHoldWhenTheyAreGiven() {
        for (String bad : new String[]{"", "a/b", "a\\b", "a?b", "a*b", "[x]", "a:b", "'quoted", "quoted'"}) {
            assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().sheetName(bad), bad);
        }
    }

    @Test
    void acceptsOrdinaryNames() {
        for (String ok : new String[]{"Data", "2025 매출", "a'b", "x".repeat(40), "A-B_C (1)"}) {
            assertDoesNotThrow(() -> ExcelWriter.<String>create().sheetName(ok), ok);
        }
    }
}
