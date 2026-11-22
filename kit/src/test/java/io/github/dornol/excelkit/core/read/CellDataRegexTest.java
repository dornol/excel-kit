package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.testsupport.DefaultLocaleShowsEnglishNumbers;
import io.github.dornol.excelkit.excel.*;
import org.junit.jupiter.api.*;

import java.io.*;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("A1: CellData currency symbol stripping")
class CellDataRegexTest {

    @Test
    void asNumber_shouldStripDollarSign() {
        assertEquals(1234, new CellData(0, "$1234").asInt());
    }

    @Test
    void asNumber_shouldStripEuroSign() {
        assertEquals(999, new CellData(0, "€999").asInt());
    }

    @Test
    void asNumber_shouldStripKoreanWon() {
        assertEquals(5000, new CellData(0, "₩5000").asInt());
    }

    @Test
    void asNumber_shouldStripKoreanWonText() {
        assertEquals(3000, new CellData(0, "3000원").asInt());
    }

    @Test
    void asNumber_shouldStripPercentSign() {
        assertNotNull(new CellData(0, "50%").asNumber());
    }

    @Test
    @DefaultLocaleShowsEnglishNumbers
    void asNumber_shouldStripCommasAndCurrency() {
        assertEquals(1234567, new CellData(0, "$1,234,567").asInt());
    }

    @Test
    void asNumber_shouldStripNonBreakingSpace() {
        assertEquals(100, new CellData(0, "100\u00A0").asInt());
    }

    @Test
    @DefaultLocaleShowsEnglishNumbers
    void asNumber_shouldHandleMixedCurrencySymbols() {
        assertEquals(1234, new CellData(0, "₩1,234").asInt());
    }
}
