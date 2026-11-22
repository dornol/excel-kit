package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

class CellDataValueConversionTest {

    // -----------------------------------------------------------------------
    // 3. CellData — asLocalDateTime(format), asLocalDate(format), asLocalTime(format),
    //    asNumber(Locale), asFloat(), asBooleanOrNull(), edge cases
    // -----------------------------------------------------------------------

    @Test
    void cellData_asLocalDateTime_withFormat() {
        CellData cell = new CellData(0, "19/07/2025 14:30:00");
        LocalDateTime result = cell.asLocalDateTime("dd/MM/yyyy HH:mm:ss");
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 0), result);
    }

    @Test
    void cellData_asLocalDateTime_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "   ");
        assertNull(cell.asLocalDateTime("yyyy-MM-dd HH:mm:ss"));
    }

    @Test
    void cellData_asLocalDate_withFormat() {
        CellData cell = new CellData(0, "19.07.2025");
        LocalDate result = cell.asLocalDate("dd.MM.yyyy");
        assertEquals(LocalDate.of(2025, 7, 19), result);
    }

    @Test
    void cellData_asLocalDate_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asLocalDate("yyyy/MM/dd"));
    }

    @Test
    void cellData_asLocalTime_withFormat() {
        CellData cell = new CellData(0, "14:30");
        LocalTime result = cell.asLocalTime("HH:mm");
        assertEquals(LocalTime.of(14, 30), result);
    }

    @Test
    void cellData_asLocalTime_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "  ");
        assertNull(cell.asLocalTime("HH:mm:ss"));
    }

    @Test
    void cellData_asLocalTime_noArg() {
        CellData cell = new CellData(0, "10:15:30");
        LocalTime result = cell.asLocalTime();
        assertEquals(LocalTime.of(10, 15, 30), result);
    }

    @Test
    void cellData_asLocalTime_noArg_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asLocalTime());
    }

    @Test
    void cellData_asNumber_withLocale() {
        CellData cell = new CellData(0, "1,234.56");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(1234.56, result.doubleValue(), 0.01);
    }

    @Test
    void cellData_asNumber_withLocale_blankReturnsNull() {
        CellData cell = new CellData(0, "   ");
        assertNull(cell.asNumber(Locale.US));
    }

    @Test
    void cellData_asNumber_withLocale_invalidThrows() {
        CellData cell = new CellData(0, "not-a-number");
        assertThrows(IllegalArgumentException.class, () -> cell.asNumber(Locale.US));
    }

    @Test
    void cellData_asNumber_withCurrencySymbols() {
        CellData cell = new CellData(0, "$1,000");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(1000, result.longValue());
    }

    @Test
    void cellData_asNumber_withPercentSign() {
        CellData cell = new CellData(0, "50%");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(50, result.longValue());
    }

    @Test
    void cellData_asFloat() {
        CellData cell = new CellData(0, "3.14");
        Float result = cell.asFloat();
        assertNotNull(result);
        assertEquals(3.14f, result, 0.01f);
    }

    @Test
    void cellData_asFloat_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asFloat());
    }

    @Test
    void cellData_asBooleanOrNull_trueValues() {
        assertEquals(Boolean.TRUE, new CellData(0, "true").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "1").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "y").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "YES").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "True").asBooleanOrNull());
    }

    @Test
    void cellData_asBooleanOrNull_falseValues() {
        assertEquals(Boolean.FALSE, new CellData(0, "false").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "0").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "no").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "n").asBooleanOrNull());
    }

    @Test
    void cellData_asBooleanOrNull_blankReturnsNull() {
        assertNull(new CellData(0, "").asBooleanOrNull());
        assertNull(new CellData(0, "   ").asBooleanOrNull());
    }

    @Test
    void cellData_negativeColumnIndex_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () -> new CellData(-1, "test"));
    }

    @Test
    void cellData_nullFormattedValue_becomesEmpty() {
        CellData cell = new CellData(0, null);
        assertEquals("", cell.formattedValue());
    }

    @Test
    void cellData_asDouble() {
        CellData cell = new CellData(0, "2.718");
        Double result = cell.asDouble();
        assertNotNull(result);
        assertEquals(2.718, result, 0.001);
    }

    @Test
    void cellData_asDouble_blankReturnsNull() {
        assertNull(new CellData(0, "").asDouble());
    }

    @Test
    void cellData_asBigDecimal() {
        CellData cell = new CellData(0, "12345.67");
        BigDecimal result = cell.asBigDecimal();
        assertNotNull(result);
        assertEquals(new BigDecimal("12345.67"), result);
    }

    @Test
    void cellData_asBigDecimal_blankReturnsNull() {
        assertNull(new CellData(0, "").asBigDecimal());
    }

    @Test
    void cellData_asInt_outOfRange_shouldThrow() {
        CellData cell = new CellData(0, String.valueOf(Long.MAX_VALUE));
        assertThrows(IllegalArgumentException.class, cell::asInt);
    }

    @Test
    void cellData_isEmpty() {
        assertTrue(new CellData(0, "").isEmpty());
        assertTrue(new CellData(0, "   ").isEmpty());
        assertFalse(new CellData(0, "x").isEmpty());
    }

    @Test
    void cellData_asBoolean_trueValues() {
        assertTrue(new CellData(0, "true").asBoolean());
        assertTrue(new CellData(0, "1").asBoolean());
        assertTrue(new CellData(0, "y").asBoolean());
        assertTrue(new CellData(0, "yes").asBoolean());
    }

    @Test
    void cellData_asBoolean_falseValues() {
        assertFalse(new CellData(0, "false").asBoolean());
        assertFalse(new CellData(0, "0").asBoolean());
        assertFalse(new CellData(0, "").asBoolean());
    }

    @Test
    void cellData_asLong() {
        CellData cell = new CellData(0, "42");
        assertEquals(42L, cell.asLong());
    }

    @Test
    void cellData_asLong_blankReturnsNull() {
        assertNull(new CellData(0, "").asLong());
    }

    @Test
    void cellData_asInt() {
        CellData cell = new CellData(0, "42");
        assertEquals(42, cell.asInt());
    }

    @Test
    void cellData_asInt_blankReturnsNull() {
        assertNull(new CellData(0, "").asInt());
    }

    @Test
    void cellConversionBuilder_localeNullThrows() {
        assertThrows(NullPointerException.class, () -> CellConversionConfig.builder().locale(null));
    }

    @Test
    void cellData_asLocalDateTime_noArg_standardFormats() {
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30),
                new CellData(0, "2025-07-19 14:30").asLocalDateTime());
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 15),
                new CellData(0, "2025-07-19 14:30:15").asLocalDateTime());
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30),
                new CellData(0, "2025/07/19 14:30").asLocalDateTime());
        // ISO_LOCAL_DATE_TIME format
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 15),
                new CellData(0, "2025-07-19T14:30:15").asLocalDateTime());
    }

    @Test
    void cellData_asLocalDateTime_noArg_blankReturnsNull() {
        assertNull(new CellData(0, "").asLocalDateTime());
    }

    @Test
    void cellData_asLocalDate_noArg_standardFormats() {
        assertEquals(LocalDate.of(2025, 7, 19), new CellData(0, "2025-07-19").asLocalDate());
        assertEquals(LocalDate.of(2025, 7, 19), new CellData(0, "2025/07/19").asLocalDate());
    }

    @Test
    void cellData_asLocalDate_noArg_blankReturnsNull() {
        assertNull(new CellData(0, "").asLocalDate());
    }
}
