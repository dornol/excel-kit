package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link io.github.dornol.excelkit.shared.CellData} class.
 */
class CellDataTest {

    // Constructor tests
    @Test
    void constructor_shouldNormalizeNullFormattedValue() {
        // Arrange & Act
        CellData cellData = new CellData(0, null);

        // Assert
        assertEquals("", cellData.formattedValue(), "formattedValue should be empty string when null is provided");
    }

    @Test
    void constructor_shouldThrowExceptionWhenColumnIndexIsNegative() {
        // Arrange
        int columnIndex = -1;
        String formattedValue = "test";

        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            new CellData(columnIndex, formattedValue);
        }, "Constructor should throw IllegalArgumentException when columnIndex is negative");
    }

    @Test
    void constructor_shouldCreateInstanceWithValidParameters() {
        // Arrange
        int columnIndex = 0;
        String formattedValue = "test";

        // Act
        CellData cellData = new CellData(columnIndex, formattedValue);

        // Assert
        assertEquals(columnIndex, cellData.columnIndex(), "columnIndex should match the provided value");
        assertEquals(formattedValue, cellData.formattedValue(), "formattedValue should match the provided value");
    }

    // asNumber tests
    @Test
    void asNumber_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertNull(result, "asNumber should return null when value is empty");
    }

    @Test
    void asNumber_shouldReturnNullWhenValueIsBlank() {
        // Arrange
        CellData cellData = new CellData(0, "   ");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertNull(result, "asNumber should return null when value is blank");
    }

    @Test
    void asNumber_shouldParseSimpleNumber() {
        // Arrange
        CellData cellData = new CellData(0, "123");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertEquals(123, result.intValue(), "asNumber should parse simple number correctly");
    }

    @Test
    void asNumber_shouldParseNumberWithCommas() {
        // Arrange
        CellData cellData = new CellData(0, "1,234,567");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertEquals(1234567, result.intValue(), "asNumber should parse number with commas correctly");
    }

    @Test
    void asNumber_shouldParseNumberWithCurrencySymbols() {
        // Arrange
        CellData cellData = new CellData(0, "$1234.56");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertEquals(1234.56, result.doubleValue(), 0.001, "asNumber should parse number with currency symbols correctly");
    }

    @Test
    void asNumber_shouldParseNumberWithKoreanWon() {
        // Arrange
        CellData cellData = new CellData(0, "1234원");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertEquals(1234, result.intValue(), "asNumber should parse number with Korean won symbol correctly");
    }

    @Test
    void asNumber_shouldParseNumberWithPercentSign() {
        // Arrange
        CellData cellData = new CellData(0, "12.34%");

        // Act
        Number result = cellData.asNumber();

        // Assert
        assertEquals(12.34, result.doubleValue(), 0.001, "asNumber should parse number with percent sign correctly");
    }

    @Test
    void asNumber_shouldThrowExceptionWhenValueCannotBeParsed() {
        // Arrange
        CellData cellData = new CellData(0, "not a number");

        // Act & Assert
        assertThrows(IllegalArgumentException.class, cellData::asNumber, 
                "asNumber should throw IllegalArgumentException when value cannot be parsed");
    }

    @Test
    void asNumber_shouldUseSpecifiedLocale() {
        // Arrange
        // Note: The CellData.asNumber() method removes commas before parsing,
        // so we need to use a format that works with the implementation
        CellData cellData = new CellData(0, "1234.56");

        // Act
        Number result = cellData.asNumber(Locale.US);

        // Assert
        assertEquals(1234.56, result.doubleValue(), 0.001, 
                "asNumber should parse number according to specified locale");
    }

    // asLong tests
    @Test
    void asLong_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        Long result = cellData.asLong();

        // Assert
        assertNull(result, "asLong should return null when value is empty");
    }

    @Test
    void asLong_shouldConvertNumberToLong() {
        // Arrange
        CellData cellData = new CellData(0, "123456789");

        // Act
        Long result = cellData.asLong();

        // Assert
        assertEquals(123456789L, result, "asLong should convert number to Long correctly");
    }

    // asInt tests
    @Test
    void asInt_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        Integer result = cellData.asInt();

        // Assert
        assertNull(result, "asInt should return null when value is empty");
    }

    @Test
    void asInt_shouldConvertNumberToInt() {
        // Arrange
        CellData cellData = new CellData(0, "12345");

        // Act
        Integer result = cellData.asInt();

        // Assert
        assertEquals(12345, result, "asInt should convert number to Integer correctly");
    }

    @Test
    void asInt_shouldThrowExceptionWhenValueIsOutOfIntRange() {
        // Arrange
        CellData cellData = new CellData(0, "2147483648"); // Integer.MAX_VALUE + 1

        // Act & Assert
        assertThrows(IllegalArgumentException.class, cellData::asInt, 
                "asInt should throw IllegalArgumentException when value is out of int range");
    }

    // asString tests
    @Test
    void asString_shouldReturnFormattedValue() {
        // Arrange
        String value = "test string";
        CellData cellData = new CellData(0, value);

        // Act
        String result = cellData.asString();

        // Assert
        assertEquals(value, result, "asString should return the formatted value as-is");
    }

    // asBoolean tests
    @ParameterizedTest
    @ValueSource(strings = {"true", "TRUE", "True", "1", "y", "Y", "yes", "YES", "Yes"})
    void asBoolean_shouldReturnTrueForTrueValues(String value) {
        // Arrange
        CellData cellData = new CellData(0, value);

        // Act
        boolean result = cellData.asBoolean();

        // Assert
        assertTrue(result, "asBoolean should return true for value: " + value);
    }

    @ParameterizedTest
    @ValueSource(strings = {"false", "FALSE", "False", "0", "n", "N", "no", "NO", "No", "", "   ", "other"})
    void asBoolean_shouldReturnFalseForFalseValues(String value) {
        // Arrange
        CellData cellData = new CellData(0, value);

        // Act
        boolean result = cellData.asBoolean();

        // Assert
        assertFalse(result, "asBoolean should return false for value: " + value);
    }

    @Test
    void asBoolean_shouldReturnFalseWhenValueIsNull() {
        // Arrange
        CellData cellData = new CellData(0, null);

        // Act
        boolean result = cellData.asBoolean();

        // Assert
        assertFalse(result, "asBoolean should return false when value is null");
    }

    // asLocalDateTime tests
    @Test
    void asLocalDateTime_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        LocalDateTime result = cellData.asLocalDateTime();

        // Assert
        assertNull(result, "asLocalDateTime should return null when value is empty");
    }

    @Test
    void asLocalDateTime_shouldParseDefaultFormat() {
        // Arrange
        CellData cellData = new CellData(0, "2025-07-22 14:30:45");

        // Act
        LocalDateTime result = cellData.asLocalDateTime();

        // Assert
        assertEquals(LocalDateTime.of(2025, 7, 22, 14, 30, 45), result, 
                "asLocalDateTime should parse date-time in default format correctly");
    }

    @Test
    void asLocalDateTime_shouldParseCustomFormat() {
        // Arrange
        CellData cellData = new CellData(0, "22/07/2025 14:30");

        // Act
        LocalDateTime result = cellData.asLocalDateTime("dd/MM/yyyy HH:mm");

        // Assert
        assertEquals(LocalDateTime.of(2025, 7, 22, 14, 30), result, 
                "asLocalDateTime should parse date-time in custom format correctly");
    }

    // asLocalDate tests
    @Test
    void asLocalDate_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        LocalDate result = cellData.asLocalDate();

        // Assert
        assertNull(result, "asLocalDate should return null when value is empty");
    }

    @Test
    void asLocalDate_shouldParseDefaultFormat() {
        // Arrange
        CellData cellData = new CellData(0, "2025-07-22");

        // Act
        LocalDate result = cellData.asLocalDate();

        // Assert
        assertEquals(LocalDate.of(2025, 7, 22), result, 
                "asLocalDate should parse date in default format correctly");
    }

    @Test
    void asLocalDate_shouldParseCustomFormat() {
        // Arrange
        CellData cellData = new CellData(0, "22/07/2025");

        // Act
        LocalDate result = cellData.asLocalDate("dd/MM/yyyy");

        // Assert
        assertEquals(LocalDate.of(2025, 7, 22), result, 
                "asLocalDate should parse date in custom format correctly");
    }

    // asLocalTime tests
    @Test
    void asLocalTime_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        LocalTime result = cellData.asLocalTime();

        // Assert
        assertNull(result, "asLocalTime should return null when value is empty");
    }

    @Test
    void asLocalTime_shouldParseDefaultFormat() {
        // Arrange
        CellData cellData = new CellData(0, "14:30:45");

        // Act
        LocalTime result = cellData.asLocalTime();

        // Assert
        assertEquals(LocalTime.of(14, 30, 45), result, 
                "asLocalTime should parse time in default format correctly");
    }

    @Test
    void asLocalTime_shouldParseCustomFormat() {
        // Arrange
        CellData cellData = new CellData(0, "14:30");

        // Act
        LocalTime result = cellData.asLocalTime("HH:mm");

        // Assert
        assertEquals(LocalTime.of(14, 30), result, 
                "asLocalTime should parse time in custom format correctly");
    }

    // asDouble tests
    @Test
    void asDouble_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        Double result = cellData.asDouble();

        // Assert
        assertNull(result, "asDouble should return null when value is empty");
    }

    @Test
    void asDouble_shouldConvertNumberToDouble() {
        // Arrange
        CellData cellData = new CellData(0, "123.456");

        // Act
        Double result = cellData.asDouble();

        // Assert
        assertEquals(123.456, result, 0.001, "asDouble should convert number to Double correctly");
    }

    // asFloat tests
    @Test
    void asFloat_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        Float result = cellData.asFloat();

        // Assert
        assertNull(result, "asFloat should return null when value is empty");
    }

    @Test
    void asFloat_shouldConvertNumberToFloat() {
        // Arrange
        CellData cellData = new CellData(0, "123.456");

        // Act
        Float result = cellData.asFloat();

        // Assert
        assertEquals(123.456f, result, 0.001f, "asFloat should convert number to Float correctly");
    }

    // asBigDecimal tests
    @Test
    void asBigDecimal_shouldReturnNullWhenValueIsEmpty() {
        // Arrange
        CellData cellData = new CellData(0, "");

        // Act
        BigDecimal result = cellData.asBigDecimal();

        // Assert
        assertNull(result, "asBigDecimal should return null when value is empty");
    }

    @Test
    void asBigDecimal_shouldConvertNumberToBigDecimal() {
        // Arrange
        CellData cellData = new CellData(0, "123.456");

        // Act
        BigDecimal result = cellData.asBigDecimal();

        // Assert
        assertEquals(new BigDecimal("123.456"), result, "asBigDecimal should convert number to BigDecimal correctly");
    }
}