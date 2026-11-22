package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.UUID;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for CellData custom conversion and default value methods.
 */
class CellDataConversionTest {

    @Test
    void floatConversionRoundsDirectlyFromTheLocalizedDecimal() {
        var us = CellConversionConfig.builder().locale(Locale.US).build();
        var german = CellConversionConfig.builder().locale(Locale.GERMANY).build();
        assertEquals(Math.nextUp(1.0f), new CellData(0, "1.0000000596046448", us).asFloat());
        assertEquals(-Math.nextUp(1.0f), new CellData(0, "-1,0000000596046448", german).asFloat());
        assertEquals(1.0f, new CellData(0, "1.000000059604644775390625", us).asFloat());
        assertEquals(-0.0f, new CellData(0, "-0", us).asFloat());
    }

    @Test
    void numericConversionRejectsEmbeddedAffixes() {
        var us = CellConversionConfig.builder().locale(Locale.US).build();
        for (String value : List.of("1$2", "12%34", "$12%", "$$12")) {
            var cell = new CellData(0, value, us);
            assertThrows(IllegalArgumentException.class, cell::asNumber, value);
            assertThrows(IllegalArgumentException.class, cell::asBigDecimal, value);
        }
        assertEquals(-1234.56, new CellData(0, "-$1,234.56", us).asDouble());
        assertEquals(new BigDecimal("25"), new CellData(0, "25 %", us).asBigDecimal());
    }

    @Test
    void floatingPointConversionsRejectNonFiniteAndUnderflowedValues() {
        var us = CellConversionConfig.builder().locale(Locale.US).build();
        for (String value : List.of("NaN", "∞", "1E309", "1E-400")) {
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, value, us).asDouble(), value);
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, value, us).asNumber(), value);
        }
        for (String value : List.of("NaN", "3.5E38", "1E-50")) {
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, value, us).asFloat(), value);
        }
        assertEquals(Double.MIN_VALUE, new CellData(0, "4.9E-324", us).asDouble());
        assertEquals(-0.0, new CellData(0, "-0", us).asDouble());
        assertEquals(0.0f, new CellData(0, "0", us).asFloat());
    }

    @Test
    void hourOfDayPatternRejectsTwentyFourRatherThanNormalizingIt() {
        assertEquals(LocalTime.MIDNIGHT, cell("00:00").asLocalTime("HH:mm"));
        assertThrows(DateTimeParseException.class, () -> cell("24:00").asLocalTime("HH:mm"));
    }

    @Test
    void longConversionRejectsFractionsNonFiniteValuesAndOverflow() {
        var us = CellConversionConfig.builder().locale(Locale.US).build();
        assertEquals(Long.MAX_VALUE, new CellData(0, "9,223,372,036,854,775,807", us).asLong());
        assertEquals(Long.MIN_VALUE, new CellData(0, "-9223372036854775808.0", us).asLong());
        for (String value : List.of("9223372036854775808", "-9223372036854775809", "NaN", "∞", "1.5")) {
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, value, us).asLong(), value);
        }
    }

    @Test
    void numericConversionValidatesGroupingBeforeRemovingSeparators() {
        var us = CellConversionConfig.builder().locale(Locale.US).build();
        var french = CellConversionConfig.builder().locale(Locale.FRANCE).build();
        for (String value : List.of("1,2", "1,234.5,67", "1,,234", "1234,567")) {
            var cell = new CellData(0, value, us);
            assertThrows(IllegalArgumentException.class, cell::asNumber, value);
            assertThrows(IllegalArgumentException.class, cell::asBigDecimal, value);
        }
        assertEquals(new BigDecimal("1234.56"), new CellData(0, "1\u202F234,56", french).asBigDecimal());
    }

    @Test
    void dateConversionRejectsInvalidCalendarDates() {
        assertEquals(LocalDate.of(2024, 2, 29), cell("02/29/24").asLocalDate());
        assertThrows(DateTimeParseException.class, () -> cell("2025-02-30").asLocalDate());
        assertThrows(DateTimeParseException.class, () -> cell("2025-02-30 12:00").asLocalDateTime());
        assertThrows(DateTimeParseException.class, () -> cell("30.02.2025").asLocalDate("dd.MM.yyyy"));
    }

    private CellData cell(String value) {
        return new CellData(0, value);
    }

    private CellData blank() {
        return new CellData(0, "");
    }

    private CellData nullCell() {
        // CellData compact constructor normalizes null to ""
        return new CellData(0, null);
    }

    // ============================================================
    // as(Function) — custom conversion
    // ============================================================
    @Nested
    class AsFunction_CustomConversion {

        @Test
        void as_convertsWithFunction() {
            var uuid = UUID.randomUUID();
            var result = cell(uuid.toString()).as(UUID::fromString);
            assertEquals(uuid, result);
        }

        @Test
        void as_returnsNullForBlank() {
            assertNull(blank().as(UUID::fromString));
        }

        @Test
        void as_customParseLogic() {
            var result = cell("3,4").as(s -> {
                String[] parts = s.split(",");
                return new int[]{Integer.parseInt(parts[0]), Integer.parseInt(parts[1])};
            });
            assertNotNull(result);
            assertArrayEquals(new int[]{3, 4}, result);
        }

        @Test
        void as_converterThatThrowsException_shouldPropagate() {
            var cellData = cell("not-a-uuid");
            assertThrows(IllegalArgumentException.class, () -> cellData.as(UUID::fromString));
        }

        @Test
        void as_whitespaceOnlyValue_treatedAsBlank() {
            assertNull(cell("   ").as(UUID::fromString));
            assertNull(cell("\t").as(UUID::fromString));
            assertNull(cell(" \n ").as(UUID::fromString));
        }

        @Test
        void as_nullFormattedValue_treatedAsBlank() {
            // CellData normalizes null to "", so this behaves as blank
            assertNull(nullCell().as(UUID::fromString));
        }

        @Test
        void as_converterReturnsNull_shouldReturnNull() {
            var result = cell("something").as(s -> null);
            assertNull(result);
        }
    }

    // ============================================================
    // as(Function, R defaultValue) — custom with default
    // ============================================================
    @Nested
    class AsFunctionWithDefault {

        @Test
        void as_withDefault_returnsConvertedValueWhenPresent() {
            var defaultUuid = UUID.randomUUID();
            var uuid = UUID.randomUUID();
            var result = cell(uuid.toString()).as(UUID::fromString, defaultUuid);
            assertEquals(uuid, result);
        }

        @Test
        void as_withDefault_returnsDefaultForBlank() {
            var defaultUuid = UUID.randomUUID();
            var result = blank().as(UUID::fromString, defaultUuid);
            assertEquals(defaultUuid, result);
        }

        @Test
        void as_withDefault_returnsDefaultForWhitespaceOnly() {
            var defaultUuid = UUID.randomUUID();
            var result = cell("   ").as(UUID::fromString, defaultUuid);
            assertEquals(defaultUuid, result);
        }

        @Test
        void as_withDefault_returnsDefaultForNullFormattedValue() {
            var defaultUuid = UUID.randomUUID();
            var result = nullCell().as(UUID::fromString, defaultUuid);
            assertEquals(defaultUuid, result);
        }

        @Test
        void as_withDefault_converterAppliedToNonBlank() {
            var result = cell("42").as(Integer::parseInt, -1);
            assertEquals(42, result);
        }
    }

    // ============================================================
    // asInt(int defaultValue)
    // ============================================================
    @Nested
    class AsIntWithDefault {

        @Test
        void asInt_returnsValueWhenPresent() {
            assertEquals(42, cell("42").asInt(0));
        }

        @Test
        void asInt_returnsDefaultForBlank() {
            assertEquals(-1, blank().asInt(-1));
        }

        @Test
        void asInt_returnsDefaultForNull() {
            assertEquals(99, nullCell().asInt(99));
        }

        @Test
        void asInt_worksWithFormattedNumbers() {
            // asNumber() strips commas before parsing
            assertEquals(1234, cell("1,234").asInt(0));
        }

        @Test
        void asInt_zeroDefault() {
            assertEquals(0, blank().asInt(0));
        }

        @Test
        void asInt_negativeValue() {
            assertEquals(-5, cell("-5").asInt(0));
        }

        @Test
        void asInt_whitespaceOnlyReturnsDefault() {
            assertEquals(7, cell("   ").asInt(7));
        }
    }

    // ============================================================
    // asLong(long defaultValue)
    // ============================================================
    @Nested
    class AsLongWithDefault {

        @Test
        void asLong_returnsValueWhenPresent() {
            assertEquals(100L, cell("100").asLong(0L));
        }

        @Test
        void asLong_returnsDefaultForBlank() {
            assertEquals(999L, blank().asLong(999L));
        }

        @Test
        void asLong_returnsDefaultForNull() {
            assertEquals(42L, nullCell().asLong(42L));
        }

        @Test
        void asLong_worksWithFormattedNumbers() {
            assertEquals(1234567L, cell("1,234,567").asLong(0L));
        }

        @Test
        void numericConversionsAcceptScientificNotationWithLocaleSpecificDecimals() {
            var us = CellConversionConfig.builder().locale(Locale.US).build();
            var german = CellConversionConfig.builder().locale(Locale.GERMANY).build();

            assertEquals(1_234_570_000_000L, new CellData(0, "1.23457E+12", us).asLong());
            assertEquals(1_234_570_000_000L, new CellData(0, "1,23457e+12", german).asLong());
            assertEquals(0.0125, new CellData(0, "1.25E-2", us).asDouble());
        }

        @Test
        void numericConversionsRejectPartiallyParsedValues() {
            var us = CellConversionConfig.builder().locale(Locale.US).build();

            assertThrows(IllegalArgumentException.class, () -> new CellData(0, "123abc", us).asLong());
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, "1.25E+", us).asDouble());
            assertThrows(IllegalArgumentException.class, () -> new CellData(0, "1.25E+-2", us).asLong());
        }

        @Test
        void asLong_zeroDefault() {
            assertEquals(0L, blank().asLong(0L));
        }

        @Test
        void asLong_negativeValue() {
            assertEquals(-100L, cell("-100").asLong(0L));
        }

        @Test
        void asLong_whitespaceOnlyReturnsDefault() {
            assertEquals(5L, cell("  ").asLong(5L));
        }
    }

    // ============================================================
    // asDouble(double defaultValue)
    // ============================================================
    @Nested
    class AsDoubleWithDefault {

        @Test
        void asDouble_returnsValueWhenPresent() {
            assertEquals(3.14, cell("3.14").asDouble(0.0), 0.001);
        }

        @Test
        void asDouble_returnsDefaultForBlank() {
            assertEquals(-1.0, blank().asDouble(-1.0), 0.001);
        }

        @Test
        void asDouble_returnsDefaultForNull() {
            assertEquals(2.5, nullCell().asDouble(2.5), 0.001);
        }

        @Test
        void asDouble_worksWithDecimalValues() {
            assertEquals(0.001, cell("0.001").asDouble(0.0), 0.0001);
        }

        @Test
        void asDouble_zeroDefault() {
            assertEquals(0.0, blank().asDouble(0.0), 0.001);
        }

        @Test
        void asDouble_negativeValue() {
            assertEquals(-9.99, cell("-9.99").asDouble(0.0), 0.001);
        }

        @Test
        void asDouble_whitespaceOnlyReturnsDefault() {
            assertEquals(1.1, cell("  \t ").asDouble(1.1), 0.001);
        }
    }

    // ============================================================
    // asString(String defaultValue)
    // ============================================================
    @Nested
    class AsStringWithDefault {

        @Test
        void asString_returnsValueWhenPresent() {
            assertEquals("hello", cell("hello").asString("default"));
        }

        @Test
        void asString_returnsDefaultForBlank() {
            assertEquals("N/A", blank().asString("N/A"));
        }

        @Test
        void asString_returnsDefaultForWhitespaceOnly() {
            assertEquals("N/A", cell("   ").asString("N/A"));
        }

        @Test
        void asString_defaultIsEmptyString() {
            assertEquals("", blank().asString(""));
        }

        @Test
        void asString_nonBlankWhitespacePaddedValue_returnsActualValue() {
            // " hello " is not blank, so should be returned as-is (not the default)
            assertEquals(" hello ", cell(" hello ").asString("default"));
        }

        @Test
        void asString_returnsDefaultForNull() {
            assertEquals("fallback", nullCell().asString("fallback"));
        }

        @Test
        void asString_tabOnlyIsBlank() {
            assertEquals("default", cell("\t").asString("default"));
        }
    }

    // ============================================================
    // Integration: round-trip with ExcelWriter + ExcelReader mapping mode
    // ============================================================
    @Nested
    class IntegrationRoundTrip {

        record Product(String name, int quantity, double price) {}

        @Test
        void roundTrip_writeAndReadWithMappingModeUsingAsAndAsInt() throws IOException {
            // Write Excel
            var baos = new ByteArrayOutputStream();
            ExcelWriter.<Product>create()
                    .column("Name", Product::name)
                    .column("Quantity", p -> p.quantity())
                    .column("Price", p -> p.price())
                    .write(Stream.of(
                            new Product("Widget", 10, 19.99),
                            new Product("Gadget", 5, 49.50),
                            new Product("", 0, 0.0)
                    ))
                    .writeTo(baos);

            // Read back with mapping mode using as(), asInt(), asDouble()
            List<Product> products = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
                ExcelReader.<Product>mapping(row -> new Product(
                        row.get("Name").asString("unknown"),
                        row.get("Quantity").asInt(0),
                        row.get("Price").asDouble(0.0)
                )).read(is, r -> {
                    assertTrue(r.success());
                    products.add(r.data());
                });
            }

            assertEquals(3, products.size());

            assertEquals("Widget", products.get(0).name());
            assertEquals(10, products.get(0).quantity());
            assertEquals(19.99, products.get(0).price(), 0.01);

            assertEquals("Gadget", products.get(1).name());
            assertEquals(5, products.get(1).quantity());
            assertEquals(49.50, products.get(1).price(), 0.01);

            // Third row: empty name should use default, 0 values should parse
            assertEquals(0, products.get(2).quantity());
            assertEquals(0.0, products.get(2).price(), 0.01);
        }

        @Test
        void roundTrip_customAsConversion() throws IOException {
            var uuid1 = UUID.randomUUID();
            var uuid2 = UUID.randomUUID();

            // Write Excel with UUID strings
            var baos = new ByteArrayOutputStream();
            ExcelWriter.<UUID>create()
                    .column("ID", UUID::toString)
                    .write(Stream.of(uuid1, uuid2))
                    .writeTo(baos);

            // Read back with as(UUID::fromString)
            List<UUID> ids = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
                ExcelReader.<UUID>mapping(row ->
                        row.get("ID").as(UUID::fromString)
                ).read(is, r -> {
                    assertTrue(r.success());
                    ids.add(r.data());
                });
            }

            assertEquals(2, ids.size());
            assertEquals(uuid1, ids.get(0));
            assertEquals(uuid2, ids.get(1));
        }
    }
}
