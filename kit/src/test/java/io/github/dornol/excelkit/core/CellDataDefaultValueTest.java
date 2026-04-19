package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for CellData default-value overloads and custom conversion edge cases.
 */
class CellDataDefaultValueTest {

    @Nested
    class DefaultOverloads {

        @Test
        void asInt_blankCell_returnsDefault() {
            assertEquals(42, new CellData(0, "").asInt(42));
            assertEquals(0, new CellData(0, "  ").asInt(0));
        }

        @Test
        void asInt_nonBlankCell_parsesValue() {
            assertEquals(7, new CellData(0, "7").asInt(42));
        }

        @Test
        void asLong_blankCell_returnsDefault() {
            assertEquals(100L, new CellData(0, "").asLong(100L));
        }

        @Test
        void asLong_nonBlankCell_parsesValue() {
            assertEquals(999L, new CellData(0, "999").asLong(0L));
        }

        @Test
        void asDouble_blankCell_returnsDefault() {
            assertEquals(3.14, new CellData(0, "").asDouble(3.14));
        }

        @Test
        void asDouble_nonBlankCell_parsesValue() {
            assertEquals(2.5, new CellData(0, "2.5").asDouble(0.0), 0.001);
        }

        @Test
        void asString_blankCell_returnsDefault() {
            assertEquals("N/A", new CellData(0, "").asString("N/A"));
        }

        @Test
        void asString_nonBlankCell_returnsValue() {
            assertEquals("hello", new CellData(0, "hello").asString("N/A"));
        }

        @Test
        void asString_nullValue_returnsDefault() {
            assertEquals("default", new CellData(0, null).asString("default"));
        }
    }

    @Nested
    class CustomConversion {

        @Test
        void as_converter_blankCell_returnsNull() {
            assertNull(new CellData(0, "").as(UUID::fromString));
        }

        @Test
        void as_converter_validValue_parses() {
            UUID expected = UUID.fromString("550e8400-e29b-41d4-a716-446655440000");
            assertEquals(expected, new CellData(0, "550e8400-e29b-41d4-a716-446655440000").as(UUID::fromString));
        }

        @Test
        void as_converterWithDefault_blankCell_returnsDefault() {
            UUID fallback = UUID.randomUUID();
            assertEquals(fallback, new CellData(0, "").as(UUID::fromString, fallback));
        }

        @Test
        void as_converterWithDefault_validValue_parses() {
            UUID fallback = UUID.randomUUID();
            UUID expected = UUID.fromString("550e8400-e29b-41d4-a716-446655440000");
            assertEquals(expected, new CellData(0, "550e8400-e29b-41d4-a716-446655440000").as(UUID::fromString, fallback));
        }
    }

    @Nested
    class IntRangeEdgeCases {

        @Test
        void asInt_maxValue_succeeds() {
            assertEquals(Integer.MAX_VALUE,
                    new CellData(0, String.valueOf(Integer.MAX_VALUE)).asInt());
        }

        @Test
        void asInt_minValue_succeeds() {
            assertEquals(Integer.MIN_VALUE,
                    new CellData(0, String.valueOf(Integer.MIN_VALUE)).asInt());
        }

        @Test
        void asInt_overMaxValue_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> new CellData(0, String.valueOf((long) Integer.MAX_VALUE + 1)).asInt());
        }

        @Test
        void asInt_underMinValue_throws() {
            assertThrows(IllegalArgumentException.class,
                    () -> new CellData(0, String.valueOf((long) Integer.MIN_VALUE - 1)).asInt());
        }
    }
}
