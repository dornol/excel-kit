package io.github.dornol.excelkit.shared;

import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for CellData custom conversion and default value methods.
 */
class CellDataConversionTest {

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
            new ExcelWriter<Product>()
                    .column("Name", Product::name)
                    .column("Quantity", p -> p.quantity())
                    .column("Price", p -> p.price())
                    .write(Stream.of(
                            new Product("Widget", 10, 19.99),
                            new Product("Gadget", 5, 49.50),
                            new Product("", 0, 0.0)
                    ))
                    .consumeOutputStream(baos);

            // Read back with mapping mode using as(), asInt(), asDouble()
            List<Product> products = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
                ExcelReader.<Product>mapping(row -> new Product(
                        row.get("Name").asString("unknown"),
                        row.get("Quantity").asInt(0),
                        row.get("Price").asDouble(0.0)
                )).build(is).read(r -> {
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
            new ExcelWriter<UUID>()
                    .column("ID", UUID::toString)
                    .write(Stream.of(uuid1, uuid2))
                    .consumeOutputStream(baos);

            // Read back with as(UUID::fromString)
            List<UUID> ids = new ArrayList<>();
            try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
                ExcelReader.<UUID>mapping(row ->
                        row.get("ID").as(UUID::fromString)
                ).build(is).read(r -> {
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
