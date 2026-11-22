package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.testsupport.DefaultLocaleShowsEnglishNumbers;
import org.junit.jupiter.api.Test;

import java.time.DateTimeException;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.OptionalDouble;
import java.util.OptionalLong;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The corners of {@link CellData}: optional accessors, strict time zones, equality. */
class CellDataDetailsTest {

    private static CellData cell(String text) {
        return new CellData(0, text);
    }

    @Test
    @DefaultLocaleShowsEnglishNumbers
    void optionalAccessorsAreEmptyForBlankCellsAndFullOtherwise() {
        assertEquals(OptionalLong.of(42), cell("42").asOptionalLong());
        assertEquals(OptionalLong.empty(), cell("").asOptionalLong());
        assertEquals(OptionalLong.empty(), cell("  ").asOptionalLong());
        assertEquals(OptionalDouble.of(2.5), cell("2.5").asOptionalDouble());
        assertEquals(OptionalDouble.empty(), cell("").asOptionalDouble());
        assertEquals(java.util.OptionalInt.of(7), cell("7").asOptionalInt());
        assertEquals(java.util.OptionalInt.empty(), cell(null).asOptionalInt());
    }

    @Test
    void strictZonedConversionRefusesTimesThatDoNotExistOrRepeat() {
        ZoneId newYork = ZoneId.of("America/New_York");
        // 2024-03-10 02:30 does not exist in New York (clocks jump from 02:00 to 03:00)
        DateTimeException gap = assertThrows(DateTimeException.class,
                () -> cell("2024-03-10 02:30:00").asZonedDateTimeStrict(newYork));
        assertTrue(gap.getMessage().contains("does not exist"), gap.getMessage());
        // 2024-11-03 01:30 happens twice
        DateTimeException overlap = assertThrows(DateTimeException.class,
                () -> cell("2024-11-03 01:30:00").asZonedDateTimeStrict(newYork));
        assertTrue(overlap.getMessage().contains("is ambiguous"), overlap.getMessage());
        // an ordinary time is fine, and a blank cell has none
        assertEquals(ZonedDateTime.of(2024, 6, 1, 12, 0, 0, 0, newYork),
                cell("2024-06-01 12:00:00").asZonedDateTimeStrict(newYork));
        assertEquals(null, cell("").asZonedDateTimeStrict(newYork));
    }

    @Test
    void strictZonedConversionWithAPatternChecksTheSameWay() {
        ZoneId newYork = ZoneId.of("America/New_York");
        assertEquals(ZonedDateTime.of(2024, 6, 1, 12, 0, 0, 0, newYork),
                cell("01/06/2024 12:00").asZonedDateTimeStrict("dd/MM/yyyy HH:mm", newYork));
        assertThrows(DateTimeException.class,
                () -> cell("10/03/2024 02:30").asZonedDateTimeStrict("dd/MM/yyyy HH:mm", newYork));
        assertEquals(null, cell(" ").asZonedDateTimeStrict("dd/MM/yyyy HH:mm", newYork));
    }

    @Test
    void lenientZonedConversionMovesAcrossAGap() {
        ZonedDateTime moved = cell("2024-03-10 02:30:00").asZonedDateTime(ZoneId.of("America/New_York"));
        assertEquals(LocalDateTime.of(2024, 3, 10, 3, 30), moved.toLocalDateTime());
    }

    @Test
    void cellsAreEqualByColumnTextAndNumericValue() {
        assertEquals(cell("a"), cell("a"));
        assertEquals(cell("a").hashCode(), cell("a").hashCode());
        assertNotEquals(cell("a"), cell("b"));
        assertNotEquals(cell("a"), new CellData(1, "a"));
        assertNotEquals(cell("a"), "a");
        assertNotEquals(cell("a"), null);

        CellConversionConfig defaults = CellConversionConfig.defaults();
        // the same number written differently is the same cell: 1.50 and 1.5
        assertEquals(new CellData(0, "1.5", defaults, "1.50"), new CellData(0, "1.5", defaults, "1.5"));
        assertEquals(new CellData(0, "1.5", defaults, "1.50").hashCode(), new CellData(0, "1.5", defaults, "1.5").hashCode());
        assertNotEquals(new CellData(0, "1.5", defaults, "1.5"), new CellData(0, "1.5", defaults, "1.6"));
        assertNotEquals(new CellData(0, "1.5", defaults, "1.5"), new CellData(0, "1.5"));
        // equality stays defined for a numeric cell that is not a number
        assertEquals(new CellData(0, "x", defaults, "not-a-number"), new CellData(0, "x", defaults, "not-a-number"));
        assertNotEquals(new CellData(0, "x", defaults, "not-a-number"), new CellData(0, "x", defaults, "other"));
        assertFalse(cell("a").equals(null));
    }

    @Test
    void toStringNamesTheColumnAndTheText() {
        assertEquals("CellData[columnIndex=3, formattedValue=hello]", new CellData(3, "hello").toString());
    }
}
