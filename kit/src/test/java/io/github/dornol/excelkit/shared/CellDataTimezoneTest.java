package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CellData#asZonedDateTime(ZoneId)} and related timezone methods.
 */
class CellDataTimezoneTest {

    @Test
    void asZonedDateTime_withZoneId_shouldAttachTimezone() {
        CellData cell = new CellData(0, "2025-06-15 14:30:00");
        ZonedDateTime result = cell.asZonedDateTime(ZoneId.of("Asia/Seoul"));

        assertNotNull(result);
        assertEquals(2025, result.getYear());
        assertEquals(6, result.getMonthValue());
        assertEquals(15, result.getDayOfMonth());
        assertEquals(14, result.getHour());
        assertEquals(30, result.getMinute());
        assertEquals(ZoneId.of("Asia/Seoul"), result.getZone());
    }

    @Test
    void asZonedDateTime_utc() {
        CellData cell = new CellData(0, "2025-01-01 00:00:00");
        ZonedDateTime result = cell.asZonedDateTime(ZoneId.of("UTC"));

        assertNotNull(result);
        assertEquals(ZoneId.of("UTC"), result.getZone());
    }

    @Test
    void asZonedDateTime_blank_returnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asZonedDateTime(ZoneId.of("Asia/Seoul")));
    }

    @Test
    void asZonedDateTime_withFormat_shouldWork() {
        CellData cell = new CellData(0, "15/06/2025 14:30");
        ZonedDateTime result = cell.asZonedDateTime("dd/MM/yyyy HH:mm", ZoneId.of("Europe/London"));

        assertNotNull(result);
        assertEquals(15, result.getDayOfMonth());
        assertEquals(6, result.getMonthValue());
        assertEquals(14, result.getHour());
        assertEquals(ZoneId.of("Europe/London"), result.getZone());
    }

    @Test
    void asZonedDateTime_withFormat_blank_returnsNull() {
        CellData cell = new CellData(0, "  ");
        assertNull(cell.asZonedDateTime("yyyy-MM-dd HH:mm", ZoneId.of("UTC")));
    }

    @Test
    void asZonedDateTime_dateWithTime_withDefaultFormats() {
        CellData cell = new CellData(0, "2025-06-15 00:00");
        ZonedDateTime result = cell.asZonedDateTime(ZoneId.of("America/New_York"));

        assertNotNull(result);
        assertEquals(2025, result.getYear());
        assertEquals(6, result.getMonthValue());
        assertEquals(15, result.getDayOfMonth());
        assertEquals(0, result.getHour());
    }

    @Test
    void asZonedDateTime_differentTimezones_differentInstant() {
        CellData cell = new CellData(0, "2025-06-15 12:00:00");

        ZonedDateTime seoul = cell.asZonedDateTime(ZoneId.of("Asia/Seoul"));
        ZonedDateTime utc = cell.asZonedDateTime(ZoneId.of("UTC"));

        assertNotNull(seoul);
        assertNotNull(utc);
        // Same local time, different instants
        assertEquals(seoul.toLocalDateTime(), utc.toLocalDateTime());
        assertNotEquals(seoul.toInstant(), utc.toInstant());
    }
}
