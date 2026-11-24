package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * A date written to a cell is the wall-clock time the program holds, whatever the JVM's default time zone is: the
 * clock a person reads in Excel does not move with the server's zone, nor around a daylight-saving change.
 * Only {@code java.util.Date} (an instant) is shown in the default zone, as documented.
 */
class DateWriteTimezoneTest {
    record Stamp(LocalDateTime dateTime, LocalDate date, LocalTime time, java.time.ZonedDateTime zoned, Date legacy) {}

    private static final String[] ZONES = {"UTC", "Asia/Seoul", "America/New_York", "Europe/London", "Europe/Berlin",
            "Australia/Lord_Howe", "Pacific/Apia", "America/Sao_Paulo", "Asia/Kolkata", "Pacific/Kiritimati"};

    /** Wall-clock times in daylight-saving gaps and overlaps of several zones, the 1900 leap-year bug and the far end. */
    private static final List<LocalDateTime> TIMES = List.of(
            LocalDateTime.of(2026, 3, 8, 2, 30), LocalDateTime.of(2026, 11, 1, 1, 30),
            LocalDateTime.of(2026, 3, 29, 1, 30), LocalDateTime.of(2026, 10, 25, 2, 30),
            LocalDateTime.of(2026, 4, 5, 2, 15), LocalDateTime.of(2026, 10, 4, 2, 15),
            LocalDateTime.of(1900, 2, 28, 12, 0), LocalDateTime.of(1900, 3, 1, 0, 0),
            LocalDateTime.of(2026, 12, 31, 23, 59, 59, 999_000_000), LocalDateTime.of(9999, 12, 31, 0, 0));

    private static byte[] write(Stamp stamp) {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Stamp>create().column("dateTime", Stamp::dateTime).column("date", Stamp::date)
                .column("time", Stamp::time).column("zoned", Stamp::zoned).column("legacy", Stamp::legacy)
                .write(List.of(stamp)).writeTo(out);
        return out.toByteArray();
    }

    @Test
    void theClockInTheCellDoesNotDependOnTheDefaultZone() throws IOException {
        TimeZone original = TimeZone.getDefault();
        try {
            for (String zone : ZONES) {
                TimeZone.setDefault(TimeZone.getTimeZone(zone));
                for (LocalDateTime time : TIMES) {
                    Date legacy = Date.from(time.atZone(ZoneId.systemDefault()).toInstant());
                    var stamp = new Stamp(time, time.toLocalDate(), time.toLocalTime().withNano(0),
                            time.atZone(ZoneId.of("Asia/Tokyo")), legacy);
                    String at = zone + " " + time;
                    try (var wb = new XSSFWorkbook(new ByteArrayInputStream(write(stamp)))) {
                        var row = wb.getSheetAt(0).getRow(1);
                        assertEquals(DateUtil.getExcelDate(time, false), row.getCell(0).getNumericCellValue(), 1e-9, at + " LocalDateTime");
                        assertEquals(DateUtil.getExcelDate(time.toLocalDate().atStartOfDay(), false),
                                row.getCell(1).getNumericCellValue(), 1e-9, at + " LocalDate");
                        assertEquals(time.toLocalTime().withNano(0).toNanoOfDay() / 86_400e9,
                                row.getCell(2).getNumericCellValue(), 1e-9, at + " LocalTime");
                        assertEquals(DateUtil.getExcelDate(time, false), row.getCell(3).getNumericCellValue(), 1e-9,
                                at + " ZonedDateTime keeps its own wall clock, not the default zone's");
                        // an instant is shown in the default zone: what a person sees is the local time of that instant
                        assertEquals(legacy.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime(),
                                row.getCell(4).getLocalDateTimeCellValue(), at + " Date");
                    }
                }
            }
        } finally {
            TimeZone.setDefault(original);
        }
    }
}
