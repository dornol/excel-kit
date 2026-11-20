package io.github.dornol.excelkit.core;

import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.CopyOnWriteArrayList;

import static io.github.dornol.excelkit.core.CellConversionSupport.strictDatePattern;

/** Isolates mutable legacy conversion defaults from the CellData value object. */
final class LegacyCellDefaults {
    private LegacyCellDefaults() {
    }

    private static final List<DateTimeFormatter> DATES = List.of(
            strictDatePattern("yyyy-MM-dd"), strictDatePattern("yyyy/MM/dd"),
            strictDatePattern("MM/dd/yy"), strictDatePattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE);
    private static final List<DateTimeFormatter> DATE_TIMES = List.of(
            strictDatePattern("yyyy-MM-dd[ HH:mm[:ss]]"),
            strictDatePattern("yyyy/MM/dd[ HH:mm[:ss]]"),
            strictDatePattern("MM/dd/yy[ HH:mm[:ss]]"),
            strictDatePattern("M/d/yy[ HH:mm[:ss]]"), DateTimeFormatter.ISO_LOCAL_DATE_TIME);
    private static volatile Locale locale = Locale.getDefault();
    private static volatile CopyOnWriteArrayList<DateTimeFormatter> dates = new CopyOnWriteArrayList<>(DATES);
    private static volatile CopyOnWriteArrayList<DateTimeFormatter> dateTimes = new CopyOnWriteArrayList<>(DATE_TIMES);

    static Locale locale() {
        return locale;
    }

    static void locale(Locale value) {
        locale = java.util.Objects.requireNonNull(value);
    }

    static List<DateTimeFormatter> dates() {
        return List.copyOf(dates);
    }

    static List<DateTimeFormatter> dateTimes() {
        return List.copyOf(dateTimes);
    }

    static void addDate(String pattern) {
        dates.add(0, strictDatePattern(pattern));
    }

    static void addDateTime(String pattern) {
        dateTimes.add(0, strictDatePattern(pattern));
    }

    static void resetDates() {
        dates = new CopyOnWriteArrayList<>(DATES);
    }

    static void resetDateTimes() {
        dateTimes = new CopyOnWriteArrayList<>(DATE_TIMES);
    }
}
