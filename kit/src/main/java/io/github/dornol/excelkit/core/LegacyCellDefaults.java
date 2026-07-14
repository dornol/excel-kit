package io.github.dornol.excelkit.core;

import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.CopyOnWriteArrayList;

/** Isolates mutable legacy conversion defaults from the CellData value object. */
final class LegacyCellDefaults {
    private LegacyCellDefaults() {
    }

    private static final List<DateTimeFormatter> DATES = List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd"), DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("MM/dd/yy"), DateTimeFormatter.ofPattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE);
    private static final List<DateTimeFormatter> DATE_TIMES = List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("MM/dd/yy[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("M/d/yy[ HH:mm[:ss]]"), DateTimeFormatter.ISO_LOCAL_DATE_TIME);
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
        dates.add(0, DateTimeFormatter.ofPattern(pattern));
    }

    static void addDateTime(String pattern) {
        dateTimes.add(0, DateTimeFormatter.ofPattern(pattern));
    }

    static void resetDates() {
        dates = new CopyOnWriteArrayList<>(DATES);
    }

    static void resetDateTimes() {
        dateTimes = new CopyOnWriteArrayList<>(DATE_TIMES);
    }
}
