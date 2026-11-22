package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellData;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.read.ExcelWorkbookInspector;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.text.NumberFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

/**
 * What a workbook contains does not depend on the JVM's default locale. Several locales write their numbers with
 * their own digits (Arabic-Indic, Persian, Bengali, Devanagari, Thai): a reference such as {@code $A$2:$A$4} or a
 * date such as {@code 2024-01-01} must still come out with ASCII digits, and what is written must read back.
 * Each case switches the default locale for its own duration.
 */
class LocaleIndependenceTest {

    /** Locales whose default numbering system is not the ASCII digits 0-9. */
    private static final List<Locale> NATIVE_DIGITS = List.of(
            Locale.forLanguageTag("ar-EG"), Locale.forLanguageTag("fa-IR"), Locale.forLanguageTag("bn-BD"),
            Locale.forLanguageTag("ne-NP"), Locale.forLanguageTag("th-TH-u-nu-thai"));

    private static final List<Locale> OTHER = List.of(
            Locale.forLanguageTag("tr-TR"), Locale.forLanguageTag("de-DE"), Locale.forLanguageTag("fr-FR"),
            Locale.forLanguageTag("hi-IN"), Locale.forLanguageTag("ko-KR"));

    /** Locales whose calendar is not Gregorian: POI would count a date's year and day in that calendar. */
    private static final List<Locale> OTHER_CALENDARS = List.of(
            Locale.forLanguageTag("th-TH"), Locale.forLanguageTag("ja-JP-u-ca-japanese"),
            Locale.forLanguageTag("th-TH-u-nu-thai"));

    private interface Body {
        void run() throws Exception;
    }

    private static void under(Locale locale, Body body) throws Exception {
        Locale previous = Locale.getDefault();
        Locale.setDefault(locale);
        try {
            body.run();
        } finally {
            Locale.setDefault(previous);
        }
    }

    private static List<Locale> all() {
        List<Locale> all = new ArrayList<>(NATIVE_DIGITS);
        all.addAll(OTHER);
        all.addAll(OTHER_CALENDARS);
        return all;
    }

    @Test
    void theNativeDigitLocalesReallyUseNonAsciiDigits() {
        // If one of them stopped doing so the tests below would pass without testing anything.
        for (Locale locale : NATIVE_DIGITS) {
            assertNotEquals("1", NumberFormat.getIntegerInstance(locale).format(1), locale.toLanguageTag());
        }
    }

    @Test
    void aNamedRangeReferencesItsRowsWithAsciiDigits() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<String>create().column("Name", s -> s).namedRange("Names", 0)
                        .write(List.of("a", "b", "c")).writeTo(out);

                try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    Name name = wb.getName("Names");
                    assertEquals("'Sheet0'!$A$2:$A$4", name.getRefersToFormula(), locale.toLanguageTag());
                }
            });
        }
    }

    @Test
    void datesAndDateTimesReadBackAsWritten() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                LocalDate date = LocalDate.of(2024, 2, 29);
                LocalDateTime dateTime = LocalDateTime.of(2024, 12, 31, 23, 59, 58);
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("Day", x -> date).column("Moment", x -> dateTime)
                        .write(List.of(new Object())).writeTo(out);

                List<List<Object>> rows = new ArrayList<>();
                ExcelReader.mapping(row -> List.<Object>of(row.get("Day").asLocalDate(), row.get("Moment").asLocalDateTime()))
                        .read(new ByteArrayInputStream(out.toByteArray()), r -> {
                            assertEquals(true, r.success(), locale.toLanguageTag() + ": " + r.messages());
                            rows.add(r.data());
                        });

                assertEquals(1, rows.size(), locale.toLanguageTag());
                assertEquals(List.of(date, dateTime), rows.get(0), locale.toLanguageTag());
            });
        }
    }

    @Test
    void aTimeOfDayReadsBackAsWritten() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("At", x -> java.time.LocalTime.of(7, 5, 9)).write(List.of(new Object())).writeTo(out);

                List<java.time.LocalTime> read = new ArrayList<>();
                ExcelReader.mapping(row -> row.get("At").asLocalTime()).read(new ByteArrayInputStream(out.toByteArray()), r -> {
                    assertEquals(true, r.success(), locale.toLanguageTag() + ": " + r.messages());
                    read.add(r.data());
                });

                assertEquals(List.of(java.time.LocalTime.of(7, 5, 9)), read, locale.toLanguageTag());
            });
        }
    }

    @Test
    void aDateCellsDisplayTextHasAsciiDigits() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("Day", x -> LocalDate.of(2024, 1, 5)).write(List.of(new Object())).writeTo(out);

                List<String> shown = new ArrayList<>();
                ExcelReader.mapping(row -> row.get("Day")).read(new ByteArrayInputStream(out.toByteArray()),
                        r -> shown.add(((CellData) r.data()).formattedValue()));

                assertEquals(List.of("2024-01-05"), shown, locale.toLanguageTag());
            });
        }
    }

    @Test
    void numbersReadBackAsWritten() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("I", x -> -123456).column("L", x -> 9_876_543_210L)
                        .column("D", x -> -1234.5).write(List.of(new Object())).writeTo(out);

                List<List<Object>> rows = new ArrayList<>();
                ExcelReader.mapping(row -> List.<Object>of(row.get("I").asInt(), row.get("L").asLong(), row.get("D").asDouble()))
                        .read(new ByteArrayInputStream(out.toByteArray()), r -> {
                            assertEquals(true, r.success(), locale.toLanguageTag() + ": " + r.messages());
                            rows.add(r.data());
                        });

                assertEquals(List.of(List.of(-123456, 9_876_543_210L, -1234.5)), rows, locale.toLanguageTag());
            });
        }
    }

    @Test
    void aDateHeaderIsDiscoveredWithAsciiDigits() throws Exception {
        for (Locale locale : all()) {
            under(locale, () -> {
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("D", x -> "v").write(List.of(new Object())).writeTo(out);
                // headers are text; a date written below them is shown with the same digits the cells use
                ByteArrayOutputStream dated = new ByteArrayOutputStream();
                ExcelWriter.<Object>create().column("Day", x -> LocalDate.of(2025, 3, 9)).write(List.of(new Object())).writeTo(dated);

                List<String> row = ExcelWorkbookInspector.create().headers(new ByteArrayInputStream(dated.toByteArray()), 0, 1);
                assertEquals(List.of("2025-03-09"), row, locale.toLanguageTag());
            });
        }
    }
}
