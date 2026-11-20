package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowData;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.TimeZone;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelValuePreservationTest {

    @Test
    void millisecondDateTimeAndTimeRoundTripWithoutRoundingToTheNextSecond() {
        var ordinary = LocalDateTime.of(2025, 4, 28, 15, 20, 30);
        var fractional = ordinary.withNano(987_000_000);
        var lateTime = LocalTime.of(23, 59, 59, 999_000_000);
        var output = new ByteArrayOutputStream();
        ExcelWriter.<LocalDateTime>create().column("Default", value -> value)
                .column("Custom", value -> value, c -> c.format("yyyy-mm-dd hh:mm:ss.000"))
                .column("Time", value -> value.getNano() == 0 ? LocalTime.NOON : lateTime)
                .write(List.of(ordinary, fractional)).writeTo(output);

        var rows = ExcelReader.mapping(row -> row).readAll(new ByteArrayInputStream(output.toByteArray()));
        assertEquals("2025-04-28 15:20:30", rows.get(0).get("Default").formattedValue());
        assertEquals(fractional, rows.get(1).get("Default").asLocalDateTime());
        assertEquals(fractional, rows.get(1).get("Custom").asLocalDateTime());
        assertEquals("23:59:59.999", rows.get(1).get("Time").formattedValue());
        assertEquals(lateTime, rows.get(1).get("Time").asLocalTime());
    }

    @Test
    void readResultsAndErrorsIncludeStoredNumericTokensInsteadOfDisplayValues() {
        var output = new ByteArrayOutputStream();
        ExcelWriter.<Long>create().column("Id", value -> value, c -> c.format("General"))
                .write(List.of(1234567890123L, 1234567890124L)).writeTo(output);
        byte[] workbook = output.toByteArray();
        var rows = new ArrayList<ReadResult<CellData>>();
        ExcelReader.mapping(row -> row.get("Id")).read(new ByteArrayInputStream(workbook), rows::add);

        assertEquals("1.23457E+12", rows.get(0).data().formattedValue());
        assertEquals("1.23457E+12", rows.get(1).data().formattedValue());
        assertEquals(1234567890123L, new BigDecimal(rows.get(0).rawValues().get(0)).longValueExact());
        assertEquals(1234567890124L, new BigDecimal(rows.get(1).rawValues().get(0)).longValueExact());

        var report = ExcelReader.mapping(row -> row.get("Id").asInt())
                .readReport(new ByteArrayInputStream(workbook), 2);
        assertEquals(rows.get(0).rawValues(), report.errors().get(0).rawValues());
        assertEquals(rows.get(1).rawValues(), report.errors().get(1).rawValues());
    }

    @Test
    void formatsExcelCalendarValuesIndependentlyOfThePoiTimeZone() throws IOException {
        var date = LocalDateTime.of(2025, 3, 9, 2, 30);
        var output = new ByteArrayOutputStream();
        ExcelWriter.<LocalDateTime>create().column("Local", value -> value)
                .write(List.of(date)).writeTo(output);
        byte[] file = output.toByteArray();

        TimeZone previous = LocaleUtil.getUserTimeZone();
        try {
            for (String zone : List.of("UTC", "America/New_York", "Asia/Seoul")) {
                var timeZone = TimeZone.getTimeZone(zone);
                LocaleUtil.setUserTimeZone(timeZone);
                ExcelReader.mapping(row -> row.get("Local")).read(new ByteArrayInputStream(file), row -> {
                    assertEquals(date, row.data().asLocalDateTime());
                    assertEquals(timeZone, LocaleUtil.getUserTimeZone());
                });
                assertEquals(List.of("2025-03-09 02:30:00"),
                        ExcelReader.getSheetHeaders(new ByteArrayInputStream(file), 0, 1));
                assertEquals(timeZone, LocaleUtil.getUserTimeZone());
            }
        } finally {
            LocaleUtil.setUserTimeZone(previous);
        }
    }

    @Test
    void infersMissingCellReferencesFromRowOrderAndExplicitGaps() throws IOException {
        byte[] file = workbookWithSheetXml("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row>
                      <c t="inlineStr"><is><t>First</t></is></c>
                      <c r="C1" t="inlineStr"><is><t>Third</t></is></c>
                      <c t="inlineStr"><is><t>Fourth</t></is></c>
                      <c t="inlineStr"><is><t>Fifth</t></is></c>
                    </row>
                    <row r="5">
                      <c><v>10</v></c><c r="C5"><v>30</v></c><c><v>40</v></c><c><f>1+1</f></c>
                    </row>
                    <row>
                      <c><v>11</v></c><c/><c><v>31</v></c><c><v>41</v></c><c><v>51</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
        assertEquals(List.of("First", "", "Third", "Fourth", "Fifth"),
                ExcelReader.getSheetHeaders(new ByteArrayInputStream(file), 0, 0));
        var results = new ArrayList<ReadResult<RowData>>();
        ExcelReader.mapping(row -> row).read(new ByteArrayInputStream(file), results::add);
        assertEquals(5L, results.get(0).fileRowNum());
        assertEquals(6L, results.get(1).fileRowNum());
        var first = results.get(0).data();
        var second = results.get(1).data();
        assertEquals(10L, first.get("First").asLong());
        assertEquals("", first.get(1).formattedValue());
        assertEquals(30L, first.get("Third").asLong());
        assertEquals(40L, first.get("Fourth").asLong());
        assertEquals("", first.get("Fifth").formattedValue());
        assertEquals(51L, second.get("Fifth").asLong());
        assertEquals(31L, second.get("Third").asLong());
    }

    @Test
    void retainsSubMillisecondDateTimesAsTextAndKeepsMillisecondValuesNumeric() {
        var base = LocalDateTime.of(2025, 4, 28, 15, 20, 30);
        var first = base.withNano(123_400_000);
        var second = base.withNano(123_800_000);
        var output = new ByteArrayOutputStream();
        ExcelWriter.<LocalDateTime>create().column("Time", time -> time)
                .write(List.of(first, second, base.withNano(123_000_000))).writeTo(output);
        var cells = ExcelReader.mapping(row -> row.get("Time"))
                .readAll(new ByteArrayInputStream(output.toByteArray()));
        assertEquals(first, cells.get(0).asLocalDateTime());
        assertEquals(second, cells.get(1).asLocalDateTime());
        assertNull(cells.get(0).rawNumber());
        assertNull(cells.get(1).rawNumber());
        assertNotNull(cells.get(2).rawNumber());
        var strict = ExcelWriter.<LocalDateTime>create().column("Time", time -> time)
                .writeErrorPolicy(ExcelWriteErrorPolicy.FAIL_FAST);
        assertThrows(ExcelWriteException.class, () -> strict.write(List.of(first)));
    }

    @Test
    void escapesLiteralSequencesAndControlsWithInlineAndSharedStrings() throws IOException {
        List<Object> values = List.of("_x0041_", "_x0041_x0042_", "_x0041\0", "a\0b\fb\u0001", "a\r\nb",
                "한글😀", "x".repeat(32760) + "_x0041_", new ExcelRichText().bold("_x").italic("0041_\f"));
        for (boolean shared : List.of(false, true)) {
            var output = new ByteArrayOutputStream();
            ExcelWriter.<Object>create(options -> options.useSharedStrings(shared))
                    .column("_x005F_", value -> value).write(values).writeTo(output);
            byte[] workbook = output.toByteArray();
            var rows = ExcelReader.forMap().readAll(new ByteArrayInputStream(workbook));
            assertEquals(values.stream().map(Object::toString).toList(),
                    rows.stream().map(row -> row.get("_x005F_")).toList());
            if (shared) {
                try (var poi = new XSSFWorkbook(new ByteArrayInputStream(workbook))) {
                    var rich = poi.getSheetAt(0).getRow(values.size()).getCell(0).getRichStringCellValue();
                    assertEquals(2, rich.numFormattingRuns());
                }
            }
        }
    }

    @Test
    void isoTypedValuesReadAsDatesDateTimesAndTimes() throws IOException {
        byte[] file = workbookWithSheetXml("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>Date</t></is></c>
                      <c r="B1" t="inlineStr"><is><t>DateTime</t></is></c>
                      <c r="C1" t="inlineStr"><is><t>Time</t></is></c>
                    </row>
                    <row r="2">
                      <c r="A2" s="1" t="d"><v>2025-04-28</v></c>
                      <c r="B2" s="1" t="d"><v>2025-04-28T10:15:21.123</v></c>
                      <c r="C2" s="1" t="d"><v>12:34:56</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
        var row = ExcelReader.mapping(value -> value).readAll(new ByteArrayInputStream(file)).get(0);
        assertEquals(LocalDate.of(2025, 4, 28), row.get("Date").asLocalDate());
        assertEquals(LocalDateTime.of(2025, 4, 28, 10, 15, 21, 123_000_000), row.get("DateTime").asLocalDateTime());
        assertEquals(LocalTime.of(12, 34, 56), row.get("Time").asLocalTime());
        assertNull(row.get("Date").rawNumber());
        assertEquals(List.of("2025-04-28", "2025-04-28T10:15:21.123", "12:34:56"),
                ExcelReader.getSheetHeaders(new ByteArrayInputStream(file), 0, 1));
    }

    @Test
    void distinguishesExcelLeapDayFromMarchAndRetainsTimeOnlyFormats() throws IOException {
        var output = new ByteArrayOutputStream();
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet();
            var header = sheet.createRow(0);
            var row = sheet.createRow(1);
            var formats = workbook.createDataFormat();
            var headings = List.of("Leap", "March", "Time", "DateTime", "Number", "Elapsed");
            var patterns = List.of("yyyy-mm-dd", "yyyy-mm-dd", "hh:mm:ss", "yyyy-mm-dd hh:mm:ss", "0", "[h]:mm");
            var values = List.of(60.0, 61.0, 60.5, 60.5, 60.0, 60.5);
            for (int column = 0; column < headings.size(); column++) {
                header.createCell(column).setCellValue(headings.get(column));
                var cell = row.createCell(column);
                var style = workbook.createCellStyle();
                style.setDataFormat(formats.getFormat(patterns.get(column)));
                cell.setCellStyle(style);
                cell.setCellValue(values.get(column));
            }
            workbook.write(output);
        }
        var row = ExcelReader.mapping(value -> value)
                .readAll(new ByteArrayInputStream(output.toByteArray())).get(0);
        assertEquals("1900-02-29", row.get("Leap").formattedValue());
        assertThrows(DateTimeParseException.class, () -> row.get("Leap").asLocalDate());
        assertEquals(LocalDate.of(1900, 3, 1), row.get("March").asLocalDate());
        assertEquals(LocalTime.NOON, row.get("Time").asLocalTime());
        assertEquals("1900-02-29T12:00:00", row.get("DateTime").formattedValue());
        assertEquals(60L, row.get("Number").asLong());
        assertEquals("1452:00", row.get("Elapsed").formattedValue());
        assertEquals(60L, row.get("Leap").rawNumber().longValueExact());
    }

    @Test
    void cellEqualityIncludesTheNumericValueAndIgnoresEquivalentNumericSpellings() {
        var first = new CellData(0, "1.23457E+12", null, "1234567890123");
        var equivalent = new CellData(0, "1.23457E+12", null, "1.234567890123E+12");
        var different = new CellData(0, "1.23457E+12", null, "1234567890124");
        assertEquals(first, equivalent);
        assertEquals(first.hashCode(), equivalent.hashCode());
        assertNotEquals(first, different);
        assertEquals(2, new HashSet<>(List.of(first, equivalent, different)).size());
    }

    private static byte[] workbookWithSheetXml(String xml) throws IOException {
        var original = new ByteArrayOutputStream();
        try (var workbook = new XSSFWorkbook()) {
            workbook.createSheet();
            var date = workbook.createCellStyle();
            date.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            workbook.write(original);
        }
        var patched = new ByteArrayOutputStream();
        try (var input = new ZipInputStream(new ByteArrayInputStream(original.toByteArray()));
             var output = new ZipOutputStream(patched)) {
            for (ZipEntry entry; (entry = input.getNextEntry()) != null;) {
                output.putNextEntry(new ZipEntry(entry.getName()));
                if ("xl/worksheets/sheet1.xml".equals(entry.getName())) {
                    output.write(xml.getBytes(StandardCharsets.UTF_8));
                } else {
                    input.transferTo(output);
                }
                output.closeEntry();
            }
        }
        return patched.toByteArray();
    }
}
