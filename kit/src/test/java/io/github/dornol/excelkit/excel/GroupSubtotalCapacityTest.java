package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.stream.IntStream;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Subtotal rows are not data rows, so {@code maxRows} does not count them, but a sheet holds at most 1,048,576 rows
 * in all: a sheet that is out of room for them turns over early instead of failing. These tests write about a million
 * rows each, and read the sheets back as raw XML (POI would need gigabytes to open them).
 */
class GroupSubtotalCapacityTest {
    record Row(int id, String key) {}

    /** The most maxRows allows: only the room left in the sheet can turn it over. */
    private static final int MAX_ROWS = 1_048_575;

    @TempDir
    Path dir;

    /** Counts how often {@code needle} occurs in the entry, reading it in chunks. */
    private static long count(ZipFile zip, String entry, String needle) throws IOException {
        byte[] pattern = needle.getBytes();
        long count = 0;
        int matched = 0;
        try (InputStream in = zip.getInputStream(zip.getEntry(entry))) {
            byte[] buffer = new byte[1 << 16];
            int read;
            while ((read = in.read(buffer)) > 0) {
                for (int i = 0; i < read; i++) {
                    if (buffer[i] == pattern[matched]) {
                        if (++matched == pattern.length) { count++; matched = 0; }
                    } else {
                        matched = buffer[i] == pattern[0] ? 1 : 0;
                    }
                }
            }
        }
        return count;
    }

    /** The first {@code limit} bytes of the entry as text. */
    private static String head(ZipFile zip, String entry, int limit) throws IOException {
        try (InputStream in = zip.getInputStream(zip.getEntry(entry))) {
            return new String(in.readNBytes(limit));
        }
    }

    private static int sheets(ZipFile zip) {
        int n = 0;
        while (zip.getEntry("xl/worksheets/sheet" + (n + 1) + ".xml") != null) n++;
        return n;
    }

    /** The style index of a cell, read from the start of the sheet. */
    private static String style(String head, String cell) {
        int at = head.indexOf("<c r=\"" + cell + "\"");
        assertTrue(at >= 0, cell + " in " + head.substring(0, Math.min(300, head.length())));
        int s = head.indexOf(" s=\"", at);
        int end = head.indexOf('"', s + 4);
        return head.substring(s + 4, end);
    }

    @Test
    void oneBigGroupTurnsTheSheetOverBeforeItRunsOutOfRoom() throws IOException {
        Path file = dir.resolve("one-group.xlsx");
        int rows = 1_048_575;
        ExcelWriter.<Row>create().column("Key", Row::key).column("Id", Row::id).maxRows(MAX_ROWS)
                .groupSubtotal(Row::key, s -> s.label("Key", k -> "t").sum("Id"))
                .write(IntStream.rangeClosed(1, rows).mapToObj(i -> new Row(i, "A"))).writeTo(file);
        try (ZipFile zip = new ZipFile(file.toFile())) {
            assertEquals(2, sheets(zip), "the sheet turns over for lack of room, not for maxRows");
            long dataRows = 0, subtotals = 0;
            for (int sheet = 1; sheet <= 2; sheet++) {
                String entry = "xl/worksheets/sheet" + sheet + ".xml";
                long cells = count(zip, entry, "<row ");
                long sub = count(zip, entry, "SUBTOTAL(9,");
                assertEquals(1, sub, "one subtotal row in sheet " + sheet);
                dataRows += cells - 1 - sub;
                subtotals += sub;
            }
            assertEquals(rows, dataRows, "no row lost");
            assertEquals(2, subtotals);
        }
    }

    @Test
    void everyRowItsOwnGroupFillsSheetsWithDataAndSubtotalRowsAlike() throws IOException {
        Path file = dir.resolve("own-groups.xlsx");
        int rows = 540_000;
        ExcelWriter.<Row>create().column("Key", Row::key).column("Id", Row::id)
                .groupSubtotal(Row::key, s -> s.label("Key", k -> "t").sum("Id"))
                .write(IntStream.range(0, rows).mapToObj(i -> new Row(i, "k" + i))).writeTo(file);
        try (ZipFile zip = new ZipFile(file.toFile())) {
            int sheets = sheets(zip);
            assertTrue(sheets >= 2, "540,000 data rows and as many subtotal rows do not fit one sheet: " + sheets);
            long dataRows = 0, subtotals = 0;
            for (int sheet = 1; sheet <= sheets; sheet++) {
                String entry = "xl/worksheets/sheet" + sheet + ".xml";
                long cells = count(zip, entry, "<row ");
                long sub = count(zip, entry, "SUBTOTAL(9,");
                assertTrue(cells <= 1_048_576, "sheet " + sheet + " holds " + cells + " rows");
                dataRows += cells - 1 - sub;
                subtotals += sub;
            }
            assertEquals(rows, dataRows);
            assertEquals(rows, subtotals, "a subtotal row for every group, on the sheet of its rows");
        }
    }

    @Test
    void theSummaryStillFitsBelowTheLastSubtotalRow() throws IOException {
        Path file = dir.resolve("summary.xlsx");
        int rows = 1_048_575;
        ExcelWriter.<Row>create().column("Key", Row::key).column("Id", Row::id).maxRows(MAX_ROWS)
                .groupSubtotal(Row::key, s -> s.label("Key", k -> "t").sum("Id"))
                .summary(s -> s.label("Key", "Total").sum("Id").average("Id").count("Id"))
                .write(IntStream.rangeClosed(1, rows).mapToObj(i -> new Row(i, "A"))).writeTo(file);
        try (ZipFile zip = new ZipFile(file.toFile())) {
            int sheets = sheets(zip);
            assertEquals(2, sheets);
            for (int sheet = 1; sheet <= sheets; sheet++) {
                String entry = "xl/worksheets/sheet" + sheet + ".xml";
                assertEquals(1, count(zip, entry, "SUBTOTAL(1,"), "the average row of sheet " + sheet);
                assertEquals(1, count(zip, entry, "SUBTOTAL(2,"), "the count row of sheet " + sheet);
                assertTrue(count(zip, entry, "<row ") <= 1_048_576);
            }
        }
    }

    @Test
    void stripesStartAgainOnASheetThatTurnedOverEarly() throws IOException {
        Path file = dir.resolve("stripes.xlsx");
        ExcelWriter.<Row>create().column("Key", Row::key).column("Id", Row::id).maxRows(MAX_ROWS).stripedRows()
                .groupSubtotal(Row::key, s -> s.label("Key", k -> "t").sum("Id"))
                .write(IntStream.rangeClosed(1, 1_048_575).mapToObj(i -> new Row(i, "A"))).writeTo(file);
        try (ZipFile zip = new ZipFile(file.toFile())) {
            assertEquals(2, sheets(zip));
            String first = head(zip, "xl/worksheets/sheet1.xml", 4000);
            String second = head(zip, "xl/worksheets/sheet2.xml", 4000);
            // data rows are rows 2, 3, 4: the second one is shaded, on both sheets
            assertNotEquals(style(first, "A2"), style(first, "A3"));
            assertEquals(style(first, "A2"), style(first, "A4"));
            assertEquals(style(first, "A2"), style(second, "A2"), "the pattern starts again: the first row is plain");
            assertEquals(style(first, "A3"), style(second, "A3"));
        }
    }

    @Test
    void aWorkbookSheetTurnsOverEarlyToo() throws IOException {
        Path file = dir.resolve("workbook.xlsx");
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<Row>sheet("Data").column("Key", Row::key).column("Id", Row::id).maxRows(MAX_ROWS)
                    .groupSubtotal(Row::key, s -> s.label("Key", k -> "t").sum("Id"))
                    .write(IntStream.rangeClosed(1, 1_048_575).mapToObj(i -> new Row(i, "A")));
            workbook.finish().writeTo(file);
        }
        try (ZipFile zip = new ZipFile(file.toFile())) {
            assertEquals(2, sheets(zip));
            assertTrue(count(zip, "xl/worksheets/sheet1.xml", "<row ") <= 1_048_576);
        }
    }
}
