package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.read.ReadProgress;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicLong;

import static org.junit.jupiter.api.Assertions.*;

/** {@code CsvReader.countRows()}: the total is known while reading, the same as {@code ExcelReader.countRows()}. */
class CsvCountRowsTest {
    @TempDir
    Path dir;

    private static ByteArrayInputStream csv(String text) {
        return new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8));
    }

    private static String rows(int n) {
        var sb = new StringBuilder("id,name\r\n");
        for (int i = 0; i < n; i++) sb.append(i).append(",n").append(i).append("\r\n");
        return sb.toString();
    }

    /** What every progress callback is told as the total, by the older and the newer callback. */
    private record Seen(List<Long> legacyTotals, List<Long> legacyProcessed, List<ReadProgress> events, long delivered) {}

    private static Seen read(CsvReader<Map<String, String>> reader, java.io.InputStream in) {
        List<Long> totals = new ArrayList<>();
        List<Long> processed = new ArrayList<>();
        List<ReadProgress> events = new ArrayList<>();
        AtomicLong delivered = new AtomicLong();
        reader.onProgress(1, (count, cursor) -> {
            processed.add(count);
            totals.add(cursor.getTotalRows());
        }).onReadProgress(1, events::add).read(in, r -> delivered.incrementAndGet());
        return new Seen(totals, processed, events, delivered.get());
    }

    @Nested
    class TheTotal {
        @Test
        void bothCallbacksKnowTheNumberOfDataRows() {
            Seen seen = read(CsvReader.forMap().countRows(), csv(rows(4)));
            assertEquals(List.of(1L, 2L, 3L, 4L), seen.legacyProcessed());
            assertEquals(List.of(4L, 4L, 4L, 4L), seen.legacyTotals());
            assertFalse(seen.events().isEmpty());
            for (ReadProgress event : seen.events()) {
                assertEquals(4, event.totalRows(), "processed " + event.processedRows() + (event.completed() ? " (completed)" : ""));
            }
            assertTrue(seen.events().get(seen.events().size() - 1).completed());
        }

        @Test
        void theReadProgressCallbackNeedsNoOlderCallback() {
            List<ReadProgress> events = new ArrayList<>();
            CsvReader.forMap().countRows().onReadProgress(2, events::add).read(csv(rows(5)), r -> {});
            assertTrue(events.stream().allMatch(e -> e.totalRows() == 5), events.toString());
        }

        @Test
        void withoutCountRowsTheTotalIsUnknownWhileReading() {
            Seen seen = read(CsvReader.forMap(), csv(rows(3)));
            assertEquals(List.of(-1L, -1L, -1L), seen.legacyTotals());
            assertTrue(seen.events().stream().allMatch(e -> e.totalRows() == -1), seen.events().toString());
        }

        @Test
        void aHeaderOnlyFileHasNoRows() {
            Seen seen = read(CsvReader.forMap().countRows(), csv("id,name\r\n"));
            assertEquals(0, seen.delivered());
            assertTrue(seen.events().stream().allMatch(e -> e.totalRows() == 0), seen.events().toString());
        }

        @Test
        void aQuotedFieldWithLineBreaksIsOneRow() {
            String text = "id,note\r\n1,\"first\nsecond\nthird\"\r\n2,\"a,b\"\r\n3,plain\r\n";
            Seen seen = read(CsvReader.forMap().countRows(), csv(text));
            assertEquals(3, seen.delivered());
            assertEquals(List.of(3L, 3L, 3L), seen.legacyTotals(), "three records, not five lines");
        }

        @Test
        void theLastLineNeedsNoLineBreakAndAnyLineEndingCounts() {
            for (String eol : new String[]{"\r\n", "\n", "\r"}) {
                String text = "id,name" + eol + "1,a" + eol + "2,b" + eol + "3,c";
                Seen seen = read(CsvReader.forMap().countRows(), csv(text));
                assertEquals(List.of(3L, 3L, 3L), seen.legacyTotals(), "line ending " + eol.replace("\r", "\\r").replace("\n", "\\n"));
            }
        }
    }

    @Nested
    class WithOtherSettings {
        @Test
        void rowsBeforeTheHeaderAreNotCounted() {
            String text = "title\r\nsubtitle\r\nid,name\r\n1,a\r\n2,b\r\n";
            Seen seen = read(CsvReader.forMap().headerRowIndex(2).countRows(), csv(text));
            assertEquals(2, seen.delivered());
            assertEquals(List.of(2L, 2L), seen.legacyTotals());
        }

        @Test
        void delimiterAndCharsetOfTheReadAreThoseOfTheCount() {
            Charset eucKr = Charset.forName("EUC-KR");
            String text = "id;이름\r\n1;가\r\n2;나\r\n3;다\r\n";
            Seen seen = read(CsvReader.forMap().delimiter(';').charset(eucKr).countRows(), new ByteArrayInputStream(text.getBytes(eucKr)));
            assertEquals(3, seen.delivered());
            assertEquals(List.of(3L, 3L, 3L), seen.legacyTotals());
        }

        @Test
        void aPathIsCountedLikeAStream() throws IOException {
            Path file = dir.resolve("rows.csv");
            Files.writeString(file, rows(6));
            List<Long> totals = new ArrayList<>();
            CsvReader.forMap().countRows().onProgress(3, (count, cursor) -> totals.add(cursor.getTotalRows())).read(file, r -> {});
            assertEquals(List.of(6L, 6L), totals);
            assertTrue(Files.exists(file), "the caller's file is left alone");
        }

        @Test
        void aSetterReaderCountsToo() {
            List<Long> totals = new ArrayList<>();
            CsvReader.setter(Holder::new).column("id", (h, cell) -> h.id = cell.asInt()).countRows()
                    .onProgress(1, (count, cursor) -> totals.add(cursor.getTotalRows())).read(csv(rows(3)), r -> {});
            assertEquals(List.of(3L, 3L, 3L), totals);
        }

        static class Holder {
            int id;
        }

        @Test
        void aSkippedRowIsCountedAsTheExcelReaderDoes() {
            String text = "id,name\r\n1,a\r\n,\r\n2,b\r\n";
            Seen seen = read(CsvReader.forMap().skipBlankRows().countRows(), csv(text));
            assertEquals(2, seen.delivered(), "the blank row is skipped");
            assertEquals(3, seen.events().get(0).totalRows(), "but it is a row of the file: the percentage can end below 100");
        }

        @Test
        void maxRowsStopsTheReadButNotTheCount() {
            Seen seen = read(CsvReader.forMap().maxRows(2).countRows(), csv(rows(5)));
            assertEquals(2, seen.delivered());
            assertTrue(seen.events().stream().allMatch(e -> e.totalRows() == 5), seen.events().toString());
        }
    }

    @Nested
    class WhenTheCountCannotBeMade {
        @Test
        void aFailureIsReportedByTheReadNotByTheCount() {
            // invalid UTF-8: the first pass fails, the total stays unknown, and the read says what is wrong
            byte[] bad = new byte[]{'i', 'd', '\n', '1', (byte) 0xFF, (byte) 0xFE, '\n'};
            CsvReadException withCount = assertThrows(CsvReadException.class,
                    () -> CsvReader.forMap().countRows().read(new ByteArrayInputStream(bad), r -> {}));
            CsvReadException without = assertThrows(CsvReadException.class,
                    () -> CsvReader.forMap().read(new ByteArrayInputStream(bad), r -> {}));
            assertEquals(without.getMessage(), withCount.getMessage());
        }

        @Test
        void aMalformedRecordFailsTheSameWayWithAndWithoutTheCount() {
            for (String bad : new String[]{"id,note\r\n1,\"never closed\r\n2,b\r\n", "id,note\r\n1,a\u0000b\r\n"}) {
                CsvReadException withCount = assertThrows(CsvReadException.class,
                        () -> CsvReader.forMap().countRows().read(csv(bad), r -> {}));
                CsvReadException without = assertThrows(CsvReadException.class,
                        () -> CsvReader.forMap().read(csv(bad), r -> {}));
                assertEquals(without.getMessage(), withCount.getMessage(), bad);
            }
        }

        @Test
        void aFileThatEndsBeforeItsHeaderFailsTheSameWay() {
            CsvReadException withCount = assertThrows(CsvReadException.class,
                    () -> CsvReader.forMap().headerRowIndex(5).countRows().read(csv("id\r\n1\r\n"), r -> {}));
            CsvReadException without = assertThrows(CsvReadException.class,
                    () -> CsvReader.forMap().headerRowIndex(5).read(csv("id\r\n1\r\n"), r -> {}));
            assertEquals(without.getMessage(), withCount.getMessage());
        }

        @Test
        void theTotalIsUnknownWhenTheCountFails() {
            // a header row that is missing: nothing to read, no total
            List<ReadProgress> events = new ArrayList<>();
            assertThrows(CsvReadException.class,
                    () -> CsvReader.forMap().countRows().onReadProgress(1, events::add).read(csv(""), r -> {}));
            assertTrue(events.stream().allMatch(e -> e.totalRows() == -1), events.toString());
        }
    }
}
