package io.github.dornol.excelkit.scale;

import io.github.dornol.excelkit.excel.read.ExcelWorkbookInspector;
import io.github.dornol.excelkit.core.read.ReadSummary;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.read.ExcelSheetInfo;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.LongStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * What must hold when the data is big. Run with {@code ./gradlew :kit:scaleTest}, which gives the test JVM a
 * 256 MB heap: a writer or reader that held its rows in memory would not get through.
 * <ul>
 *   <li>every row arrives, in order, with the values it was written with, also across rollover sheets;</li>
 *   <li>the heap does not grow with the number of rows (measured after a garbage collection every 500,000 rows);</li>
 *   <li>no temp file is left behind (the leak guard checks that after each test).</li>
 * </ul>
 */
@Tag("scale")
class ScaleTest {

    private static final long EXCEL_ROWS = Long.getLong("excelkit.scale.excelRows", 2_500_000L);
    private static final long CSV_ROWS = Long.getLong("excelkit.scale.csvRows", 3_000_000L);
    private static final long SAMPLE_EVERY = 500_000L;
    /** What may be added to the heap between the first sample and the last, while the data stays out of it. */
    private static final long ALLOWED_GROWTH_BYTES = 64L * 1024 * 1024;
    private static final LocalDate BASE = LocalDate.of(2000, 1, 1);

    @TempDir
    Path dir;

    /** One generated row; everything in it follows from {@code id}, so a reader can check it. */
    private record Line(long id) {
        String name() {
            return "name-" + id;
        }

        int score() {
            return (int) (id % 1000);
        }

        LocalDate day() {
            return BASE.plusDays(id % 3650);
        }
    }

    private record Seen(long id, String name, int score, LocalDate day, long number) {
    }

    private static long usedAfterGc() {
        for (int i = 0; i < 3; i++) System.gc();
        Runtime runtime = Runtime.getRuntime();
        return runtime.totalMemory() - runtime.freeMemory();
    }

    /** Samples the heap every {@link #SAMPLE_EVERY} rows and fails if it grew by more than the allowance. */
    private static final class HeapWatch {
        private final String what;
        private final List<Long> samples = new ArrayList<>();

        HeapWatch(String what) {
            this.what = what;
        }

        void row(long index) {
            if (index % SAMPLE_EVERY == 0) samples.add(usedAfterGc());
        }

        void verify() {
            System.out.println("[scale] " + what + " heap after gc (MiB): "
                    + samples.stream().map(bytes -> bytes / (1024 * 1024)).toList());
            assertTrue(samples.size() >= 2, what + ": too few samples to say anything (" + samples.size() + ")");
            long growth = samples.get(samples.size() - 1) - samples.get(0);
            assertTrue(growth < ALLOWED_GROWTH_BYTES, what + ": the heap grew by " + growth / (1024 * 1024)
                    + " MiB while the rows went through " + samples);
        }
    }

    // ---- Excel

    @Test
    void excelRowsSurviveRolloverInOrderWithBoundedHeap() throws IOException {
        Path file = dir.resolve("big.xlsx");
        HeapWatch writing = new HeapWatch("writing " + EXCEL_ROWS + " Excel rows");
        long started = System.nanoTime();
        AtomicLong produced = new AtomicLong();

        ExcelWriter.<Line>create()
                .column("Id", Line::id, c -> c.type(ExcelDataType.LONG))
                .column("Name", Line::name)
                .column("Score", Line::score)
                .column("Day", Line::day)
                .rowNumberColumn("No")
                .write(LongStream.range(0, EXCEL_ROWS).mapToObj(id -> {
                    writing.row(produced.getAndIncrement());
                    return new Line(id);
                }))
                .writeTo(file);
        long written = System.nanoTime();
        System.out.printf("[scale] wrote %,d Excel rows in %.1fs, %,d bytes%n", EXCEL_ROWS, (written - started) / 1e9, Files.size(file));
        writing.verify();

        long expectedSheets = (EXCEL_ROWS + 999_999) / 1_000_000;
        List<ExcelSheetInfo> sheets = ExcelWorkbookInspector.create().sheetNames(file);
        assertEquals(expectedSheets, sheets.size(), "one sheet per million rows: " + sheets);
        assertEquals(sheets.size(), sheets.stream().map(ExcelSheetInfo::name).distinct().count(), "sheet names are unique: " + sheets);

        HeapWatch reading = new HeapWatch("reading " + EXCEL_ROWS + " Excel rows");
        AtomicLong expectedId = new AtomicLong();
        AtomicLong total = new AtomicLong();
        for (ExcelSheetInfo sheet : sheets) {
            long startId = expectedId.get();
            ReadSummary summary = ExcelReader.<Seen>mapping(row -> new Seen(row.get("Id").asLong(), row.get("Name").asString(),
                            row.get("Score").asInt(), row.get("Day").asLocalDate(), row.get("No").asLong()))
                    .sheetIndex(sheet.index())
                    .readWithSummary(file, result -> {
                        assertTrue(result.success(), "row " + expectedId.get() + " did not read: " + result.messages());
                        Seen seen = result.data();
                        long id = expectedId.getAndIncrement();
                        Line line = new Line(id);
                        // a mismatch names the first row that is wrong, not the millionth
                        if (seen.id() != id || !seen.name().equals(line.name()) || seen.score() != line.score()
                                || !seen.day().equals(line.day()) || seen.number() != id + 1) {
                            throw new AssertionError("row " + id + " on sheet " + sheet.name() + " came back as " + seen);
                        }
                        reading.row(total.getAndIncrement());
                    });
            assertEquals(Math.min(1_000_000L, EXCEL_ROWS - startId), summary.totalRows(), "rows on " + sheet.name());
            assertEquals(0, summary.errorRows(), "errors on " + sheet.name());
        }
        System.out.printf("[scale] read %,d Excel rows in %.1fs%n", total.get(), (System.nanoTime() - written) / 1e9);
        assertEquals(EXCEL_ROWS, total.get());
        reading.verify();
        Files.delete(file);
    }

    // ---- CSV

    @Test
    void csvRowsSurviveInOrderWithBoundedHeap() throws IOException {
        Path file = dir.resolve("big.csv");
        HeapWatch writing = new HeapWatch("writing " + CSV_ROWS + " CSV rows");
        AtomicLong produced = new AtomicLong();
        long started = System.nanoTime();

        try (OutputStream out = new BufferedOutputStream(Files.newOutputStream(file))) {
            CsvWriter.<Line>create()
                    .column("Id", Line::id).column("Name", Line::name).column("Score", Line::score).column("Day", Line::day)
                    .write(LongStream.range(0, CSV_ROWS).mapToObj(id -> {
                        writing.row(produced.getAndIncrement());
                        return new Line(id);
                    }))
                    .writeTo(out);
        }
        long written = System.nanoTime();
        System.out.printf("[scale] wrote %,d CSV rows in %.1fs, %,d bytes%n", CSV_ROWS, (written - started) / 1e9, Files.size(file));
        writing.verify();

        HeapWatch reading = new HeapWatch("reading " + CSV_ROWS + " CSV rows");
        AtomicLong expectedId = new AtomicLong();
        AtomicLong checksum = new AtomicLong();
        ReadSummary summary = CsvReader.<Seen>mapping(row -> new Seen(row.get("Id").asLong(), row.get("Name").asString(),
                        row.get("Score").asInt(), row.get("Day").asLocalDate(), 0))
                .readWithSummary(file, result -> {
                    assertTrue(result.success(), "row " + expectedId.get() + " did not read: " + result.messages());
                    Seen seen = result.data();
                    long id = expectedId.getAndIncrement();
                    Line line = new Line(id);
                    if (seen.id() != id || !seen.name().equals(line.name()) || seen.score() != line.score()
                            || !seen.day().equals(line.day())) {
                        throw new AssertionError("row " + id + " came back as " + seen);
                    }
                    checksum.addAndGet(seen.id());
                    reading.row(id);
                });
        System.out.printf("[scale] read %,d CSV rows in %.1fs%n", summary.totalRows(), (System.nanoTime() - written) / 1e9);
        assertEquals(CSV_ROWS, summary.totalRows());
        assertEquals(CSV_ROWS * (CSV_ROWS - 1) / 2, checksum.get(), "the ids add up");
        reading.verify();
    }

    /** A cell value that does not compress well, so the file is as big as the row count suggests. */
    private static String noise(long n) {
        long x = n * 0x9E3779B97F4A7C15L;
        x ^= x >>> 29;
        x *= 0xBF58476D1CE4E5B9L;
        x ^= x >>> 32;
        return Long.toHexString(x) + Long.toHexString(x * 31);
    }

    /**
     * An encrypted workbook of well over 100 MB is written and read back within the 256 MB heap. POI used to
     * build the encrypted container in heap, which needed about twice the size of the file.
     */
    @Test
    void encryptingABigWorkbookDoesNotNeedItsSizeInHeap() throws IOException {
        long rows = 1_000_000L;
        Path file = dir.resolve("encrypted.xlsx");
        ExcelWriter<Long> writer = ExcelWriter.<Long>create().password("secret");
        for (int c = 0; c < 6; c++) {
            int column = c;
            writer.column("C" + c, n -> noise(n * 8 + column));
        }
        writer.write(LongStream.range(0, rows).boxed().toList()).writeTo(file);

        assertTrue(Files.size(file) > 100L * 1024 * 1024, "the test needs a file bigger than the heap allows in memory");
        AtomicLong count = new AtomicLong();
        AtomicLong mismatches = new AtomicLong();
        ExcelReader.forMap().password("secret").read(file, result -> {
            long id = count.getAndIncrement();
            if (!result.success() || !noise(id * 8).equals(result.data().get("C0"))
                    || !noise(id * 8 + 5).equals(result.data().get("C5"))) {
                mismatches.incrementAndGet();
            }
        });

        assertEquals(rows, count.get());
        assertEquals(0, mismatches.get());
    }
}
