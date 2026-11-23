package io.github.dornol.excelkit.scale;

import io.github.dornol.excelkit.core.read.ReadSummary;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.LongStream;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The features that keep something per merged region or per key, at a size where that shows: run with
 * {@code ./gradlew :kit:scaleTest} (256 MB heap). Unlike the plain row streams of {@link ScaleTest}, these hold state
 * that grows with the data, so each test says how much it may hold instead of expecting a flat heap.
 */
@Tag("scale")
class FeatureScaleTest {

    private static final long ROWS = Long.getLong("excelkit.scale.featureRows", 1_000_000L);

    @TempDir
    Path dir;

    private static long usedMiB() {
        for (int i = 0; i < 3; i++) System.gc();
        Runtime r = Runtime.getRuntime();
        return (r.totalMemory() - r.freeMemory()) / (1024 * 1024);
    }

    private record Line(long id) {
        /** A department every {@code RUN} rows, so the merge makes one region per {@code RUN} rows. */
        String dept() {
            return "dept-" + (id / RUN);
        }

        String name() {
            return "name-" + id;
        }
    }

    private static final long RUN = 5;

    private static long countMergedRegions(Path xlsx, String sheetEntry) throws IOException {
        try (ZipFile zip = new ZipFile(xlsx.toFile()); var in = zip.getInputStream(zip.getEntry(sheetEntry))) {
            // a merged region is written as <mergeCell ref="..."/>; count them without parsing the sheet
            long count = 0;
            byte[] pattern = "<mergeCell ".getBytes();
            byte[] buffer = new byte[1 << 16];
            int matched = 0, read;
            while ((read = in.read(buffer)) > 0) {
                for (int i = 0; i < read; i++) {
                    if (buffer[i] == pattern[matched]) {
                        if (++matched == pattern.length) { count++; matched = 0; }
                    } else matched = buffer[i] == pattern[0] ? 1 : 0;
                }
            }
            return count;
        }
    }

    @Test
    void mergeSameValuesOverARolloverSheetWithManyRegions() throws IOException {
        Path file = dir.resolve("merged.xlsx");
        long before = usedMiB();
        long started = System.nanoTime();
        ExcelWriter.<Line>create()
                .column("Dept", Line::dept, c -> c.mergeSameValues())
                .column("Name", Line::name)
                .write(LongStream.range(0, ROWS).mapToObj(Line::new))
                .writeTo(file);
        double seconds = (System.nanoTime() - started) / 1e9;
        long regions = countMergedRegions(file, "xl/worksheets/sheet1.xml");
        System.out.printf("[scale] merged %,d rows into %,d regions in %.1fs, file %,d bytes, heap before %d MiB%n",
                ROWS, regions, seconds, Files.size(file), before);
        // the sheet holds up to 1,000,000 rows: one region per RUN rows
        assertEquals(Math.min(ROWS, 1_000_000L) / RUN, regions);
        Files.delete(file);
    }

    private static long countOccurrences(Path xlsx, String sheetEntry, String needle) throws IOException {
        try (ZipFile zip = new ZipFile(xlsx.toFile()); var in = zip.getInputStream(zip.getEntry(sheetEntry))) {
            long count = 0;
            byte[] pattern = needle.getBytes();
            byte[] buffer = new byte[1 << 16];
            int matched = 0, read;
            while ((read = in.read(buffer)) > 0) {
                for (int i = 0; i < read; i++) {
                    if (buffer[i] == pattern[matched]) {
                        if (++matched == pattern.length) { count++; matched = 0; }
                    } else matched = buffer[i] == pattern[0] ? 1 : 0;
                }
            }
            return count;
        }
    }

    @Test
    void groupSubtotalsKeepNothingPerGroupAndTurnTheSheetOverBeforeItIsFull() throws IOException {
        Path file = dir.resolve("subtotals.xlsx");
        AtomicLong peak = new AtomicLong();
        long started = System.nanoTime();
        // a subtotal row for every RUN rows: with the default maxRows the sheet would run out of room for them
        ExcelWriter.<Line>create()
                .column("Dept", Line::dept)
                .column("Name", Line::name)
                .column("Id", Line::id)
                .groupSubtotal(Line::dept, s -> s.label("Name", k -> "total").sum("Id"))
                .onProgress(250_000, (total, cursor) -> peak.accumulateAndGet(usedMiB(), Math::max))
                .write(LongStream.range(0, ROWS).mapToObj(Line::new))
                .writeTo(file);
        double seconds = (System.nanoTime() - started) / 1e9;
        long dataRows = 0, subtotalRows = 0;
        int sheets = 0;
        try (ZipFile zip = new ZipFile(file.toFile())) {
            while (zip.getEntry("xl/worksheets/sheet" + (sheets + 1) + ".xml") != null) sheets++;
        }
        for (int sheet = 1; sheet <= sheets; sheet++) {
            String entry = "xl/worksheets/sheet" + sheet + ".xml";
            long rows = countOccurrences(file, entry, "<row ");
            long sub = countOccurrences(file, entry, "SUBTOTAL(9,");
            assertTrue(rows <= 1_048_576, "sheet " + sheet + " holds " + rows + " rows");
            dataRows += rows - 1 - sub;
            subtotalRows += sub;
        }
        System.out.printf("[scale] %,d rows with %,d subtotal rows over %d sheets in %.1fs, peak heap after gc %d MiB%n",
                dataRows, subtotalRows, sheets, seconds, peak.get());
        assertEquals(ROWS, dataRows, "no data row lost");
        // a subtotal row for every run of RUN rows, plus one more where a sheet ends inside a run (its last group is closed there)
        long runs = (ROWS + RUN - 1) / RUN;
        assertTrue(subtotalRows >= runs && subtotalRows <= runs + sheets - 1,
                subtotalRows + " subtotal rows for " + runs + " runs over " + sheets + " sheets");
        assertTrue(ROWS < 900_000 || sheets >= 2, "the rows and their subtotals do not fit one sheet");
        assertTrue(peak.get() < 200, "the subtotals held " + peak.get() + " MiB");
        Files.delete(file);
    }

    @Test
    void fillMergedCellsReadsABigSheetWithManyRegionsInBoundedHeap() throws IOException {
        Path file = dir.resolve("merged-read.xlsx");
        long rows = Math.min(ROWS, 1_000_000L);
        ExcelWriter.<Line>create()
                .column("Dept", Line::dept, c -> c.mergeSameValues())
                .column("Name", Line::name)
                .write(LongStream.range(0, rows).mapToObj(Line::new))
                .writeTo(file);
        // this file has 200,000 regions, the most the option accepts; the writer keeps the covered cells' values, so
        // the check is the cost of the scan and that the rows come back right, not the filling itself
        long started = System.nanoTime();
        AtomicLong seen = new AtomicLong();
        AtomicLong peak = new AtomicLong();
        ReadSummary summary = ExcelReader.<String>mapping(row -> row.get("Dept").asString()).fillMergedCells()
                .readWithSummary(file, result -> {
                    assertTrue(result.success(), result.messages().toString());
                    long n = seen.getAndIncrement();
                    assertEquals("dept-" + (n / RUN), result.data(), "row " + n);
                    if (n % 250_000 == 0) peak.accumulateAndGet(usedMiB(), Math::max);
                });
        System.out.printf("[scale] read %,d rows with %,d merged regions in %.1fs, peak heap after gc %d MiB%n",
                seen.get(), rows / RUN, (System.nanoTime() - started) / 1e9, peak.get());
        assertEquals(rows, summary.totalRows());
        assertTrue(peak.get() < 200, "the scan of the merged regions held " + peak.get() + " MiB");
        Files.delete(file);
    }

    @Test
    void uniqueKeyHoldsOneEntryPerKey() throws IOException {
        Path file = dir.resolve("keys.csv");
        long rows = ROWS;
        try (OutputStream out = new BufferedOutputStream(Files.newOutputStream(file))) {
            CsvWriter.<Line>create().column("Id", Line::id).column("Name", Line::name)
                    .write(LongStream.range(0, rows).mapToObj(Line::new)).writeTo(out);
        }
        AtomicLong seen = new AtomicLong();
        AtomicLong errors = new AtomicLong();
        AtomicLong peak = new AtomicLong();
        long started = System.nanoTime();
        CsvReader.<String>mapping(row -> row.get("Id").asString()).uniqueKey("Id")
                .readWithSummary(file, result -> {
                    long n = seen.getAndIncrement();
                    if (!result.success()) errors.incrementAndGet();
                    if (n % 250_000 == 0) peak.accumulateAndGet(usedMiB(), Math::max);
                });
        System.out.printf("[scale] uniqueKey over %,d distinct keys in %.1fs, peak heap after gc %d MiB, errors %d%n",
                rows, (System.nanoTime() - started) / 1e9, peak.get(), errors.get());
        assertEquals(0, errors.get());
        assertEquals(rows, seen.get());
        // measured: about 95 MiB for 1,000,000 keys (some 85 bytes a key); a sheet holds at most about a million rows
        assertTrue(peak.get() < 160, "the keys held " + peak.get() + " MiB for " + rows + " keys");
        Files.delete(file);
    }
}
