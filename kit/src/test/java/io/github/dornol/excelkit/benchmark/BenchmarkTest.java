package io.github.dornol.excelkit.benchmark;

import io.github.dornol.excelkit.csv.CsvMapReader;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelMapReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Performance benchmarks for excel-kit read and write operations.
 * <p>
 * These are not micro-benchmarks (no JMH); they measure end-to-end wall-clock time
 * and output size for representative workloads. Run with:
 * <pre>{@code ./gradlew :kit:benchmark}</pre>
 */
@Tag("benchmark")
class BenchmarkTest {

    @TempDir
    Path tempDir;

    // ============================================================
    // Excel Writing
    // ============================================================

    @Test
    void excel_100k_rows_5_columns() throws IOException {
        int rows = 100_000;
        int cols = 5;
        Path file = tempDir.resolve("bench_100k_5col.xlsx");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        ExcelHandler handler = new ExcelWriter<int[]>()
                .addColumn("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Name", r -> "User-" + r[0])
                .addColumn("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .addColumn("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .addColumn("Note", r -> "This is a sample note for row " + r[0])
                .write(generateRows(rows, cols));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("Excel 100K rows × 5 cols", rows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_1m_rows_5_columns() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("bench_1m_5col.xlsx");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        ExcelHandler handler = new ExcelWriter<int[]>()
                .addColumn("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Name", r -> "User-" + r[0])
                .addColumn("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .addColumn("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .addColumn("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("Excel 1M rows × 5 cols", rows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_100k_rows_50_columns() throws IOException {
        int rows = 100_000;
        int cols = 50;
        Path file = tempDir.resolve("bench_100k_50col.xlsx");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        ExcelWriter<int[]> writer = new ExcelWriter<>();
        for (int i = 0; i < cols; i++) {
            final int idx = i;
            writer.addColumn("Col" + i, r -> r[idx % r.length], c -> c.type(ExcelDataType.INTEGER));
        }
        ExcelHandler handler = writer.write(generateRows(rows, cols));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("Excel 100K rows × 50 cols", rows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_rollover_100k_per_sheet_x_10() throws IOException {
        int totalRows = 1_000_000;
        int rowsPerSheet = 100_000;
        Path file = tempDir.resolve("bench_rollover_10sheets.xlsx");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        ExcelHandler handler = new ExcelWriter<int[]>(rowsPerSheet)
                .addColumn("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .addColumn("Name", r -> "User-" + r[0])
                .addColumn("Value", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .write(generateRows(totalRows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("Excel rollover 100K×10 sheets", totalRows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    @Test
    void excelWorkbook_multiSheet_different_types() throws IOException {
        int rowsPerSheet = 50_000;
        Path file = tempDir.resolve("bench_workbook_multi.xlsx");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
            wb.<int[]>sheet("Numbers")
                    .column("ID", r -> r[0])
                    .column("Value", r -> r[1])
                    .write(generateRows(rowsPerSheet, 3));

            wb.<String>sheet("Strings")
                    .column("Name", s -> s)
                    .column("Upper", s -> s.toUpperCase())
                    .write(IntStream.range(0, rowsPerSheet).mapToObj(i -> "Item-" + i));

            ExcelHandler handler = wb.finish();
            try (OutputStream os = Files.newOutputStream(file)) {
                handler.consumeOutputStream(os);
            }
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("ExcelWorkbook 2 sheets × 50K rows", rowsPerSheet * 2, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    // ============================================================
    // CSV Writing
    // ============================================================

    @Test
    void csv_100k_rows_5_columns() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("bench_100k_5col.csv");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        var handler = new CsvWriter<int[]>()
                .column("ID", r -> r[0])
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1])
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("CSV 100K rows × 5 cols", rows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    @Test
    void csv_1m_rows_5_columns() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("bench_1m_5col.csv");

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        var handler = new CsvWriter<int[]>()
                .column("ID", r -> r[0])
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1])
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.consumeOutputStream(os);
        }

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;
        long fileSize = Files.size(file);

        printResult("CSV 1M rows × 5 cols", rows, elapsed, fileSize, peakMem);
        assertTrue(Files.exists(file));
    }

    // ============================================================
    // Excel Reading
    // ============================================================

    @Test
    void excelRead_100k_rows_mapReader() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("read_100k.xlsx");

        // Write test file
        try (OutputStream os = Files.newOutputStream(file)) {
            new ExcelWriter<int[]>()
                    .addColumn("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Name", r -> "User-" + r[0])
                    .addColumn("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                    .write(generateRows(rows, 3))
                    .consumeOutputStream(os);
        }

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        AtomicLong count = new AtomicLong();
        new ExcelMapReader()
                .build(new FileInputStream(file.toFile()))
                .read(r -> count.incrementAndGet());

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;

        printResult("Excel Read 100K rows (MapReader)", rows, elapsed, Files.size(file), peakMem);
        assertEquals(rows, count.get());
    }

    @Test
    void excelRead_100k_rows_typedReader() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("read_typed_100k.xlsx");

        try (OutputStream os = Files.newOutputStream(file)) {
            new ExcelWriter<int[]>()
                    .addColumn("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Name", r -> "User-" + r[0])
                    .addColumn("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                    .write(generateRows(rows, 3))
                    .consumeOutputStream(os);
        }

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        AtomicLong count = new AtomicLong();
        ExcelReader.<String[]>mapping(row -> new String[]{
                row.get("ID").asString(),
                row.get("Name").asString(),
                row.get("Score").asString()
        }).build(new FileInputStream(file.toFile()))
                .read(r -> count.incrementAndGet());

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;

        printResult("Excel Read 100K rows (typed)", rows, elapsed, Files.size(file), peakMem);
        assertEquals(rows, count.get());
    }

    // ============================================================
    // CSV Reading
    // ============================================================

    @Test
    void csvRead_1m_rows_mapReader() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("read_1m.csv");

        try (OutputStream os = Files.newOutputStream(file)) {
            new CsvWriter<int[]>()
                    .column("ID", r -> r[0])
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1])
                    .write(generateRows(rows, 3))
                    .consumeOutputStream(os);
        }

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        AtomicLong count = new AtomicLong();
        new CsvMapReader()
                .build(new FileInputStream(file.toFile()))
                .read(r -> count.incrementAndGet());

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;

        printResult("CSV Read 1M rows (MapReader)", rows, elapsed, Files.size(file), peakMem);
        assertEquals(rows, count.get());
    }

    @Test
    void csvRead_1m_rows_typedReader() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("read_typed_1m.csv");

        try (OutputStream os = Files.newOutputStream(file)) {
            new CsvWriter<int[]>()
                    .column("ID", r -> r[0])
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1])
                    .write(generateRows(rows, 3))
                    .consumeOutputStream(os);
        }

        long startMem = usedMemoryMB();
        long startTime = System.currentTimeMillis();

        AtomicLong count = new AtomicLong();
        CsvReader.<String[]>mapping(row -> new String[]{
                row.get("ID").asString(),
                row.get("Name").asString(),
                row.get("Score").asString()
        }).build(new FileInputStream(file.toFile()))
                .read(r -> count.incrementAndGet());

        long elapsed = System.currentTimeMillis() - startTime;
        long peakMem = usedMemoryMB() - startMem;

        printResult("CSV Read 1M rows (typed)", rows, elapsed, Files.size(file), peakMem);
        assertEquals(rows, count.get());
    }

    // ============================================================
    // Helpers
    // ============================================================

    private Stream<int[]> generateRows(int count, int cols) {
        return IntStream.range(0, count)
                .mapToObj(i -> {
                    int[] row = new int[Math.max(cols, 3)];
                    row[0] = i;
                    row[1] = i * 10;
                    row[2] = i;
                    for (int j = 3; j < row.length; j++) {
                        row[j] = i + j;
                    }
                    return row;
                });
    }

    private long usedMemoryMB() {
        Runtime rt = Runtime.getRuntime();
        rt.gc();
        return (rt.totalMemory() - rt.freeMemory()) / (1024 * 1024);
    }

    private void printResult(String label, int rows, long elapsedMs, long fileBytes, long memDeltaMB) {
        double seconds = elapsedMs / 1000.0;
        double fileMB = fileBytes / (1024.0 * 1024.0);
        double rowsPerSec = rows / seconds;
        System.out.printf("[BENCH] %-40s | %,10d rows | %6.2fs | %7.2f MB | ~%,d rows/s | mem delta ~%d MB%n",
                label, rows, seconds, fileMB, (long) rowsPerSec, memDeltaMB);
    }
}
