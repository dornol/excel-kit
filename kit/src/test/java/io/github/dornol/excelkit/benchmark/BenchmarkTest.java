package io.github.dornol.excelkit.benchmark;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelWorkbookReader;
import io.github.dornol.excelkit.excel.WorkbookReadReport;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.TestInfo;
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

    private BenchmarkProbe probe;
    private String scenario;

    @BeforeEach
    void identifyScenario(TestInfo info) {
        scenario = info.getTestMethod().orElseThrow().getName();
    }

    @AfterEach
    void stopSampler() throws IOException {
        if (probe != null) probe.close();
    }

    // ============================================================
    // Excel Writing
    // ============================================================

    @Test
    void excel_100k_rows_5_columns() throws IOException {
        int rows = 100_000;
        int cols = 5;
        Path file = tempDir.resolve("bench_100k_5col.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelHandler handler = ExcelWriter.<int[]>create()
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "This is a sample note for row " + r[0])
                .write(generateRows(rows, cols));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 100K rows × 5 cols", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_1m_rows_5_columns() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("bench_1m_5col.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelHandler handler = ExcelWriter.<int[]>create()
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 1M rows × 5 cols", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_100k_rows_50_columns() throws IOException {
        int rows = 100_000;
        int cols = 50;
        Path file = tempDir.resolve("bench_100k_50col.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelWriter<int[]> writer = ExcelWriter.<int[]>create();
        for (int i = 0; i < cols; i++) {
            final int idx = i;
            writer.column("Col" + i, r -> r[idx % r.length], c -> c.type(ExcelDataType.INTEGER));
        }
        ExcelHandler handler = writer.write(generateRows(rows, cols));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 100K rows × 50 cols", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_rollover_100k_per_sheet_x_10() throws IOException {
        int totalRows = 1_000_000;
        int rowsPerSheet = 100_000;
        Path file = tempDir.resolve("bench_rollover_10sheets.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelHandler handler = ExcelWriter.<int[]>create().maxRows(rowsPerSheet)
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Value", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .write(generateRows(totalRows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel rollover 100K×10 sheets", totalRows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void excelWorkbook_multiSheet_different_types() throws IOException {
        int rowsPerSheet = 50_000;
        Path file = tempDir.resolve("bench_workbook_multi.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        try (ExcelWorkbook wb = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
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
                handler.writeTo(os);
            }
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("ExcelWorkbook 2 sheets × 50K rows", rowsPerSheet * 2, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    // ============================================================
    // Excel Password Encryption
    //
    // POI's agile encryption buffers the entire OLE2 container in heap
    // (no file-backed variant). These benchmarks quantify the time &
    // memory cost vs a plain (unencrypted) write of the same data so
    // users sizing heap for encrypted exports have a reference point.
    // ============================================================

    @Test
    void excel_100k_rows_5_columns_encrypted() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("bench_100k_5col_encrypted.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelWriter.<int[]>create()
                .password("benchPw")
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "This is a sample note for row " + r[0])
                .write(generateRows(rows, 5))
                .writeTo(file);

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 100K × 5 cols ENCRYPTED", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
        // Sanity-check the file is actually encrypted (OLE2 magic bytes).
        byte[] head = Files.readAllBytes(file);
        assertEquals((byte) 0xD0, head[0]);
        assertEquals((byte) 0xCF, head[1]);
    }

    @Test
    void excel_1m_rows_5_columns_encrypted() throws IOException {
        // The encryption step buffers the whole XLSX payload in heap.
        // For 1M rows the plain XLSX is ~30-40MB, so plan ~hundreds of
        // MB of heap for the POIFS/Encryptor step on top of the SXSSF
        // window. This benchmark confirms the end-to-end path completes
        // and gives a concrete number for the guide.
        int rows = 1_000_000;
        Path file = tempDir.resolve("bench_1m_5col_encrypted.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelWriter.<int[]>create()
                .password("benchPw")
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Row " + r[0])
                .write(generateRows(rows, 5))
                .writeTo(file);

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 1M × 5 cols ENCRYPTED", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void excel_100k_rows_encrypted_via_handler_writeTo_path() throws IOException {
        // Exercises the late-binding path: handler.writeTo(Path, String)
        // added in 0.16.6. Same workload as the writer-bound encrypted
        // benchmark so results are directly comparable.
        int rows = 100_000;
        Path file = tempDir.resolve("bench_100k_writeto_path.xlsx");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        ExcelHandler handler = ExcelWriter.<int[]>create()
                .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Row " + r[0])
                .write(generateRows(rows, 5));

        handler.writeTo(file, "benchPw");

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("Excel 100K × 5 cols writeTo(Path,pw)", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    // ============================================================
    // CSV Writing
    // ============================================================

    @Test
    void csv_100k_rows_5_columns() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("bench_100k_5col.csv");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        var handler = CsvWriter.<int[]>create()
                .column("ID", r -> r[0])
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1])
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("CSV 100K rows × 5 cols", rows, elapsed, fileSize, resources);
        assertTrue(Files.exists(file));
    }

    @Test
    void csv_1m_rows_5_columns() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("bench_1m_5col.csv");

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        var handler = CsvWriter.<int[]>create()
                .column("ID", r -> r[0])
                .column("Name", r -> "User-" + r[0])
                .column("Score", r -> r[1])
                .column("Active", r -> r[2] % 2 == 0 ? "Y" : "N")
                .column("Note", r -> "Note-" + r[0])
                .write(generateRows(rows, 5));

        try (OutputStream os = Files.newOutputStream(file)) {
            handler.writeTo(os);
        }

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();
        long fileSize = Files.size(file);

        printResult("CSV 1M rows × 5 cols", rows, elapsed, fileSize, resources);
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
            ExcelWriter.<int[]>create()
                    .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                    .write(generateRows(rows, 3))
                    .writeTo(os);
        }

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        AtomicLong count = new AtomicLong();
        ExcelReader.forMap()
                .read(new FileInputStream(file.toFile()), r -> count.incrementAndGet());

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();

        printResult("Excel Read 100K rows (MapReader)", rows, elapsed, Files.size(file), resources);
        assertEquals(rows, count.get());
    }

    @Test
    void excelRead_100k_rows_typedReader() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("read_typed_100k.xlsx");

        try (OutputStream os = Files.newOutputStream(file)) {
            ExcelWriter.<int[]>create()
                    .column("ID", r -> r[0], c -> c.type(ExcelDataType.INTEGER))
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1], c -> c.type(ExcelDataType.DOUBLE))
                    .write(generateRows(rows, 3))
                    .writeTo(os);
        }

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        AtomicLong count = new AtomicLong();
        ExcelReader.<String[]>mapping(row -> new String[]{
                row.get("ID").asString(),
                row.get("Name").asString(),
                row.get("Score").asString()
        }).read(new FileInputStream(file.toFile()), r -> count.incrementAndGet());

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();

        printResult("Excel Read 100K rows (typed)", rows, elapsed, Files.size(file), resources);
        assertEquals(rows, count.get());
    }

    @Test
    void excelRead_100k_rows_pathInput() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("read_path_100k.xlsx");
        writeExcelRows(file, rows);

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();
        AtomicLong count = new AtomicLong();
        ExcelReader.forMap().read(file, result -> count.incrementAndGet());

        printResult("Excel Read 100K rows (Path)", rows,
                System.nanoTime() - startTime, Files.size(file), probe.stop());
        assertEquals(rows, count.get());
    }

    @Test
    void excelRead_100k_rows_inputStreamWithCountRows() throws IOException {
        int rows = 100_000;
        Path file = tempDir.resolve("read_count_rows_100k.xlsx");
        writeExcelRows(file, rows);

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();
        AtomicLong count = new AtomicLong();
        try (FileInputStream input = new FileInputStream(file.toFile())) {
            ExcelReader.forMap().countRows().read(input, result -> count.incrementAndGet());
        }

        printResult("Excel Read 100K rows (InputStream + countRows)", rows,
                System.nanoTime() - startTime, Files.size(file), probe.stop());
        assertEquals(rows, count.get());
    }

    @Test
    void excelRead_multiSheet_50kRowsEach_sharedPackage() throws IOException {
        int rowsPerSheet = 50_000;
        Path file = tempDir.resolve("read_multi_sheet_50k.xlsx");
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<int[]>sheet("Numbers")
                    .column("ID", row -> row[0])
                    .column("Value", row -> row[1])
                    .write(generateRows(rowsPerSheet, 3));
            workbook.<String>sheet("Strings")
                    .column("Name", value -> value)
                    .write(IntStream.range(0, rowsPerSheet).mapToObj(i -> "Item-" + i));
            workbook.finish().writeTo(file);
        }

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();
        AtomicLong count = new AtomicLong();
        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet("Numbers", ExcelReader.forMap(), result -> count.incrementAndGet())
                .sheet("Strings", ExcelReader.forMap(), result -> count.incrementAndGet())
                .read(file);

        printResult("ExcelWorkbookReader 2 sheets × 50K rows", rowsPerSheet * 2,
                System.nanoTime() - startTime, Files.size(file), probe.stop());
        assertTrue(report.successful());
        assertEquals(rowsPerSheet * 2L, count.get());
    }

    // ============================================================
    // CSV Reading
    // ============================================================

    @Test
    void csvRead_1m_rows_mapReader() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("read_1m.csv");

        try (OutputStream os = Files.newOutputStream(file)) {
            CsvWriter.<int[]>create()
                    .column("ID", r -> r[0])
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1])
                    .write(generateRows(rows, 3))
                    .writeTo(os);
        }

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        AtomicLong count = new AtomicLong();
        CsvReader.forMap()
                .read(new FileInputStream(file.toFile()), r -> count.incrementAndGet());

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();

        printResult("CSV Read 1M rows (MapReader)", rows, elapsed, Files.size(file), resources);
        assertEquals(rows, count.get());
    }

    @Test
    void csvRead_1m_rows_typedReader() throws IOException {
        int rows = 1_000_000;
        Path file = tempDir.resolve("read_typed_1m.csv");

        try (OutputStream os = Files.newOutputStream(file)) {
            CsvWriter.<int[]>create()
                    .column("ID", r -> r[0])
                    .column("Name", r -> "User-" + r[0])
                    .column("Score", r -> r[1])
                    .write(generateRows(rows, 3))
                    .writeTo(os);
        }

        probe = BenchmarkProbe.start(tempDir);
        long startTime = System.nanoTime();

        AtomicLong count = new AtomicLong();
        CsvReader.<String[]>mapping(row -> new String[]{
                row.get("ID").asString(),
                row.get("Name").asString(),
                row.get("Score").asString()
        }).read(new FileInputStream(file.toFile()), r -> count.incrementAndGet());

        long elapsed = System.nanoTime() - startTime;
        var resources = probe.stop();

        printResult("CSV Read 1M rows (typed)", rows, elapsed, Files.size(file), resources);
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

    private void writeExcelRows(Path file, int rows) throws IOException {
        try (OutputStream output = Files.newOutputStream(file)) {
            ExcelWriter.<int[]>create()
                    .column("ID", row -> row[0], c -> c.type(ExcelDataType.INTEGER))
                    .column("Name", row -> "User-" + row[0])
                    .column("Score", row -> row[1], c -> c.type(ExcelDataType.DOUBLE))
                    .write(generateRows(rows, 3))
                    .writeTo(output);
        }
    }

    private void printResult(String label, int rows, long elapsedNanos, long fileBytes,
                             BenchmarkProbe.Snapshot resources) throws IOException {
        double seconds = elapsedNanos / 1_000_000_000.0;
        System.out.printf(java.util.Locale.ROOT,
                "[BENCH] %s | %,d rows | %.3fs | %.2f MiB output | %,.0f rows/s | sampled heap peak %.2f MiB | temp peak %.2f MiB | temp remaining %d bytes%n",
                label, rows, seconds, fileBytes / 1048576.0, rows / seconds, resources.heapPeakBytes() / 1048576.0,
                resources.tempPeakBytes() / 1048576.0, resources.tempRemainingBytes());
        BenchmarkReport.append(scenario, rows, elapsedNanos, fileBytes, resources);
    }
}
