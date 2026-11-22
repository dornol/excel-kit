package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.read.ReadAbortException;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.core.read.ReadSecurityPolicy;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelChartConfig;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.testsupport.Damage;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A reader is handed whatever a client uploads. Starting from valid files, this damages them in seeded random ways
 * (cut off, bits flipped, ranges overwritten, zeroed, duplicated or removed, and, for workbooks, the XML inside the
 * archive damaged and zipped again so that it reaches the parsers) and reads them. The only acceptable outcomes are
 * a normal read, whose rows may report errors, or one of the library's own exceptions: never another kind of
 * exception, a hang, or an {@link OutOfMemoryError}. A failure names the seed and the damage, so it can be repeated.
 */
class CorruptedInputFuzzTest {

    /** {@code -Dexcelkit.fuzz.scale=20} damages 20 times as many files, for a deeper one-off search. */
    private static final int SCALE = Integer.getInteger("excelkit.fuzz.scale", 1);
    private static final int RAW_MUTATIONS = 150 * SCALE;
    private static final int XML_MUTATIONS = 150 * SCALE;
    private static final int SECONDS_PER_READ = 30;

    private record Row(String name, int qty, double price, LocalDate day, LocalDateTime moment) {
    }

    private static List<Row> rows(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Row("name \"" + i + "\", é한", i, i * 1.5,
                LocalDate.of(2024, 1, 1).plusDays(i), LocalDateTime.of(2024, 1, 1, 0, 0).plusMinutes(i * 37L))).toList();
    }

    private static ExcelWriter<Row> columns(ExcelWriter<Row> writer) {
        return writer.column("Name", Row::name).column("Qty", Row::qty, c -> c.type(ExcelDataType.INTEGER))
                .column("Price", Row::price).column("Day", Row::day).column("Moment", Row::moment);
    }

    private static byte[] plain() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        columns(ExcelWriter.create()).write(rows(40)).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] rich() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        columns(ExcelWriter.<Row>create().maxRows(15).table("FuzzTable").autoFilter(true).freezeRows(1))
                .summary(s -> s.sum("Qty").average("Price"))
                .conditionalFormatting(cf -> cf.columns(1).greaterThan("5", io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_RED))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"))
                .write(rows(40)).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] encrypted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        columns(ExcelWriter.<Row>create().password("pw")).write(rows(25)).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] csv() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.<Row>create().column("Name", Row::name).column("Qty", Row::qty).column("Price", Row::price)
                .column("Day", Row::day).write(rows(40)).writeTo(out);
        return out.toByteArray();
    }

    // ---- reading -----------------------------------------------------------------------------------------

    private interface Read {
        void read(byte[] data) throws Exception;
    }

    private static final ExecutorService POOL = Executors.newCachedThreadPool(runnable -> {
        Thread thread = new Thread(runnable, "fuzz-reader");
        thread.setDaemon(true);
        return thread;
    });

    /** How the reads of the current run ended, so a test can tell that the damage did something. */
    private static final java.util.concurrent.atomic.AtomicInteger READ_OK = new java.util.concurrent.atomic.AtomicInteger();
    private static final java.util.concurrent.atomic.AtomicInteger READ_REFUSED = new java.util.concurrent.atomic.AtomicInteger();

    /** Returns null when the outcome is acceptable, otherwise what went wrong. */
    private static String attempt(Read read, byte[] data) {
        Future<Throwable> future = POOL.submit(() -> {
            try {
                read.read(data);
                return null;
            } catch (Throwable t) {
                return t;
            }
        });
        Throwable thrown;
        try {
            thrown = future.get(SECONDS_PER_READ, TimeUnit.SECONDS);
        } catch (TimeoutException e) {
            future.cancel(true);
            return "no result after " + SECONDS_PER_READ + " s (hang?)";
        } catch (Exception e) {
            return "harness: " + e;
        }
        if (thrown == null) {
            READ_OK.incrementAndGet();
            return null;
        }
        if (thrown instanceof ExcelKitException || thrown instanceof ReadAbortException) {
            READ_REFUSED.incrementAndGet();
            return null;
        }
        StackTraceElement top = thrown.getStackTrace().length > 0 ? thrown.getStackTrace()[0] : null;
        return thrown.getClass().getName() + ": " + String.valueOf(thrown.getMessage()).replace('\n', ' ')
                + (top == null ? "" : " @ " + top);
    }

    private static final Read EXCEL_DEFAULT = data -> ExcelReader.forMap().read(new ByteArrayInputStream(data), r -> { });
    private static final Read EXCEL_UPLOAD = data -> ExcelReader.forMap().limits(ReadLimits.UPLOAD)
            .securityPolicy(ReadSecurityPolicy.STRICT).read(new ByteArrayInputStream(data), r -> { });
    private static final Read EXCEL_INSPECT = data -> {
        var inspector = ExcelWorkbookInspector.create();
        var sheets = inspector.sheetNames(new ByteArrayInputStream(data));
        if (!sheets.isEmpty()) inspector.headers(new ByteArrayInputStream(data), 0, 1);
    };

    private void fuzz(String label, byte[] base, boolean zip, List<Read> readers) throws Exception {
        Map<String, List<String>> problems = new LinkedHashMap<>();
        READ_OK.set(0);
        READ_REFUSED.set(0);
        int attempts = 0;
        Random r = new Random(label.hashCode() + 31L * SCALE);
        for (int i = 0; i < RAW_MUTATIONS + (zip ? XML_MUTATIONS : 0); i++) {
            byte[][] result = new byte[1][];
            String how = i < RAW_MUTATIONS ? Damage.raw(r, base, result) : Damage.insideArchive(r, base, result);
            for (int k = 0; k < readers.size(); k++) {
                attempts++;
                String problem = attempt(readers.get(k), result[0]);
                if (problem != null) {
                    String key = problem.replaceAll("\\d+", "#");
                    problems.computeIfAbsent(key, x -> new ArrayList<>())
                            .add("[" + label + " #" + i + " reader " + k + ": " + how + "] " + problem);
                }
            }
        }
        StringBuilder report = new StringBuilder();
        problems.forEach((key, list) -> report.append("\n  ").append(list.size()).append("x ").append(list.get(0)));
        System.out.printf("FUZZ %-9s %5d reads: %5d read normally, %5d refused with a library exception, %d other%n",
                label, attempts, READ_OK.get(), READ_REFUSED.get(), attempts - READ_OK.get() - READ_REFUSED.get());
        assertTrue(problems.isEmpty(), label + ": " + problems.size() + " kind(s) of unacceptable outcome in "
                + attempts + " reads:" + report);
        // the damage must matter: some files are refused, and some still read (damage that lands in unused bytes)
        assertTrue(READ_REFUSED.get() > attempts / 10, label + ": too few damaged files were refused (" + READ_REFUSED.get() + " of " + attempts + ")");
        assertTrue(READ_OK.get() > 0, label + ": every damaged file was refused, so the damage is not subtle enough");
    }

    @Test
    void aDamagedPlainWorkbookEndsInAReadOrALibraryException() throws Exception {
        fuzz("plain", plain(), true, List.of(EXCEL_DEFAULT, EXCEL_UPLOAD, EXCEL_INSPECT));
    }

    @Test
    void aDamagedWorkbookWithTablesChartsAndSeveralSheetsToo() throws Exception {
        fuzz("rich", rich(), true, List.of(EXCEL_DEFAULT, EXCEL_UPLOAD, EXCEL_INSPECT));
    }

    @Test
    void aDamagedEncryptedWorkbook() throws Exception {
        Read withPassword = data -> ExcelReader.forMap().password("pw").read(new ByteArrayInputStream(data), r -> { });
        Read inspect = data -> ExcelWorkbookInspector.create().password("pw").sheetNames(new ByteArrayInputStream(data));
        fuzz("encrypted", encrypted(), false, List.of(withPassword, inspect));
    }

    @Test
    void aDamagedCsvFile() throws Exception {
        Read read = data -> CsvReader.forMap().read(new ByteArrayInputStream(data), r -> { });
        Read upload = data -> CsvReader.forMap().limits(ReadLimits.UPLOAD).read(new ByteArrayInputStream(data), r -> { });
        fuzz("csv", csv(), false, List.of(read, upload));
    }

    @Test
    void theHarnessFlagsWhatIsNotALibraryException() {
        byte[] any = {1, 2, 3};

        assertTrue(attempt(data -> { throw new NullPointerException("npe"); }, any).contains("NullPointerException"));
        assertTrue(attempt(data -> { throw new java.io.IOException("io"); }, any).contains("IOException"));
        assertTrue(attempt(data -> { throw new IllegalStateException("poi"); }, any).contains("IllegalStateException"));
        assertTrue(attempt(data -> { }, any) == null);
        assertTrue(attempt(data -> { throw new io.github.dornol.excelkit.excel.read.ExcelReadException("refused"); }, any) == null);
    }
}
