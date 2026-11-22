package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.testsupport.Damage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * An application that lets users bring their own template hands {@link ExcelTemplateWriter} whatever file they
 * uploaded. Starting from valid templates, this damages them in seeded random ways and fills them in: the only
 * acceptable outcomes are a finished workbook, one of the library's exceptions, or the {@link IOException} the
 * constructor declares. Not another kind of exception, not a hang. A failure names the seed and the damage.
 */
class TemplateInputFuzzTest {

    /** {@code -Dexcelkit.fuzz.scale=20} damages 20 times as many templates. */
    private static final int SCALE = Integer.getInteger("excelkit.fuzz.scale", 1);
    private static final int RAW = 150 * SCALE;
    private static final int XML = 150 * SCALE;
    private static final int SECONDS = 30;

    private record Item(String name, int qty) {
    }

    private static byte[] simple() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = wb.createSheet("Report");
            sheet.createRow(0).createCell(0).setCellValue("Title");
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 2));
            sheet.setColumnWidth(0, 20 * 256);
            wb.createSheet("Other").createRow(0).createCell(0).setCellValue("note");
            wb.write(out);
            return out.toByteArray();
        }
    }

    /** The template of {@link TemplatePreservationTest}: picture, chart, table, formatting, validation, names, panes. */
    private static byte[] rich() throws Exception {
        var method = TemplatePreservationTest.class.getDeclaredMethod("buildTemplate");
        method.setAccessible(true);
        return (byte[]) method.invoke(null);
    }

    private static void fillIn(byte[] template) throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B2", "value");
            writer.cell("A9", 12.5);
            writer.<Item>list(40).column("name", Item::name).column("qty", Item::qty)
                    .summary(s -> s.sum("qty"))
                    .write(List.of(new Item("a", 1), new Item("b", 2), new Item("c", 3)));
            writer.sheet(1);
            writer.cell("A3", "later");
            writer.finish().writeTo(OutputStream.nullOutputStream());
        }
    }

    private static final ExecutorService POOL = Executors.newCachedThreadPool(runnable -> {
        Thread thread = new Thread(runnable, "template-fuzz");
        thread.setDaemon(true);
        return thread;
    });

    private static int finished;
    private static int refused;

    /** Returns null when the outcome is acceptable, otherwise what went wrong. */
    private static String attempt(byte[] template) {
        Future<Throwable> future = POOL.submit(() -> {
            try {
                fillIn(template);
                return null;
            } catch (Throwable t) {
                return t;
            }
        });
        Throwable thrown;
        try {
            thrown = future.get(SECONDS, TimeUnit.SECONDS);
        } catch (TimeoutException e) {
            future.cancel(true);
            return "no result after " + SECONDS + " s (hang?)";
        } catch (Exception e) {
            return "harness: " + e;
        }
        if (thrown == null) {
            finished++;
            return null;
        }
        if (thrown instanceof ExcelKitException || thrown instanceof IOException) {
            refused++;
            return null;
        }
        StackTraceElement top = thrown.getStackTrace().length > 0 ? thrown.getStackTrace()[0] : null;
        // the first frame of ours says which call of the writer the exception came through
        StackTraceElement own = java.util.Arrays.stream(thrown.getStackTrace())
                .filter(f -> f.getClassName().startsWith("io.github.dornol.excelkit.excel")
                        && !f.getClassName().contains("Test")).findFirst().orElse(null);
        return thrown.getClass().getName() + ": " + String.valueOf(thrown.getMessage()).replace('\n', ' ')
                + (top == null ? "" : " @ " + top) + (own == null ? "" : " via " + own);
    }

    private void fuzz(String label, byte[] base) throws Exception {
        Map<String, List<String>> problems = new LinkedHashMap<>();
        finished = 0;
        refused = 0;
        Random r = new Random(label.hashCode() + 31L * SCALE);
        int attempts = 0;
        for (int i = 0; i < RAW + XML; i++) {
            byte[][] result = new byte[1][];
            String how = i < RAW ? Damage.raw(r, base, result) : Damage.insideArchive(r, base, result);
            attempts++;
            String problem = attempt(result[0]);
            if (problem != null) {
                problems.computeIfAbsent(problem.replaceAll("\\d+", "#"), x -> new ArrayList<>())
                        .add("[" + label + " #" + i + ": " + how + "] " + problem);
            }
        }
        StringBuilder report = new StringBuilder();
        problems.forEach((key, list) -> report.append("\n  ").append(list.size()).append("x ").append(list.get(0)));
        System.out.printf("TPLFUZZ %-7s %5d templates: %5d filled in, %5d refused, %d other%n", label, attempts, finished, refused,
                attempts - finished - refused);
        assertTrue(problems.isEmpty(), label + ": " + problems.size() + " kind(s) of unacceptable outcome in " + attempts + " templates:" + report);
        assertTrue(refused > attempts / 10, label + ": too few damaged templates were refused (" + refused + " of " + attempts + ")");
        assertTrue(finished > 0, label + ": every damaged template was refused, so the damage is not subtle enough");
    }

    @Test
    void aDamagedSimpleTemplateIsFilledInOrRefusedWithAnExpectedException() throws Exception {
        fuzz("simple", simple());
    }

    @Test
    void aDamagedTemplateWithPicturesChartsAndATableToo() throws Exception {
        fuzz("rich", rich());
    }
}
