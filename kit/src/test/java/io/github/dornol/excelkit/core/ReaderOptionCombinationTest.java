package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Predicate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Every reader option is run on its own and together with every other one, against the same data as a CSV
 * and as an XLSX file, through every way of reading. The readers share their option handling, so for any
 * combination of options:
 * <ul>
 *   <li>the CSV and the XLSX file give the same rows, row numbers, errors and progress events;</li>
 *   <li>reading from a {@link Path} and from an {@link InputStream} agree;</li>
 *   <li>the summary, the report and {@code readWhile} agree with the rows that were delivered.</li>
 * </ul>
 * No expected values are written down per combination: the readers are checked against each other, so an
 * option that behaves differently in one reader, or only in one way of reading, shows up as a difference.
 */
class ReaderOptionCombinationTest {

    @TempDir
    Path tempDir;

    static class Row {
        String name;
        Integer score;
        String note;
    }

    /** The shape of the file and of the column mapping, which options change. */
    private static final class Spec {
        int preamble;
        boolean messyHeader;
        boolean duplicateHeader;
        boolean longCell;
        boolean aliasColumns;
        boolean indexColumns;
        boolean requiredNote;
    }

    /** State shared with the options of one read. */
    private static final class Run {
        final AtomicInteger delivered = new AtomicInteger();
        final List<String> progress = new ArrayList<>();
    }

    private record Feature(String name, Consumer<Spec> shape,
                           BiConsumer<AbstractReader<Row, ?>, Run> configure) {
        static Feature option(String name, BiConsumer<AbstractReader<Row, ?>, Run> configure) {
            return new Feature(name, spec -> { }, configure);
        }

        static Feature file(String name, Consumer<Spec> shape, BiConsumer<AbstractReader<Row, ?>, Run> configure) {
            return new Feature(name, shape, configure);
        }
    }

    private static final List<Feature> FEATURES = List.of(
            Feature.file("preamble", s -> s.preamble = 2, (r, run) -> r.headerRowIndex(2)),
            Feature.file("headerPolicy", s -> s.messyHeader = true,
                    (r, run) -> r.headerPolicy(HeaderPolicy.TRIM_CASE_INSENSITIVE)),
            Feature.file("duplicateHeader.FIRST", s -> s.duplicateHeader = true,
                    (r, run) -> r.duplicateHeaderPolicy(DuplicateHeaderPolicy.FIRST)),
            Feature.file("aliasColumns", s -> s.aliasColumns = true, (r, run) -> { }),
            Feature.file("indexColumns", s -> s.indexColumns = true, (r, run) -> { }),
            Feature.file("requiredNote", s -> s.requiredNote = true, (r, run) -> { }),
            Feature.option("skipBlankRows", (r, run) -> r.skipBlankRows()),
            Feature.option("stopAtBlankRows", (r, run) -> r.stopAtBlankRows(2)),
            Feature.option("maxRows", (r, run) -> r.maxRows(3)),
            Feature.option("maxErrors", (r, run) -> r.maxErrors(0)),
            Feature.option("strictHeaders", (r, run) -> r.strictHeaders()),
            Feature.file("limits", s -> s.longCell = true, (r, run) -> r.limits(new ReadLimits(-1, -1, -1, 12, -1))),
            Feature.option("cancel", (r, run) -> r.cancellationToken(() -> run.delivered.get() >= 2)),
            Feature.option("onProgress", (r, run) -> r.onProgress(2, (rows, cursor) -> run.progress.add("rows=" + rows))),
            Feature.option("onReadProgress", (r, run) -> r.onReadProgress(2, p ->
                    run.progress.add("read=" + p.processedRows() + "/" + p.successRows() + "/" + p.errorRows()
                            + (p.completed() ? "/completed" : "") + (p.cancelled() ? "/cancelled" : "")))));

    // ---- the data

    /** The logical rows below the header: {@code null} is a blank row. */
    private static List<String[]> body(Spec spec) {
        List<String[]> rows = new ArrayList<>(Arrays.asList(
                new String[]{"alice", "10", "x"},
                new String[]{"bob", "20", ""},
                null,
                new String[]{"carol", "abc", "y"},
                new String[]{"dave", "40", spec.longCell ? "a-note-longer-than-twelve" : "z"},
                null,
                null,
                new String[]{"erin", "50", "w"}));
        return rows;
    }

    private static String[] header(Spec spec) {
        if (spec.duplicateHeader) return new String[]{"Name", "Name", "Score", "Note"};
        if (spec.messyHeader) return new String[]{" NAME ", "score", "Note "};
        if (spec.aliasColumns) return new String[]{"Full name", "Score", "Note"};
        return new String[]{"Name", "Score", "Note"};
    }

    /** All rows of the file, header and junk included; {@code null} is a blank row. */
    private static List<String[]> fileRows(Spec spec) {
        List<String[]> rows = new ArrayList<>();
        for (int i = 0; i < spec.preamble; i++) rows.add(new String[]{"junk" + i});
        String[] header = header(spec);
        rows.add(header);
        for (String[] row : body(spec)) {
            if (row != null && spec.duplicateHeader) {
                rows.add(new String[]{row[0], "dup", row[1], row[2]});
            } else {
                rows.add(row);
            }
        }
        return rows;
    }

    private Path csvFile(Spec spec) throws IOException {
        StringBuilder text = new StringBuilder();
        for (String[] row : fileRows(spec)) text.append(row == null ? "" : String.join(",", row)).append('\n');
        Path file = Files.createTempFile(tempDir, "data", ".csv");
        Files.writeString(file, text.toString(), StandardCharsets.UTF_8);
        return file;
    }

    private Path xlsxFile(Spec spec) throws IOException {
        Path file = Files.createTempFile(tempDir, "data", ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file)) {
            XSSFSheet sheet = workbook.createSheet("S");
            List<String[]> rows = fileRows(spec);
            for (int r = 0; r < rows.size(); r++) {
                String[] row = rows.get(r);
                var target = sheet.createRow(r);
                if (row == null) { // a blank row of a CSV is a row without values here, as Excel keeps blank rows
                    target.createCell(0).setBlank();
                    continue;
                }
                for (int c = 0; c < row.length; c++) {
                    if (!row[c].isEmpty()) target.createCell(c).setCellValue(row[c]);
                }
            }
            workbook.write(out);
        }
        return file;
    }

    // ---- readers

    /** One format's reader, ready to read in each way. */
    private interface Format {
        String name();

        Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException;

        Outcome readPath(Path file, Spec spec, Run run, Feature[] features);

        Outcome readStream(Path file, Spec spec, Run run, Feature[] features);

        Outcome summary(Path file, Spec spec, Run run, Feature[] features);

        Outcome report(Path file, Spec spec, Run run, Feature[] features);

        Outcome whileAll(Path file, Spec spec, Run run, Feature[] features);
    }

    /** What a read produced, in a form that does not depend on the file format. */
    private record Outcome(List<String> rows, String failure, long total, long success, long errors,
                           Boolean stoppedEarly, List<String> progress) {
    }

    private static <R extends AbstractReader<Row, R>> R configured(R reader, Spec spec, Run run, Feature[] features) {
        if (spec.indexColumns) {
            reader.columnAt(0 + (spec.duplicateHeader ? 0 : 0), (row, cell) -> row.name = cell.asString());
            int shift = spec.duplicateHeader ? 2 : 1;
            reader.columnAt(shift, (row, cell) -> row.score = cell.asInt());
            reader.columnAt(shift + 1, (row, cell) -> row.note = cell.asString());
        } else {
            if (spec.aliasColumns) {
                reader.column(List.of("Nope", "Full name"), (row, cell) -> row.name = cell.asString());
            } else {
                reader.column("Name", (row, cell) -> row.name = cell.asString());
            }
            reader.column("Score", (row, cell) -> row.score = cell.asInt());
            reader.column("Note", (row, cell) -> row.note = cell.asString());
        }
        if (spec.requiredNote) reader.required();
        for (Feature feature : features) feature.configure().accept(reader, run);
        return reader;
    }

    private static String describe(ReadResult<Row> result) {
        String data = result.data() == null ? "-"
                : result.data().name + "|" + result.data().score + "|" + result.data().note;
        // An XLSX has no cell for a trailing empty value; a CSV has an empty field. Not a difference in meaning.
        List<String> raw = new ArrayList<>(result.rawValues());
        while (!raw.isEmpty() && (raw.get(raw.size() - 1) == null || raw.get(raw.size() - 1).isEmpty())) {
            raw.remove(raw.size() - 1);
        }
        return (result.success() ? "ok " : "err " + result.errorType() + " ") + data
                + " row=" + result.fileRowNum() + " raw=" + raw;
    }

    /** Failures compare by kind: the message and exception class of the two formats differ by design. */
    private static String kind(Throwable failure) {
        if (failure == null) return null;
        if (failure instanceof ReadLimitExceededException limit) return "limit:" + limit.limit();
        if (failure instanceof ReadAbortException) return "abort";
        if (failure instanceof ReadSecurityException) return "security";
        Throwable cause = failure;
        while (cause.getCause() != null) cause = cause.getCause();
        return "read-failed:" + cause.getClass().getSimpleName();
    }

    private static final class Collector {
        final List<String> rows = new ArrayList<>();

        void accept(ReadResult<Row> result, Run run) {
            rows.add(describe(result));
            run.delivered.incrementAndGet();
        }
    }

    private static Outcome outcome(Collector collector, Throwable failure, ReadSummary summary, Run run) {
        long total = summary != null ? summary.totalRows() : collector.rows.size();
        long success = summary != null ? summary.successRows() : collector.rows.stream().filter(r -> r.startsWith("ok")).count();
        long errors = summary != null ? summary.errorRows() : collector.rows.stream().filter(r -> r.startsWith("err")).count();
        return new Outcome(collector.rows, kind(failure), total, success, errors,
                summary == null ? null : summary.stoppedEarly(), List.copyOf(run.progress));
    }

    private record CsvFormat() implements Format {
        public String name() { return "CSV"; }

        public Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException { return test.csvFile(spec); }

        private CsvReader<Row> reader(Spec spec, Run run, Feature[] features) {
            return configured(CsvReader.setter(Row::new), spec, run, features);
        }

        public Outcome readPath(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            try { reader(spec, run, features).read(file, r -> collector.accept(r, run)); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, null, run);
        }

        public Outcome readStream(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            try (InputStream in = Files.newInputStream(file)) {
                reader(spec, run, features).read(in, r -> collector.accept(r, run));
            } catch (RuntimeException | IOException e) { failure = e; }
            return outcome(collector, failure, null, run);
        }

        public Outcome summary(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            ReadSummary summary = null;
            try { summary = reader(spec, run, features).readWithSummary(file, r -> collector.accept(r, run)); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, summary, run);
        }

        public Outcome report(Path file, Spec spec, Run run, Feature[] features) {
            Throwable failure = null;
            ReadReport report = null;
            try { report = reader(spec, run, features).readReport(file, 100); } catch (RuntimeException e) { failure = e; }
            List<String> errors = new ArrayList<>();
            if (report != null) report.errors().forEach(e -> errors.add("err " + e.type() + " row=" + e.fileRowNum()));
            return new Outcome(errors, kind(failure), report == null ? 0 : report.summary().totalRows(),
                    report == null ? 0 : report.summary().successRows(), report == null ? 0 : report.summary().errorRows(),
                    report == null ? null : report.summary().stoppedEarly(), List.copyOf(run.progress));
        }

        public Outcome whileAll(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            ReadSummary summary = null;
            Predicate<ReadResult<Row>> all = r -> { collector.accept(r, run); return true; };
            try { summary = reader(spec, run, features).readWhile(file, all); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, summary, run);
        }
    }

    private record ExcelFormat() implements Format {
        public String name() { return "XLSX"; }

        public Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException { return test.xlsxFile(spec); }

        private ExcelReader<Row> reader(Spec spec, Run run, Feature[] features) {
            return configured(ExcelReader.setter(Row::new), spec, run, features);
        }

        public Outcome readPath(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            try { reader(spec, run, features).read(file, r -> collector.accept(r, run)); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, null, run);
        }

        public Outcome readStream(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            try (InputStream in = Files.newInputStream(file)) {
                reader(spec, run, features).read(in, r -> collector.accept(r, run));
            } catch (RuntimeException | IOException e) { failure = e; }
            return outcome(collector, failure, null, run);
        }

        public Outcome summary(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            ReadSummary summary = null;
            try { summary = reader(spec, run, features).readWithSummary(file, r -> collector.accept(r, run)); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, summary, run);
        }

        public Outcome report(Path file, Spec spec, Run run, Feature[] features) {
            Throwable failure = null;
            ReadReport report = null;
            try { report = reader(spec, run, features).readReport(file, 100); } catch (RuntimeException e) { failure = e; }
            List<String> errors = new ArrayList<>();
            if (report != null) report.errors().forEach(e -> errors.add("err " + e.type() + " row=" + e.fileRowNum()));
            return new Outcome(errors, kind(failure), report == null ? 0 : report.summary().totalRows(),
                    report == null ? 0 : report.summary().successRows(), report == null ? 0 : report.summary().errorRows(),
                    report == null ? null : report.summary().stoppedEarly(), List.copyOf(run.progress));
        }

        public Outcome whileAll(Path file, Spec spec, Run run, Feature[] features) {
            Collector collector = new Collector();
            Throwable failure = null;
            ReadSummary summary = null;
            Predicate<ReadResult<Row>> all = r -> { collector.accept(r, run); return true; };
            try { summary = reader(spec, run, features).readWhile(file, all); } catch (RuntimeException e) { failure = e; }
            return outcome(collector, failure, summary, run);
        }
    }

    // ---- the test

    @Test
    void everyOptionAgreesAcrossFormatsAndWaysOfReading() throws Exception {
        List<String> problems = new ArrayList<>();
        int combinations = 0;
        for (int i = 0; i < FEATURES.size(); i++) {
            check(problems, FEATURES.get(i));
            combinations++;
            for (int j = i + 1; j < FEATURES.size(); j++) {
                check(problems, FEATURES.get(i), FEATURES.get(j));
                combinations++;
            }
        }
        check(problems, FEATURES.toArray(Feature[]::new));
        check(problems);
        assertTrue(combinations > 100, "too few combinations: " + combinations);
        assertTrue(problems.isEmpty(), problems.size() + " reader option problem(s):\n  " + String.join("\n  ", problems));
    }

    /**
     * The comparison above only says the readers agree; these say what they agree on, so that two readers
     * that both deliver nothing (or both ignore an option) cannot pass it.
     */
    @Test
    void knownOptionsGiveTheExpectedRows() throws Exception {
        Spec spec = new Spec();
        for (Format format : List.of(new CsvFormat(), new ExcelFormat())) {
            Path file = format.file(this, spec);

            Outcome plain = format.readPath(file, spec, new Run(), new Feature[0]);
            assertEquals(8, plain.rows().size(), format.name() + " delivers blank rows too: " + plain.rows());
            assertEquals(1, plain.rows().stream().filter(r -> r.startsWith("err MAPPING carol")).count(), format.name());

            Feature skip = FEATURES.stream().filter(f -> f.name().equals("skipBlankRows")).findFirst().orElseThrow();
            Feature max = FEATURES.stream().filter(f -> f.name().equals("maxRows")).findFirst().orElseThrow();
            Outcome limited = format.readPath(file, spec, new Run(), new Feature[]{skip, max});
            assertEquals(3, limited.rows().size(), format.name() + ": " + limited.rows());
            assertTrue(limited.rows().get(0).startsWith("ok alice"), limited.rows().get(0));
            assertTrue(limited.rows().get(2).startsWith("err MAPPING carol"), limited.rows().get(2));

            Feature strict = FEATURES.stream().filter(f -> f.name().equals("maxErrors")).findFirst().orElseThrow();
            Outcome aborted = format.readPath(file, spec, new Run(), new Feature[]{strict});
            assertEquals("abort", aborted.failure(), format.name());
        }
    }

    private void check(List<String> problems, Feature... features) throws IOException {
        String label = features.length == 0 ? "(no options)"
                : String.join(" + ", Arrays.stream(features).map(Feature::name).toList());
        Spec spec = new Spec();
        for (Feature feature : features) feature.shape().accept(spec);
        boolean cancelled = Arrays.stream(features).anyMatch(f -> f.name().equals("cancel"));

        List<Format> formats = List.of(new CsvFormat(), new ExcelFormat());
        Outcome[] byFormat = new Outcome[formats.size()];
        for (int f = 0; f < formats.size(); f++) {
            Format format = formats.get(f);
            Path file = format.file(this, spec);
            Outcome rows = format.readPath(file, spec, new Run(), features);
            byFormat[f] = rows;

            same(problems, label, format.name() + ": a Path and an InputStream", rows,
                    format.readStream(file, spec, new Run(), features));
            Outcome summary = format.summary(file, spec, new Run(), features);
            same(problems, label, format.name() + ": read and readWithSummary rows", rows.rows(), summary.rows());
            long deliveredOk = rows.rows().stream().filter(r -> r.startsWith("ok")).count();
            same(problems, label, format.name() + ": the summary counts the delivered rows",
                    List.of((long) rows.rows().size(), deliveredOk),
                    summary.failure() != null ? List.of((long) rows.rows().size(), deliveredOk)
                            : List.of(summary.total(), summary.success()));
            Outcome whileAll = format.whileAll(file, spec, new Run(), features);
            same(problems, label, format.name() + ": readWhile(always) and readWithSummary", summary, whileAll);
            if (!cancelled) {
                Outcome report = format.report(file, spec, new Run(), features);
                // A failed read has no counts of its own (the rows delivered before it are not a summary).
                boolean failed = summary.failure() != null || report.failure() != null;
                same(problems, label, format.name() + ": readReport and readWithSummary counts",
                        failed ? List.of(String.valueOf(summary.failure()))
                                : List.of(summary.total(), summary.success(), summary.errors(), String.valueOf(summary.failure())),
                        failed ? List.of(String.valueOf(report.failure()))
                                : List.of(report.total(), report.success(), report.errors(), String.valueOf(report.failure())));
            }
        }
        same(problems, label, "CSV and XLSX", byFormat[0], byFormat[1]);
    }

    private static void same(List<String> problems, String label, String what, Object expected, Object actual) {
        if (!expected.equals(actual)) {
            problems.add("[" + label + "] " + what + " differ:\n      " + expected + "\n      " + actual);
        }
    }
}
