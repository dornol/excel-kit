package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.ss.usermodel.CellStyle;
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
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Every reader option is run on its own and together with every other one, against the same data as a CSV
 * file, as an XLSX file and as the third sheet of an XLSX workbook, in the three ways of mapping rows
 * (setters, a row mapper, maps), through every way of reading. The readers share their option handling, so
 * for any combination of options:
 * <ul>
 *   <li>the three files give the same rows, row numbers, errors and progress events;</li>
 *   <li>reading from a {@link Path} and from an {@link InputStream} agree;</li>
 *   <li>the summary, the report and {@code readWhile} agree with the rows that were delivered.</li>
 * </ul>
 * No expected values are written down per combination: the readers are checked against each other, so an
 * option that behaves differently in one reader, in one mapping mode or in one way of reading shows up as a
 * difference. {@link #knownOptionsGiveTheExpectedRows()} pins down what they agree on.
 */
class ReaderOptionCombinationTest {

    @TempDir
    Path tempDir;

    static class Row {
        String name;
        Integer score;
        String note;
        LocalDate when;
        Double price;
        LocalDate day;
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
        /** Numbers and dates are number and date cells in the XLSX, text in the CSV. */
        boolean typedCells;
        /** Text a locale and a date pattern are needed to read: {@code 1.234,50} and {@code 05/03/2024}. */
        boolean conversion;
    }

    /** State shared with the options of one read. */
    private static final class Run {
        final AtomicInteger delivered = new AtomicInteger();
        final List<String> progress = new ArrayList<>();
    }

    private enum Mode { SETTER, MAPPING, MAP }

    private record Feature(String name, Consumer<Spec> shape, BiConsumer<AbstractReader<?, ?>, Run> configure) {
        static Feature option(String name, BiConsumer<AbstractReader<?, ?>, Run> configure) {
            return new Feature(name, spec -> { }, configure);
        }

        static Feature file(String name, Consumer<Spec> shape, BiConsumer<AbstractReader<?, ?>, Run> configure) {
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
            Feature.file("typedCells", s -> s.typedCells = true, (r, run) -> { }),
            Feature.file("cellConversion", s -> s.conversion = true,
                    (r, run) -> r.cellConversion(c -> c.locale(Locale.GERMANY).addDateFormat("dd/MM/yyyy"))),
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

    /** The logical rows below the header: {@code null} is a blank row; cells are Strings, or typed when asked. */
    private static List<Object[]> body(Spec spec) {
        return new ArrayList<>(Arrays.asList(
                row(spec, 0, "alice", 10, "x"),
                row(spec, 1, "bob", 20, ""),
                null,
                row(spec, 3, "carol", "abc", "y"),
                row(spec, 4, "dave", 40, spec.longCell ? "a-note-longer-than-twelve" : "z"),
                null,
                null,
                row(spec, 7, "erin", 50, "w")));
    }

    private static Object[] row(Spec spec, int index, String name, Object score, String note) {
        List<Object> cells = new ArrayList<>();
        cells.add(name);
        if (spec.duplicateHeader) cells.add("dup");
        cells.add(spec.typedCells ? score : String.valueOf(score));
        cells.add(note);
        if (spec.typedCells) cells.add(LocalDate.of(2024, 3, 1).plusDays(index));
        else cells.add(LocalDate.of(2024, 3, 1).plusDays(index).toString());
        if (spec.conversion) {
            cells.add("1.234,50");
            cells.add("05/03/2024");
        }
        return cells.toArray();
    }

    private static String[] header(Spec spec) {
        List<String> names = new ArrayList<>();
        names.add(spec.messyHeader ? " NAME " : spec.aliasColumns ? "Full name" : "Name");
        if (spec.duplicateHeader) names.add("Name");
        names.add(spec.messyHeader ? "score" : "Score");
        names.add(spec.messyHeader ? "Note " : "Note");
        names.add("When");
        if (spec.conversion) {
            names.add("Price");
            names.add("Day");
        }
        return names.toArray(String[]::new);
    }

    /** All rows of the file, header and junk included; {@code null} is a blank row. */
    private static List<Object[]> fileRows(Spec spec) {
        List<Object[]> rows = new ArrayList<>();
        for (int i = 0; i < spec.preamble; i++) rows.add(new Object[]{"junk" + i});
        rows.add(header(spec));
        rows.addAll(body(spec));
        return rows;
    }

    private Path csvFile(Spec spec) throws IOException {
        StringBuilder text = new StringBuilder();
        for (Object[] row : fileRows(spec)) {
            text.append(row == null ? "" : String.join(",", Arrays.stream(row)
                    .map(cell -> String.valueOf(cell)).map(cell -> cell.contains(",") ? "\"" + cell + "\"" : cell)
                    .toList())).append('\n');
        }
        Path file = Files.createTempFile(tempDir, "data", ".csv");
        Files.writeString(file, text.toString(), StandardCharsets.UTF_8);
        return file;
    }

    private static void fill(XSSFSheet sheet, CellStyle dateStyle, List<Object[]> rows) {
        for (int r = 0; r < rows.size(); r++) {
            Object[] row = rows.get(r);
            var target = sheet.createRow(r);
            if (row == null) { // a blank row of a CSV is a row without values here, as Excel keeps blank rows
                target.createCell(0).setBlank();
                continue;
            }
            for (int c = 0; c < row.length; c++) {
                Object value = row[c];
                if (value instanceof String text) {
                    if (!text.isEmpty()) target.createCell(c).setCellValue(text);
                } else if (value instanceof Integer number) {
                    target.createCell(c).setCellValue(number);
                } else if (value instanceof LocalDate date) {
                    var cell = target.createCell(c);
                    cell.setCellValue(date);
                    cell.setCellStyle(dateStyle);
                }
            }
        }
    }

    /** The data as the first sheet, or as the third of three when {@code thirdOfThree} (the others are decoys). */
    private Path xlsxFile(Spec spec, boolean thirdOfThree) throws IOException {
        Path file = Files.createTempFile(tempDir, "data", ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file)) {
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            if (thirdOfThree) {
                for (String decoy : List.of("decoy1", "decoy2")) {
                    fill(workbook.createSheet(decoy), dateStyle, List.of(
                            new Object[]{"Other", "Columns"}, new Object[]{"zed", "9"}, new Object[]{"yan", "8"}));
                }
            }
            fill(workbook.createSheet("S"), dateStyle, fileRows(spec));
            workbook.write(out);
        }
        return file;
    }

    // ---- readers, in each way of reading

    private interface Reading<T> {
        void read(Path file, Consumer<ReadResult<T>> consumer);

        void read(InputStream in, Consumer<ReadResult<T>> consumer);

        ReadSummary summary(Path file, Consumer<ReadResult<T>> consumer);

        ReadReport report(Path file, int maxErrors);

        ReadSummary whileAll(Path file, Predicate<ReadResult<T>> predicate);
    }

    private record CsvReading<T>(CsvReader<T> reader) implements Reading<T> {
        public void read(Path file, Consumer<ReadResult<T>> consumer) { reader.read(file, consumer); }
        public void read(InputStream in, Consumer<ReadResult<T>> consumer) { reader.read(in, consumer); }
        public ReadSummary summary(Path file, Consumer<ReadResult<T>> consumer) { return reader.readWithSummary(file, consumer); }
        public ReadReport report(Path file, int maxErrors) { return reader.readReport(file, maxErrors); }
        public ReadSummary whileAll(Path file, Predicate<ReadResult<T>> predicate) { return reader.readWhile(file, predicate); }
    }

    private record ExcelReading<T>(ExcelReader<T> reader) implements Reading<T> {
        public void read(Path file, Consumer<ReadResult<T>> consumer) { reader.read(file, consumer); }
        public void read(InputStream in, Consumer<ReadResult<T>> consumer) { reader.read(in, consumer); }
        public ReadSummary summary(Path file, Consumer<ReadResult<T>> consumer) { return reader.readWithSummary(file, consumer); }
        public ReadReport report(Path file, int maxErrors) { return reader.readReport(file, maxErrors); }
        public ReadSummary whileAll(Path file, Predicate<ReadResult<T>> predicate) { return reader.readWhile(file, predicate); }
    }

    /** One file format: how to write the data, and how to read it back in each mapping mode. */
    private interface Variant {
        String name();

        Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException;

        Supplier<Reading<Row>> rows(Mode mode, Spec spec, Run run, Feature[] features);

        Supplier<Reading<Map<String, String>>> maps(Run run, Feature[] features);
    }

    private static <R extends AbstractReader<?, R>> R apply(R reader, Run run, Feature[] features) {
        for (Feature feature : features) feature.configure().accept(reader, run);
        return reader;
    }

    /** The setter-mode column mapping of the {@link Row}, as the options ask for it. */
    private static <R extends AbstractReader<Row, R>> R columns(R reader, Spec spec) {
        int shift = spec.duplicateHeader ? 1 : 0;
        if (spec.indexColumns) {
            reader.columnAt(0, (row, cell) -> row.name = cell.asString());
            reader.columnAt(1 + shift, (row, cell) -> row.score = cell.asInt());
            reader.columnAt(2 + shift, (row, cell) -> row.note = cell.asString());
            if (spec.requiredNote) reader.required();
            reader.columnAt(3 + shift, (row, cell) -> row.when = cell.asLocalDate());
            if (spec.conversion) {
                reader.columnAt(4 + shift, (row, cell) -> row.price = cell.asDouble());
                reader.columnAt(5 + shift, (row, cell) -> row.day = cell.asLocalDate());
            }
            return reader;
        }
        if (spec.aliasColumns) {
            reader.column(List.of("Nope", "Full name"), (row, cell) -> row.name = cell.asString());
        } else {
            reader.column("Name", (row, cell) -> row.name = cell.asString());
        }
        reader.column("Score", (row, cell) -> row.score = cell.asInt());
        reader.column("Note", (row, cell) -> row.note = cell.asString());
        if (spec.requiredNote) reader.required();
        reader.column("When", (row, cell) -> row.when = cell.asLocalDate());
        if (spec.conversion) {
            reader.column("Price", (row, cell) -> row.price = cell.asDouble());
            reader.column("Day", (row, cell) -> row.day = cell.asLocalDate());
        }
        return reader;
    }

    private static Row mapRow(RowData row, Spec spec) {
        Row mapped = new Row();
        mapped.name = row.get("Name").asString();
        mapped.score = row.get("Score").asInt();
        mapped.note = row.get("Note").asString();
        mapped.when = row.get("When").asLocalDate();
        if (spec.conversion) {
            mapped.price = row.get("Price").asDouble();
            mapped.day = row.get("Day").asLocalDate();
        }
        return mapped;
    }

    private record CsvVariant() implements Variant {
        public String name() { return "CSV"; }

        public Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException { return test.csvFile(spec); }

        public Supplier<Reading<Row>> rows(Mode mode, Spec spec, Run run, Feature[] features) {
            return () -> new CsvReading<>(apply(mode == Mode.SETTER
                    ? columns(CsvReader.setter(Row::new), spec)
                    : CsvReader.mapping(row -> mapRow(row, spec)), run, features));
        }

        public Supplier<Reading<Map<String, String>>> maps(Run run, Feature[] features) {
            return () -> new CsvReading<>(apply(CsvReader.forMap(), run, features));
        }
    }

    private record ExcelVariant(boolean thirdOfThree) implements Variant {
        public String name() { return thirdOfThree ? "XLSX (third sheet)" : "XLSX"; }

        public Path file(ReaderOptionCombinationTest test, Spec spec) throws IOException {
            return test.xlsxFile(spec, thirdOfThree);
        }

        private <R extends AbstractReader<?, R>> R sheet(ExcelReader<?> reader, R configured) {
            return configured;
        }

        public Supplier<Reading<Row>> rows(Mode mode, Spec spec, Run run, Feature[] features) {
            return () -> {
                ExcelReader<Row> reader = mode == Mode.SETTER
                        ? columns(ExcelReader.setter(Row::new), spec)
                        : ExcelReader.mapping(row -> mapRow(row, spec));
                if (thirdOfThree) reader.sheetIndex(2);
                return new ExcelReading<>(apply(reader, run, features));
            };
        }

        public Supplier<Reading<Map<String, String>>> maps(Run run, Feature[] features) {
            return () -> {
                ExcelReader<Map<String, String>> reader = ExcelReader.forMap();
                if (thirdOfThree) reader.sheetIndex(2);
                return new ExcelReading<>(apply(reader, run, features));
            };
        }
    }

    // ---- what a read produced

    /** What a read produced, in a form that does not depend on the file format. */
    private record Outcome(List<String> rows, String failure, long total, long success, long errors,
                           Boolean stoppedEarly, List<String> progress) {
    }

    /** Failures compare by kind: the message and exception class of the formats differ by design. */
    private static String kind(Throwable failure) {
        if (failure == null) return null;
        if (failure instanceof ReadLimitExceededException limit) return "limit:" + limit.limit();
        if (failure instanceof ReadAbortException) return "abort";
        if (failure instanceof ReadSecurityException) return "security";
        Throwable cause = failure;
        while (cause.getCause() != null) cause = cause.getCause();
        return "read-failed:" + cause.getClass().getSimpleName();
    }

    private static <T> String describe(ReadResult<T> result, Function<T, String> data, boolean raw) {
        // An XLSX has no cell for a trailing empty value; a CSV has an empty field. Not a difference in meaning.
        List<String> values = new ArrayList<>(result.rawValues());
        while (!values.isEmpty() && (values.get(values.size() - 1) == null || values.get(values.size() - 1).isEmpty())) {
            values.remove(values.size() - 1);
        }
        return (result.success() ? "ok " : "err " + result.errorType() + " ")
                + (result.data() == null ? "-" : data.apply(result.data()))
                + " row=" + result.fileRowNum() + (raw ? " raw=" + values : "");
    }

    private static String describeRow(Row row) {
        return row.name + "|" + row.score + "|" + row.note + "|" + row.when + "|" + row.price + "|" + row.day;
    }

    /** One variant read in each way, for one mapping mode. */
    private record Terminals<T>(Supplier<Reading<T>> reading, Function<T, String> data, boolean raw, Run run) {
        private Outcome result(List<String> rows, Throwable failure, ReadSummary summary) {
            long ok = rows.stream().filter(r -> r.startsWith("ok")).count();
            return new Outcome(rows, kind(failure), summary != null ? summary.totalRows() : rows.size(),
                    summary != null ? summary.successRows() : ok, summary != null ? summary.errorRows() : rows.size() - ok,
                    summary == null ? null : summary.stoppedEarly(), List.copyOf(run.progress));
        }

        private Consumer<ReadResult<T>> collect(List<String> rows) {
            return r -> {
                rows.add(describe(r, data, raw));
                run.delivered.incrementAndGet();
            };
        }

        Outcome path(Path file) {
            List<String> rows = new ArrayList<>();
            Throwable failure = null;
            try { reading.get().read(file, collect(rows)); } catch (RuntimeException e) { failure = e; }
            return result(rows, failure, null);
        }

        Outcome stream(Path file) {
            List<String> rows = new ArrayList<>();
            Throwable failure = null;
            try (InputStream in = Files.newInputStream(file)) {
                reading.get().read(in, collect(rows));
            } catch (RuntimeException | IOException e) { failure = e; }
            return result(rows, failure, null);
        }

        Outcome summary(Path file) {
            List<String> rows = new ArrayList<>();
            Throwable failure = null;
            ReadSummary summary = null;
            try { summary = reading.get().summary(file, collect(rows)); } catch (RuntimeException e) { failure = e; }
            return result(rows, failure, summary);
        }

        Outcome whileAll(Path file) {
            List<String> rows = new ArrayList<>();
            Throwable failure = null;
            ReadSummary summary = null;
            Consumer<ReadResult<T>> collect = collect(rows);
            try { summary = reading.get().whileAll(file, r -> { collect.accept(r); return true; }); } catch (RuntimeException e) { failure = e; }
            return result(rows, failure, summary);
        }

        Outcome report(Path file) {
            Throwable failure = null;
            ReadReport report = null;
            try { report = reading.get().report(file, 100); } catch (RuntimeException e) { failure = e; }
            ReadSummary summary = report == null ? null : report.summary();
            return new Outcome(List.of(), kind(failure), summary == null ? 0 : summary.totalRows(),
                    summary == null ? 0 : summary.successRows(), summary == null ? 0 : summary.errorRows(),
                    summary == null ? null : summary.stoppedEarly(), List.copyOf(run.progress));
        }
    }

    // ---- the tests

    @Test
    void everyOptionAgreesAcrossFormatsModesAndWaysOfReading() throws Exception {
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
     * The comparison above only says the readers agree; these say what they agree on, so that readers that
     * both deliver nothing (or both ignore an option) cannot pass it.
     */
    @Test
    void knownOptionsGiveTheExpectedRows() throws Exception {
        Spec spec = new Spec();
        Feature skip = feature("skipBlankRows");
        Feature max = feature("maxRows");
        Feature strict = feature("maxErrors");
        for (Variant variant : variants()) {
            Path file = variant.file(this, spec);
            for (Mode mode : List.of(Mode.SETTER, Mode.MAPPING)) {
                String where = variant.name() + " " + mode;
                Outcome plain = terminals(variant, mode, spec, new Run(), new Feature[0]).path(file);
                assertEquals(8, plain.rows().size(), where + " delivers blank rows too: " + plain.rows());
                // A row mapper that fails leaves no instance behind, so only the setter mode names the row.
                assertEquals(1, plain.rows().stream().filter(r -> r.startsWith("err MAPPING")).count(), where + plain.rows());
                assertTrue(plain.rows().stream().filter(r -> r.startsWith("err")).allMatch(r -> r.contains("row=5")), where);

                Outcome limited = terminals(variant, mode, spec, new Run(), new Feature[]{skip, max}).path(file);
                assertEquals(3, limited.rows().size(), where + ": " + limited.rows());
                assertTrue(limited.rows().get(0).startsWith("ok alice"), limited.rows().get(0));
                assertTrue(limited.rows().get(2).startsWith("err MAPPING"), limited.rows().get(2));

                assertEquals("abort", terminals(variant, mode, spec, new Run(), new Feature[]{strict}).path(file).failure(), where);
            }
        }
    }

    @Test
    void mapModeDeliversTheHeaderNamesAsKeys() throws Exception {
        Spec spec = new Spec();
        for (Variant variant : variants()) {
            Path file = variant.file(this, spec);
            Run run = new Run();
            Outcome outcome = new Terminals<>(variant.maps(run, new Feature[0]), Object::toString, true, run).path(file);
            assertEquals(8, outcome.rows().size(), variant.name() + ": " + outcome.rows());
            assertTrue(outcome.rows().get(0).startsWith("ok {Name=alice, Score=10, Note=x"),
                    variant.name() + ": " + outcome.rows().get(0));
        }
    }

    /** Typed cells and locale-dependent text come out as the values they stand for, in every reader. */
    @Test
    void typedCellsAndConversionGiveTheExpectedValues() throws Exception {
        Spec spec = new Spec();
        feature("typedCells").shape().accept(spec);
        feature("cellConversion").shape().accept(spec);
        Feature conversion = feature("cellConversion");
        for (Variant variant : variants()) {
            Path file = variant.file(this, spec);
            for (Mode mode : List.of(Mode.SETTER, Mode.MAPPING)) {
                Outcome outcome = terminals(variant, mode, spec, new Run(), new Feature[]{conversion}).path(file);
                assertTrue(outcome.rows().get(0).startsWith("ok alice|10|x|2024-03-01|1234.5|2024-03-05"),
                        variant.name() + " " + mode + ": " + outcome.rows().get(0));
            }
        }
    }

    private static Feature feature(String name) {
        return FEATURES.stream().filter(f -> f.name().equals(name)).findFirst().orElseThrow();
    }

    private static List<Variant> variants() {
        return List.of(new CsvVariant(), new ExcelVariant(false), new ExcelVariant(true));
    }

    private Terminals<Row> terminals(Variant variant, Mode mode, Spec spec, Run run, Feature[] features) {
        return new Terminals<>(variant.rows(mode, spec, run, features), ReaderOptionCombinationTest::describeRow,
                !spec.typedCells, run);
    }

    private void check(List<String> problems, Feature... features) throws IOException {
        String names = features.length == 0 ? "(no options)"
                : String.join(" + ", Arrays.stream(features).map(Feature::name).toList());
        Spec spec = new Spec();
        for (Feature feature : features) feature.shape().accept(spec);
        boolean cancelled = Arrays.stream(features).anyMatch(f -> f.name().equals("cancel"));

        for (Mode mode : Mode.values()) {
            String label = mode + ": " + names;
            List<Outcome> byVariant = new ArrayList<>();
            for (Variant variant : variants()) {
                Path file = variant.file(this, spec);
                Terminals<?> terminals = mode == Mode.MAP
                        ? new Terminals<>(variant.maps(new Run(), features), Object::toString, !spec.typedCells, new Run())
                        : terminals(variant, mode, spec, new Run(), features);
                // Every call needs its own Run (its counters and events), so each way of reading gets a fresh one.
                byVariant.add(checkWays(problems, label, variant, mode, spec, features, file, cancelled));
            }
            for (int v = 1; v < byVariant.size(); v++) {
                same(problems, label, variants().get(0).name() + " and " + variants().get(v).name(), byVariant.get(0), byVariant.get(v));
            }
        }
    }

    /** Reads one file in each way, checks them against each other and returns the plain read. */
    private Outcome checkWays(List<String> problems, String label, Variant variant, Mode mode, Spec spec,
                              Feature[] features, Path file, boolean cancelled) {
        Function<Run, Terminals<?>> fresh = run -> mode == Mode.MAP
                ? new Terminals<>(variant.maps(run, features), Object::toString, !spec.typedCells, run)
                : terminals(variant, mode, spec, run, features);
        Outcome rows = fresh.apply(new Run()).path(file);
        same(problems, label, variant.name() + ": a Path and an InputStream", rows, fresh.apply(new Run()).stream(file));

        Outcome summary = fresh.apply(new Run()).summary(file);
        same(problems, label, variant.name() + ": read and readWithSummary rows", rows.rows(), summary.rows());
        long deliveredOk = rows.rows().stream().filter(r -> r.startsWith("ok")).count();
        List<Long> delivered = List.of((long) rows.rows().size(), deliveredOk);
        same(problems, label, variant.name() + ": the summary counts the delivered rows", delivered,
                summary.failure() != null ? delivered : List.of(summary.total(), summary.success()));
        same(problems, label, variant.name() + ": readWhile(always) and readWithSummary", summary,
                fresh.apply(new Run()).whileAll(file));
        if (!cancelled) {
            Outcome report = fresh.apply(new Run()).report(file);
            // A failed read has no counts of its own (the rows delivered before it are not a summary).
            boolean failed = summary.failure() != null || report.failure() != null;
            same(problems, label, variant.name() + ": readReport and readWithSummary counts",
                    failed ? List.of(String.valueOf(summary.failure()))
                            : List.of(summary.total(), summary.success(), summary.errors(), String.valueOf(summary.failure())),
                    failed ? List.of(String.valueOf(report.failure()))
                            : List.of(report.total(), report.success(), report.errors(), String.valueOf(report.failure())));
        }
        return rows;
    }

    private static void same(List<String> problems, String label, String what, Object expected, Object actual) {
        if (!expected.equals(actual)) {
            problems.add("[" + label + "] " + what + " differ:\n      " + expected + "\n      " + actual);
        }
    }
}
