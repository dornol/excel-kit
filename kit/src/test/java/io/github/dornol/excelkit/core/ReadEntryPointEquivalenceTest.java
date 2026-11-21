package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelSheetInfo;
import io.github.dornol.excelkit.excel.ExcelWorkbookReader;
import io.github.dornol.excelkit.excel.SheetReadReport;
import io.github.dornol.excelkit.excel.WorkbookReadReport;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Every way to hand input to a reader (an {@code InputStream}, a {@code Path}, an
 * {@link InputStreamSource}) and every way to consume the rows must give the same result for the
 * same file. Each entry point is run for all three kinds of input and compared with the
 * {@code InputStream} variant, whose result is also checked against what the file contains.
 */
class ReadEntryPointEquivalenceTest {

    enum Kind { STREAM, PATH, SOURCE }

    static class Row {
        String name;
        Integer qty;

        @Override
        public String toString() {
            return name + ":" + qty;
        }
    }

    /** Header, a good row, a row whose quantity is not a number, and another good row. */
    private static final String CSV = "Name,Qty\nA,1\nB,x\nC,3\n";
    private static final String CLEAN_CSV = "Name,Qty\nA,1\nC,3\n";

    @TempDir
    Path dir;

    // ---------------------------------------------------------------- fixtures

    private static byte[] csv(String text) {
        return text.getBytes(StandardCharsets.UTF_8);
    }

    /** A workbook with the given text cells, header row first. */
    private static byte[] workbook(String sheetName, String[][] rows) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet(sheetName);
            for (int r = 0; r < rows.length; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < rows[r].length; c++) {
                    row.createCell(c).setCellValue(rows[r][c]);
                }
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static final String[][] PEOPLE = {{"Name", "Qty"}, {"A", "1"}, {"B", "x"}, {"C", "3"}};
    private static final String[][] CLEAN_PEOPLE = {{"Name", "Qty"}, {"A", "1"}, {"C", "3"}};

    private static byte[] twoSheetWorkbook(String[][] first, String[][] second) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            String[][][] data = {first, second};
            String[] names = {"People", "Other"};
            for (int s = 0; s < 2; s++) {
                var sheet = workbook.createSheet(names[s]);
                for (int r = 0; r < data[s].length; r++) {
                    var row = sheet.createRow(r);
                    for (int c = 0; c < data[s][r].length; c++) {
                        row.createCell(c).setCellValue(data[s][r][c]);
                    }
                }
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static <R extends AbstractReader<Row, R>> R withColumns(R reader) {
        return reader
                .column("Name", (row, cell) -> row.name = cell.asString())
                .column("Qty", (row, cell) -> row.qty = cell.asInt());
    }

    private static CsvReader<Row> csvReader() {
        return withColumns(CsvReader.setter(Row::new));
    }

    private static ExcelReader<Row> excelReader() {
        return withColumns(ExcelReader.setter(Row::new));
    }

    // ------------------------------------------------------- normalised results

    private static String describe(ReadResult<Row> result) {
        return result.success() ? "ok:" + result.data() : "err:" + result.errorType() + "@" + result.fileRowNum();
    }

    private static String describe(RowError error) {
        return "err:" + error.type() + "@" + error.fileRowNum();
    }

    private static List<String> describeErrors(List<RowError> errors) {
        return errors.stream().map(ReadEntryPointEquivalenceTest::describe).toList();
    }

    private static List<Long> counts(ReadSummary summary) {
        return List.of(summary.totalRows(), summary.successRows(), summary.errorRows());
    }

    // ------------------------------------------------------- input dispatch

    interface Opening<R> {
        R apply(Kind kind) throws Exception;
    }

    private static <R> R by(Kind kind, Opening<R> opening) {
        try {
            return opening.apply(kind);
        } catch (RuntimeException e) {
            throw e;
        } catch (Exception e) {
            throw new AssertionError(e);
        }
    }

    /** Hands the same bytes to a reader as a stream, a file and an input source. */
    private final class Input {
        private final byte[] bytes;
        private final Path path;

        Input(byte[] bytes, String name) throws IOException {
            this.bytes = bytes;
            this.path = Files.write(dir.resolve(name), bytes);
        }

        InputStream stream() {
            return new ByteArrayInputStream(bytes);
        }

        InputStreamSource source() {
            return () -> new ByteArrayInputStream(bytes);
        }

        <R> R choose(Kind kind, Function<InputStream, R> stream, Function<Path, R> file,
                     Function<InputStreamSource, R> source) {
            return switch (kind) {
                case STREAM -> stream.apply(stream());
                case PATH -> file.apply(path);
                case SOURCE -> source.apply(source());
            };
        }
    }

    // ================================================================== CSV

    @Nested
    class Csv {
        private Input input(String text) throws IOException {
            return new Input(csv(text), "rows.csv");
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWithOneCallback(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            input.choose(kind, s -> { csvReader().read(s, r -> seen.add(describe(r))); return null; },
                    p -> { csvReader().read(p, r -> seen.add(describe(r))); return null; },
                    src -> { csvReader().read(src, r -> seen.add(describe(r))); return null; });

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWithSuccessAndErrorCallbacks(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            Consumer<Row> ok = row -> seen.add("ok:" + row);
            Consumer<RowError> error = e -> seen.add(describe(e));
            input.choose(kind, s -> { csvReader().read(s, ok, error); return null; },
                    p -> { csvReader().read(p, ok, error); return null; },
                    src -> { csvReader().read(src, ok, error); return null; });

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readStrictDeliversTheRowsBeforeTheFirstErrorAndThenAborts(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            Consumer<Row> ok = row -> seen.add(row.toString());

            assertThrows(ReadAbortException.class, () -> input.choose(kind,
                    s -> { csvReader().readStrict(s, ok); return null; },
                    p -> { csvReader().readStrict(p, ok); return null; },
                    src -> { csvReader().readStrict(src, ok); return null; }));

            assertEquals(List.of("A:1"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readStrictReadsACleanFileCompletely(Kind kind) throws IOException {
            Input input = input(CLEAN_CSV);
            List<String> seen = new ArrayList<>();
            Consumer<Row> ok = row -> seen.add(row.toString());
            input.choose(kind, s -> { csvReader().readStrict(s, ok); return null; },
                    p -> { csvReader().readStrict(p, ok); return null; },
                    src -> { csvReader().readStrict(src, ok); return null; });

            assertEquals(List.of("A:1", "C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWhileStopsWhenThePredicateSaysSo(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            java.util.function.Predicate<ReadResult<Row>> firstTwo = r -> {
                seen.add(describe(r));
                return seen.size() < 2;
            };
            ReadSummary summary = input.choose(kind, s -> csvReader().readWhile(s, firstTwo),
                    p -> csvReader().readWhile(p, firstTwo), src -> csvReader().readWhile(src, firstTwo));

            assertEquals(List.of("ok:A:1", "err:MAPPING@3"), seen);
            assertTrue(summary.stoppedEarly());
            assertEquals(List.of(2L, 1L, 1L), counts(summary));
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readAllReturnsTheRowsOfACleanFile(Kind kind) throws IOException {
            Input input = input(CLEAN_CSV);
            List<Row> rows = input.choose(kind, s -> csvReader().readAll(s),
                    p -> csvReader().readAll(p), src -> csvReader().readAll(src));

            assertEquals(List.of("A:1", "C:3"), rows.stream().map(Row::toString).toList());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readAllFailsOnABadRow(Kind kind) throws IOException {
            Input input = input(CSV);
            assertThrows(ReadAbortException.class, () -> input.choose(kind, s -> csvReader().readAll(s),
                    p -> csvReader().readAll(p), src -> csvReader().readAll(src)));
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readReportSummarisesTheFile(Kind kind) throws IOException {
            Input input = input(CSV);
            ReadReport report = input.choose(kind, s -> csvReader().readReport(s, 10),
                    p -> csvReader().readReport(p, 10), src -> csvReader().readReport(src, 10));

            assertEquals(List.of(3L, 2L, 1L), counts(report.summary()));
            assertEquals(List.of("err:MAPPING@3"), describeErrors(report.errors()));
            assertFalse(report.errorsTruncated());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readDetectedBehavesLikeRead(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            Consumer<ReadResult<Row>> collect = r -> seen.add(describe(r));
            input.choose(kind, s -> { csvReader().readDetected(s, collect); return null; },
                    p -> { csvReader().readDetected(p, collect); return null; },
                    src -> { csvReader().readDetected(src, collect); return null; });

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readDetectedWithSummaryReportsTheCounts(Kind kind) throws IOException {
            Input input = input(CSV);
            List<String> seen = new ArrayList<>();
            Consumer<ReadResult<Row>> collect = r -> seen.add(describe(r));
            ReadSummary summary = input.choose(kind, s -> csvReader().readDetectedWithSummary(s, collect),
                    p -> csvReader().readDetectedWithSummary(p, collect),
                    src -> csvReader().readDetectedWithSummary(src, collect));

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
            assertEquals(List.of(3L, 2L, 1L), counts(summary));
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void openGivesASessionThatReadsTheFile(Kind kind) throws IOException {
            Input input = input(CLEAN_CSV);
            List<String> seen = new ArrayList<>();
            try (ReadSession<Row> session = input.choose(kind, s -> csvReader().open(s),
                    p -> csvReader().open(p), src -> csvReader().open(src))) {
                session.read(r -> seen.add(describe(r)));
            }

            assertEquals(List.of("ok:A:1", "ok:C:3"), seen);
        }
    }

    // ================================================================ Excel

    @Nested
    class Excel {
        private Input input(String[][] rows) throws IOException {
            return new Input(workbook("Sheet1", rows), "rows.xlsx");
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWithOneCallback(Kind kind) throws IOException {
            Input input = input(PEOPLE);
            List<String> seen = new ArrayList<>();
            input.choose(kind, s -> { excelReader().read(s, r -> seen.add(describe(r))); return null; },
                    p -> { excelReader().read(p, r -> seen.add(describe(r))); return null; },
                    src -> { excelReader().read(src, r -> seen.add(describe(r))); return null; });

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWithSuccessAndErrorCallbacks(Kind kind) throws IOException {
            Input input = input(PEOPLE);
            List<String> seen = new ArrayList<>();
            Consumer<Row> ok = row -> seen.add("ok:" + row);
            Consumer<RowError> error = e -> seen.add(describe(e));
            input.choose(kind, s -> { excelReader().read(s, ok, error); return null; },
                    p -> { excelReader().read(p, ok, error); return null; },
                    src -> { excelReader().read(src, ok, error); return null; });

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readStrictDeliversTheRowsBeforeTheFirstErrorAndThenAborts(Kind kind) throws IOException {
            Input input = input(PEOPLE);
            List<String> seen = new ArrayList<>();
            Consumer<Row> ok = row -> seen.add(row.toString());

            assertThrows(ReadAbortException.class, () -> input.choose(kind,
                    s -> { excelReader().readStrict(s, ok); return null; },
                    p -> { excelReader().readStrict(p, ok); return null; },
                    src -> { excelReader().readStrict(src, ok); return null; }));

            assertEquals(List.of("A:1"), seen);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readWhileStopsWhenThePredicateSaysSo(Kind kind) throws IOException {
            Input input = input(PEOPLE);
            List<String> seen = new ArrayList<>();
            java.util.function.Predicate<ReadResult<Row>> firstTwo = r -> {
                seen.add(describe(r));
                return seen.size() < 2;
            };
            ReadSummary summary = input.choose(kind, s -> excelReader().readWhile(s, firstTwo),
                    p -> excelReader().readWhile(p, firstTwo), src -> excelReader().readWhile(src, firstTwo));

            assertEquals(List.of("ok:A:1", "err:MAPPING@3"), seen);
            assertTrue(summary.stoppedEarly());
            assertEquals(List.of(2L, 1L, 1L), counts(summary));
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readAllReturnsTheRowsOfACleanFile(Kind kind) throws IOException {
            Input input = input(CLEAN_PEOPLE);
            List<Row> rows = input.choose(kind, s -> excelReader().readAll(s),
                    p -> excelReader().readAll(p), src -> excelReader().readAll(src));

            assertEquals(List.of("A:1", "C:3"), rows.stream().map(Row::toString).toList());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void readReportSummarisesTheFile(Kind kind) throws IOException {
            Input input = input(PEOPLE);
            ReadReport report = input.choose(kind, s -> excelReader().readReport(s, 10),
                    p -> excelReader().readReport(p, 10), src -> excelReader().readReport(src, 10));

            assertEquals(List.of(3L, 2L, 1L), counts(report.summary()));
            assertEquals(List.of("err:MAPPING@3"), describeErrors(report.errors()));
            assertFalse(report.errorsTruncated());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void sheetNamesAndHeadersCanBeReadFromAnyKindOfInput(Kind kind) throws IOException {
            Input input = input(PEOPLE);

            List<ExcelSheetInfo> sheets = input.choose(kind, ExcelReader::getSheetNames,
                    ExcelReader::getSheetNames, ExcelReader::getSheetNames);
            List<String> headers = input.choose(kind, s -> ExcelReader.getSheetHeaders(s, 0, 0),
                    p -> ExcelReader.getSheetHeaders(p, 0, 0), src -> ExcelReader.getSheetHeaders(src, 0, 0));

            assertEquals(List.of("Sheet1"), sheets.stream().map(ExcelSheetInfo::name).toList());
            assertEquals(List.of("Name", "Qty"), headers);
        }
    }

    // ============================================================ file detection

    @Nested
    class Detection {
        @ParameterizedTest
        @EnumSource(Kind.class)
        void aCsvFileIsRecognisedWhateverTheKindOfInput(Kind kind) throws IOException {
            Input input = new Input(csv(CSV), "rows.csv");

            TabularFileType type = input.choose(kind, TabularFileDetector::detect,
                    TabularFileDetector::detect, TabularFileDetector::detect);
            TabularDetectionResult detailed = input.choose(kind, TabularFileDetector::detectDetailed,
                    TabularFileDetector::detectDetailed, TabularFileDetector::detectDetailed);
            TabularDetectionResult sampled = input.choose(kind, s -> TabularFileDetector.detectDetailed(s, 64),
                    p -> TabularFileDetector.detectDetailed(p), src -> TabularFileDetector.detectDetailed(src, 64));

            assertEquals(TabularFileType.CSV, type);
            assertEquals(TabularFileType.CSV, detailed.type());
            assertEquals(',', detailed.delimiter());
            assertEquals(TabularFileType.CSV, sampled.type());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void anXlsxFileIsRecognisedWhateverItsName(Kind kind) throws IOException {
            Input input = new Input(workbook("Sheet1", PEOPLE), "not-an-xlsx-name.csv");

            TabularFileType type = input.choose(kind, TabularFileDetector::detect,
                    TabularFileDetector::detect, TabularFileDetector::detect);
            TabularDetectionResult detailed = input.choose(kind, TabularFileDetector::detectDetailed,
                    TabularFileDetector::detectDetailed, TabularFileDetector::detectDetailed);

            assertEquals(TabularFileType.XLSX, type);
            assertEquals(TabularFileType.XLSX, detailed.type());
        }

        @Test
        void allKindsOfInputGiveTheSameDetailedResult() throws IOException {
            Input input = new Input(csv(CSV), "rows.csv");

            TabularDetectionResult fromStream = TabularFileDetector.detectDetailed(input.stream());

            assertEquals(fromStream, TabularFileDetector.detectDetailed(input.path));
            assertEquals(fromStream, TabularFileDetector.detectDetailed(input.source()));
        }
    }

    // ===================================================== several sheets at once

    @Nested
    class Workbook {
        private final ExcelKitSchema<Row> schema = ExcelKitSchema.<Row>builder()
                .column("Name", row -> row.name, (row, cell) -> row.name = cell.asString())
                .column("Qty", row -> row.qty, (row, cell) -> row.qty = cell.asInt())
                .build();

        private WorkbookReadReport read(Kind kind, Input input, ExcelWorkbookReader reader) {
            return input.choose(kind, reader::read, reader::read, reader::read);
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void eachSheetIsReadByItsOwnReaderWhateverTheKindOfInput(Kind kind) throws IOException {
            Input input = new Input(twoSheetWorkbook(PEOPLE, CLEAN_PEOPLE), "two.xlsx");
            List<String> people = new ArrayList<>();
            List<String> other = new ArrayList<>();

            WorkbookReadReport report = read(kind, input, ExcelWorkbookReader.create()
                    .sheet("People", excelReader(), r -> people.add(describe(r)))
                    .sheet(1, schema, Row::new, null, r -> other.add(describe(r))));

            assertEquals(List.of("ok:A:1", "err:MAPPING@3", "ok:C:3"), people);
            assertEquals(List.of("ok:A:1", "ok:C:3"), other);
            assertEquals(List.of("People", "Other"), report.sheets().stream().map(SheetReadReport::name).toList());
            assertEquals(List.of(3L, 2L, 1L), counts(report.sheet("People").summary()));
            assertEquals(List.of(2L, 2L, 0L), counts(report.sheet("Other").summary()));
            assertEquals(List.of("err:MAPPING@3"), describeErrors(report.sheet("People").errors()));
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void sheetsGivenByNameAndBySchemaBehaveTheSame(Kind kind) throws IOException {
            Input input = new Input(twoSheetWorkbook(PEOPLE, CLEAN_PEOPLE), "two.xlsx");
            List<String> viaReader = new ArrayList<>();
            List<String> viaSchema = new ArrayList<>();

            read(kind, input, ExcelWorkbookReader.create().sheet("People", excelReader(), r -> viaReader.add(describe(r))));
            read(kind, input, ExcelWorkbookReader.create().sheet("People", schema, Row::new, null,
                    r -> viaSchema.add(describe(r))));

            assertEquals(viaReader, viaSchema);
            assertEquals(3, viaSchema.size());
        }

        @ParameterizedTest
        @EnumSource(Kind.class)
        void maxErrorsPerSheetCapsTheErrorsReportedForASheet(Kind kind) throws IOException {
            String[][] twoBadRows = {{"Name", "Qty"}, {"A", "x"}, {"B", "y"}, {"C", "3"}};
            Input input = new Input(twoSheetWorkbook(twoBadRows, CLEAN_PEOPLE), "bad.xlsx");

            WorkbookReadReport report = read(kind, input, ExcelWorkbookReader.create()
                    .maxErrorsPerSheet(1)
                    .sheet("People", excelReader(), r -> { }));

            SheetReadReport people = report.sheet("People");
            assertEquals(1, people.errors().size());
            assertTrue(people.errorsTruncated());
        }
    }

    // ================================================================ CSV options

    @Nested
    class CsvOptions {
        private List<String> pairs(String text, Consumer<CsvReader<String[]>> configure) {
            List<String> out = new ArrayList<>();
            CsvReader<String[]> reader = CsvReader.<String[]>setter(() -> new String[2])
                    .column("Name", (row, cell) -> row[0] = cell.asString())
                    .column("Qty", (row, cell) -> row[1] = cell.asString());
            configure.accept(reader);
            reader.read(new ByteArrayInputStream(csv(text)), result -> out.add(
                    result.success() ? result.data()[0] + "|" + result.data()[1] : "error"));
            return out;
        }

        @Test
        void strictQuotesDropsTextOutsideTheQuotes() {
            String text = "\"Name\",\"Qty\"\n\"A\"junk,\"1\"\n";

            assertEquals(List.of("A|1"), pairs(text, reader -> reader.strictQuotes(true)));
            assertEquals(List.of("Ajunk|1"), pairs(text, reader -> reader.strictQuotes(false)));
        }

        @Test
        void strictQuotesNeedsAQuotedHeaderRow() {
            String unquotedHeader = "Name,Qty\n\"A\",\"1\"\n";

            assertEquals(List.of("A|1"), pairs(unquotedHeader, reader -> reader.strictQuotes(false)));
            assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                    () -> pairs(unquotedHeader, reader -> reader.strictQuotes(true)));
        }

        @Test
        void aBlankedHeaderUnderStrictQuotesIsExplained() {
            String unquotedHeader = "Name,Qty\n\"A\",\"1\"\n";

            var error = assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                    () -> pairs(unquotedHeader, reader -> reader.strictQuotes(true)));

            assertTrue(error.getMessage().contains("strictQuotes"), error.getMessage());
            assertTrue(error.getMessage().contains("every header cell must be quoted"), error.getMessage());
            assertTrue(error.getMessage().contains("Header aliases"), "the original cause is quoted");
            assertTrue(error.getCause() instanceof ExcelKitException, "and still attached");
        }

        @Test
        void aMissingHeaderIsNamedInTheMessage() {
            String wrongHeader = "Wrong,Qty\n\"A\",\"1\"\n";

            var error = assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                    () -> pairs(wrongHeader, reader -> { }));

            assertTrue(error.getMessage().startsWith("Failed to read CSV: "), error.getMessage());
            assertTrue(error.getMessage().contains("Header aliases [Name] not found"), error.getMessage());
            assertTrue(error.getMessage().contains("Available headers: [Wrong, Qty]"), error.getMessage());
        }

        @Test
        void aMissingHeaderUnderStrictQuotesIsNotBlamedOnQuoting() {
            String wrongHeader = "\"Wrong\",\"Qty\"\n\"A\",\"1\"\n";

            var error = assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                    () -> pairs(wrongHeader, reader -> reader.strictQuotes(true)));

            assertTrue(error.getMessage().contains("Header aliases [Name] not found"), error.getMessage());
            assertFalse(error.getMessage().contains("strictQuotes"), error.getMessage());
        }

        @Test
        void ignoreLeadingWhiteSpaceSkipsBlanksBeforeAQuotedValue() {
            String text = "Name,Qty\nA, \"1\"\n";

            assertEquals(List.of("A|1"), pairs(text, reader -> reader.ignoreLeadingWhiteSpace(true)));
        }

        @Test
        void blanksAroundAnUnquotedValueAreKeptEitherWay() {
            String text = "Name,Qty\n  A, 1\n";

            assertEquals(List.of("  A| 1"), pairs(text, reader -> reader.ignoreLeadingWhiteSpace(true)));
            assertEquals(List.of("  A| 1"), pairs(text, reader -> reader.ignoreLeadingWhiteSpace(false)));
        }

        @Test
        void requireHeadersIsAnAliasOfStrictHeaders() {
            assertEquals(csvReader().strictHeaders().snapshotReadOptions(),
                    csvReader().requireHeaders().snapshotReadOptions());
            assertTrue(csvReader().requireHeaders().snapshotReadOptions().strictHeaders());
        }
    }
}
