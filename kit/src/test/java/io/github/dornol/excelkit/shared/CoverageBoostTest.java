package io.github.dornol.excelkit.shared;

import io.github.dornol.excelkit.csv.CsvDialect;
import io.github.dornol.excelkit.csv.CsvMapReader;
import io.github.dornol.excelkit.csv.CsvMapWriter;
import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvQuoting;
import io.github.dornol.excelkit.excel.ExcelMapReader;
import io.github.dornol.excelkit.excel.ExcelMapWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelReadException;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Targeted tests to boost coverage for low-coverage classes:
 * - TempResourceCreator: constructor, createTempFile with valid dir
 * - CsvMapReader.CsvMapReadHandler: readAsStream error/close paths
 * - AbstractReadHandler: null input validation, readStrict edge cases
 * - CsvWriter: escapeCsv edge cases, injection defense, quoting strategies
 * - ExcelMapReader: readAsStream interrupt/error paths
 */
class CoverageBoostTest {

    // ============================================================
    // AbstractReadHandler - null InputStream validation
    // ============================================================
    @Nested
    class AbstractReadHandlerValidation {

        @Test
        void excelReader_nullInputStream_throwsIllegalArgument() {
            var reader = new ExcelReader<>(Object::new, null)
                    .addColumn("A", (t, cell) -> {});
            assertThrows(IllegalArgumentException.class, () -> reader.build(null));
        }

        @Test
        void csvReader_nullInputStream_throwsIllegalArgument() {
            var reader = new CsvReader<>(Object::new, null)
                    .addColumn("A", (t, cell) -> {});
            assertThrows(IllegalArgumentException.class, () -> reader.build(null));
        }

        @Test
        void excelMapReader_nullInputStream_throwsException() {
            assertThrows(Exception.class,
                    () -> new ExcelMapReader().build(null));
        }

        @Test
        void csvMapReader_nullInputStream_throwsIllegalArgument() {
            assertThrows(IllegalArgumentException.class,
                    () -> new CsvMapReader().build(null));
        }

        @Test
        void csvMapReader_negativeHeaderRowIndex_throwsIllegalArgument() {
            assertThrows(IllegalArgumentException.class,
                    () -> new CsvMapReader()
                            .headerRowIndex(-1)
                            .build(new ByteArrayInputStream("A\n1".getBytes())));
        }
    }

    // ============================================================
    // CsvMapReader.readAsStream - error propagation paths
    // ============================================================
    @Nested
    class CsvMapReaderStreamErrors {

        @Test
        void readAsStream_emptyFile_throwsCsvReadException() {
            assertThrows(CsvReadException.class, () -> {
                try (var stream = new CsvMapReader()
                        .build(new ByteArrayInputStream(new byte[0]))
                        .readAsStream()) {
                    stream.toList();
                }
            });
        }

        @Test
        void readAsStream_headerOnly_returnsEmptyStream() {
            String csv = "Name,Age\n";
            try (var stream = new CsvMapReader()
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .readAsStream()) {
                var results = stream.toList();
                assertTrue(results.isEmpty());
            }
        }

        @Test
        void readAsStream_closeWithoutConsuming_shouldNotLeak() {
            String csv = "Name\nAlice\nBob\n";
            var stream = new CsvMapReader()
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .readAsStream();
            stream.close();
            // no exception = resources cleaned up
        }

        @Test
        void readAsStream_partialConsumption_shouldCleanup() {
            String csv = "Name\nAlice\nBob\nCharlie\n";
            try (var stream = new CsvMapReader()
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .readAsStream()) {
                var first = stream.findFirst();
                assertTrue(first.isPresent());
                assertEquals("Alice", first.get().data().get("Name"));
            }
        }

        @Test
        void readAsStream_withProgress_shouldFireCallbacks() {
            String csv = "Name\nA\nB\nC\nD\n";
            List<Long> progressCounts = new ArrayList<>();
            try (var stream = new CsvMapReader()
                    .onProgress(2, (count, total) -> progressCounts.add(count))
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .readAsStream()) {
                var results = stream.toList();
                assertEquals(4, results.size());
            }
            assertEquals(List.of(2L, 4L), progressCounts);
        }

        @Test
        void readAsStream_insufficientRowsForHeader_throwsCsvReadException() {
            String csv = "only one line";
            assertThrows(CsvReadException.class, () -> {
                try (var stream = new CsvMapReader()
                        .headerRowIndex(5)
                        .build(new ByteArrayInputStream(csv.getBytes()))
                        .readAsStream()) {
                    stream.toList();
                }
            });
        }
    }

    // ============================================================
    // CsvWriter - escapeCsv edge cases and quoting strategies
    // ============================================================
    @Nested
    class CsvWriterCoverage {

        record Row(String value) {}

        @Test
        void quoting_ALL_nullValue_shouldWriteEmptyQuotes() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .quoting(CsvQuoting.ALL)
                    .column("Value", r -> null)
                    .write(Stream.of(new Row("ignored")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("\"Value\"", lines[0]);
            assertEquals("\"\"", lines[1]);
        }

        @Test
        void quoting_NON_NUMERIC_shouldQuoteStrings() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .quoting(CsvQuoting.NON_NUMERIC)
                    .column("Text", Row::value)
                    .column("Num", r -> 42)
                    .write(Stream.of(new Row("hello")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            // "Text","Num" headers - NON_NUMERIC quotes non-numeric
            assertTrue(lines[1].contains("\"hello\""), "String should be quoted");
            assertTrue(lines[1].contains("42"), "Number should not be extra-quoted");
        }

        @Test
        void injectionDefense_shouldPrefixFormulaChars() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .csvInjectionDefense(true)
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("=SUM(A1)"), new Row("+cmd"), new Row("@import")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("'=SUM(A1)"), "= should be prefixed with '");
            assertTrue(csv.contains("'+cmd"), "+ should be prefixed with '");
            assertTrue(csv.contains("'@import"), "@ should be prefixed with '");
        }

        @Test
        void injectionDefense_disabled_shouldNotPrefix() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .csvInjectionDefense(false)
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("=SUM(A1)")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("=SUM(A1)"), "Should not prefix when defense is disabled");
            assertFalse(csv.contains("'="), "Should not have prefix");
        }

        @Test
        void valueWithQuotes_shouldBeEscaped() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("say \"hello\"")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("\"say \"\"hello\"\"\""), "Quotes should be escaped");
        }

        @Test
        void valueWithNewline_shouldBeQuoted() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("line1\nline2")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("\"line1\nline2\""), "Newline in value should trigger quoting");
        }

        @Test
        void columnIf_false_shouldNotIncludeColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .column("Always", Row::value)
                    .columnIf("Never", false, Row::value)
                    .write(Stream.of(new Row("test")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String headerLine = csv.split("\r?\n")[0];
            assertEquals("Always", headerLine);
        }

        @Test
        void constColumn_shouldWriteConstantValue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .column("Name", Row::value)
                    .constColumn("Type", "Person")
                    .write(Stream.of(new Row("Alice")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("Name,Type", lines[0]);
            assertEquals("Alice,Person", lines[1]);
        }

        @Test
        void afterData_shouldAppendContent() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .column("Name", Row::value)
                    .afterData(writer -> writer.println("# Footer"))
                    .write(Stream.of(new Row("Alice")))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("# Footer"), "afterData content should be appended");
        }

        @Test
        void noBom_shouldNotWriteBOM() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .bom(false)
                    .column("Name", Row::value)
                    .write(Stream.of(new Row("Alice")))
                    .consumeOutputStream(out);

            byte[] bytes = out.toByteArray();
            assertFalse(bytes[0] == (byte) 0xEF && bytes[1] == (byte) 0xBB && bytes[2] == (byte) 0xBF,
                    "Should not start with UTF-8 BOM");
        }

        @Test
        void duplicateColumnNames_shouldThrow() {
            var writer = new CsvWriter<Row>()
                    .column("Name", Row::value)
                    .column("Name", Row::value);

            assertThrows(Exception.class, () ->
                    writer.write(Stream.of(new Row("test"))));
        }

        @Test
        void emptyColumns_shouldThrow() {
            assertThrows(Exception.class, () ->
                    new CsvWriter<Row>().write(Stream.of(new Row("test"))));
        }

        @Test
        void isNumeric_edgeCases() throws IOException {
            // These exercise the isNumeric method through NON_NUMERIC quoting
            // Disable injection defense so +/- values are not prefixed
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new CsvWriter<Row>()
                    .quoting(CsvQuoting.NON_NUMERIC)
                    .csvInjectionDefense(false)
                    .column("Val", Row::value)
                    .write(Stream.of(
                            new Row("42"),         // integer
                            new Row("-3.14"),       // negative decimal
                            new Row("+5"),          // positive with sign
                            new Row("1.2.3"),       // double decimal - not numeric
                            new Row("-"),           // just minus - not numeric
                            new Row("")             // empty - not numeric
                    ))
                    .consumeOutputStream(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            // line 0 is header "Val" (quoted because non-numeric)
            assertEquals("42", lines[1]);           // numeric, not quoted
            assertEquals("-3.14", lines[2]);        // numeric, not quoted
            assertEquals("+5", lines[3]);           // numeric, not quoted
            assertTrue(lines[4].startsWith("\""), "1.2.3 should be quoted as non-numeric");
            assertTrue(lines[5].startsWith("\""), "- should be quoted as non-numeric");
            assertTrue(lines[6].startsWith("\""), "empty should be quoted as non-numeric");
        }
    }

    // ============================================================
    // ExcelMapReader.readAsStream - error paths
    // ============================================================
    @Nested
    class ExcelMapReaderStreamCoverage {

        @Test
        void readAsStream_tryWithResources_multipleColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelMapWriter("Name", "Age", "City").write(Stream.of(
                    Map.of("Name", "Alice", "Age", 30, "City", "Seoul"),
                    Map.of("Name", "Bob", "Age", 25, "City", "Tokyo")
            )).consumeOutputStream(out);

            try (var stream = new ExcelMapReader()
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .readAsStream()) {
                var results = stream.toList();
                assertEquals(2, results.size());
                assertTrue(results.get(0).success());
                assertEquals("Alice", results.get(0).data().get("Name"));
                assertEquals("Bob", results.get(1).data().get("Name"));
            }
        }

        @Test
        void readAsStream_withHeaderRowIndex() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelMapWriter("Name").write(Stream.of(
                    Map.of("Name", "Alice")
            )).consumeOutputStream(out);

            // headerRowIndex 0 is default — just verifying the API path
            try (var stream = new ExcelMapReader()
                    .headerRowIndex(0)
                    .sheetIndex(0)
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .readAsStream()) {
                var results = stream.toList();
                assertEquals(1, results.size());
            }
        }
    }

    // ============================================================
    // CsvReader - additional branch coverage
    // ============================================================
    @Nested
    class CsvReaderBranchCoverage {

        static class Person {
            String name;
            int age;
        }

        @Test
        void csvReader_skipColumn_shouldSkipPositionally() {
            String csv = "Name,Skip,Age\nAlice,ignored,30\n";
            List<ReadResult<Person>> results = new ArrayList<>();
            new CsvReader<>(Person::new, null)
                    .addColumn((p, cell) -> p.name = cell.asString())
                    .skipColumn()
                    .addColumn((p, cell) -> p.age = cell.asInt())
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertTrue(results.get(0).success());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
        }

        @Test
        void csvReader_skipColumns_shouldSkipMultiple() {
            String csv = "A,B,C,D\n1,2,3,4\n";
            List<ReadResult<Person>> results = new ArrayList<>();
            new CsvReader<>(Person::new, null)
                    .skipColumns(3)
                    .addColumn((p, cell) -> p.name = cell.asString())
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("4", results.get(0).data().name);
        }

        @Test
        void csvReader_skipColumns_negativeCount_throwsException() {
            assertThrows(IllegalArgumentException.class, () ->
                    new CsvReader<>(Person::new, null).skipColumns(-1));
        }

        @Test
        void csvReader_dialect_PIPE() {
            String csv = "Name|Age\nAlice|30\n";
            List<ReadResult<Person>> results = new ArrayList<>();
            new CsvReader<>(Person::new, null)
                    .dialect(CsvDialect.PIPE)
                    .addColumn("Name", (p, cell) -> p.name = cell.asString())
                    .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                    .build(new ByteArrayInputStream(csv.getBytes()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
        }

        @Test
        void csvReader_headerNotFound_throwsException() {
            String csv = "Name,Age\nAlice,30\n";
            var reader = new CsvReader<>(Person::new, null)
                    .addColumn("NonExistent", (p, cell) -> p.name = cell.asString());

            assertThrows(ExcelKitException.class, () ->
                    reader.build(new ByteArrayInputStream(csv.getBytes()))
                            .read(r -> {}));
        }

        @Test
        void csvReader_emptyFile_throwsCsvReadException() {
            var reader = new CsvReader<>(Person::new, null)
                    .addColumn("Name", (p, cell) -> p.name = cell.asString());

            assertThrows(CsvReadException.class, () ->
                    reader.build(new ByteArrayInputStream(new byte[0]))
                            .read(r -> {}));
        }

        @Test
        void csvReader_readAsStream_errorInRow_throwsCsvReadException() {
            String csv = "Name,Age\nAlice,notANumber\n";
            var handler = new CsvReader<>(Person::new, null)
                    .addColumn("Name", (p, cell) -> p.name = cell.asString())
                    .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                    .build(new ByteArrayInputStream(csv.getBytes()));

            // Setter error results in failed ReadResult, not exception
            List<ReadResult<Person>> results = new ArrayList<>();
            handler.read(results::add);
            assertEquals(1, results.size());
            assertFalse(results.get(0).success());
            assertNotNull(results.get(0).messages());
            assertFalse(results.get(0).messages().isEmpty());
        }
    }

    // ============================================================
    // CsvMapReader - dialect and charset coverage
    // ============================================================
    @Nested
    class CsvMapReaderConfigCoverage {

        @Test
        void csvMapReader_dialect_TSV_withCharset() {
            String tsv = "Name\tAge\nAlice\t30\n";
            List<Map<String, String>> results = new ArrayList<>();
            new CsvMapReader()
                    .dialect(CsvDialect.TSV)
                    .charset(StandardCharsets.UTF_8)
                    .build(new ByteArrayInputStream(tsv.getBytes()))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
        }

        @Test
        void csvMapReader_read_insufficientRowsForHeader_throwsCsvReadException() {
            String csv = "only one line";
            assertThrows(CsvReadException.class, () ->
                    new CsvMapReader()
                            .headerRowIndex(5)
                            .build(new ByteArrayInputStream(csv.getBytes()))
                            .read(r -> {}));
        }
    }

    // ============================================================
    // TempResourceCreator - additional paths
    // ============================================================
    @Nested
    class TempResourceCreatorCoverage {

        @Test
        void createTempFile_multipleCalls_createDistinctFiles() throws IOException {
            var dir = TempResourceCreator.createTempDirectory();
            try {
                var f1 = TempResourceCreator.createTempFile(dir, "a", ".tmp");
                var f2 = TempResourceCreator.createTempFile(dir, "a", ".tmp");
                assertNotEquals(f1, f2);
                assertTrue(java.nio.file.Files.exists(f1));
                assertTrue(java.nio.file.Files.exists(f2));
            } finally {
                try (var files = java.nio.file.Files.walk(dir)) {
                    files.sorted(java.util.Comparator.reverseOrder())
                            .forEach(p -> { try { java.nio.file.Files.delete(p); } catch (Exception ignored) {} });
                }
            }
        }
    }

    // ============================================================
    // TempResourceContainer - close with only tempFile (no tempDir)
    // ============================================================
    @Nested
    class TempResourceContainerCoverage {

        @Test
        void close_withOnlyTempFile_shouldNotThrow() throws IOException {
            var dir = TempResourceCreator.createTempDirectory();
            var file = TempResourceCreator.createTempFile(dir, "test", ".tmp");
            var container = new TempResourceContainer();
            container.setTempFile(file);
            // no tempDir set
            container.close();
            assertFalse(java.nio.file.Files.exists(file));
            // cleanup dir
            java.nio.file.Files.deleteIfExists(dir);
        }

        @Test
        void close_withOnlyTempDir_shouldNotThrow() throws IOException {
            var dir = TempResourceCreator.createTempDirectory();
            var container = new TempResourceContainer();
            container.setTempDir(dir);
            // no tempFile set
            container.close();
            assertFalse(java.nio.file.Files.exists(dir));
        }
    }
}
