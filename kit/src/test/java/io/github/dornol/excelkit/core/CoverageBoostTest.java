package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvDialect;
import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvQuoting;
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










    }

    // ============================================================
    // CsvMapReader.readAsStream - error propagation paths
    // ============================================================
    @Nested
    class CsvMapReaderStreamErrors {












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
            CsvWriter.<Row>create()
                    .quoting(CsvQuoting.ALL)
                    .column("Value", r -> null)
                    .write(Stream.of(new Row("ignored")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("\"Value\"", lines[0]);
            assertEquals("\"\"", lines[1]);
        }

        @Test
        void quoting_NON_NUMERIC_shouldQuoteStrings() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .quoting(CsvQuoting.NON_NUMERIC)
                    .column("Text", Row::value)
                    .column("Num", r -> 42)
                    .write(Stream.of(new Row("hello")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            // "Text","Num" headers - NON_NUMERIC quotes non-numeric
            assertTrue(lines[1].contains("\"hello\""), "String should be quoted");
            assertTrue(lines[1].contains("42"), "Number should not be extra-quoted");
        }

        @Test
        void injectionDefense_shouldPrefixFormulaChars() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .csvInjectionDefense(true)
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("=SUM(A1)"), new Row("+cmd"), new Row("@import")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("'=SUM(A1)"), "= should be prefixed with '");
            assertTrue(csv.contains("'+cmd"), "+ should be prefixed with '");
            assertTrue(csv.contains("'@import"), "@ should be prefixed with '");
        }

        @Test
        void injectionDefense_disabled_shouldNotPrefix() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .csvInjectionDefense(false)
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("=SUM(A1)")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("=SUM(A1)"), "Should not prefix when defense is disabled");
            assertFalse(csv.contains("'="), "Should not have prefix");
        }

        @Test
        void valueWithQuotes_shouldBeEscaped() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("say \"hello\"")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("\"say \"\"hello\"\"\""), "Quotes should be escaped");
        }

        @Test
        void valueWithNewline_shouldBeQuoted() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .column("Val", Row::value)
                    .write(Stream.of(new Row("line1\nline2")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("\"line1\nline2\""), "Newline in value should trigger quoting");
        }

        @Test
        void columnIf_false_shouldNotIncludeColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .column("Always", Row::value)
                    .columnIf("Never", false, Row::value)
                    .write(Stream.of(new Row("test")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String headerLine = csv.split("\r?\n")[0];
            assertEquals("Always", headerLine);
        }

        @Test
        void constColumn_shouldWriteConstantValue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .column("Name", Row::value)
                    .constColumn("Type", "Person")
                    .write(Stream.of(new Row("Alice")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("Name,Type", lines[0]);
            assertEquals("Alice,Person", lines[1]);
        }

        @Test
        void afterData_shouldAppendContent() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .column("Name", Row::value)
                    .afterData(writer -> writer.println("# Footer"))
                    .write(Stream.of(new Row("Alice")))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            assertTrue(csv.contains("# Footer"), "afterData content should be appended");
        }

        @Test
        void noBom_shouldNotWriteBOM() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
                    .bom(false)
                    .column("Name", Row::value)
                    .write(Stream.of(new Row("Alice")))
                    .writeTo(out);

            byte[] bytes = out.toByteArray();
            assertFalse(bytes[0] == (byte) 0xEF && bytes[1] == (byte) 0xBB && bytes[2] == (byte) 0xBF,
                    "Should not start with UTF-8 BOM");
        }

        @Test
        void duplicateColumnNames_shouldThrow() {
            var writer = CsvWriter.<Row>create()
                    .column("Name", Row::value)
                    .column("Name", Row::value);

            assertThrows(Exception.class, () ->
                    writer.write(Stream.of(new Row("test"))));
        }

        @Test
        void emptyColumns_shouldThrow() {
            assertThrows(Exception.class, () ->
                    CsvWriter.<Row>create().write(Stream.of(new Row("test"))));
        }

        @Test
        void isNumeric_edgeCases() throws IOException {
            // These exercise the isNumeric method through NON_NUMERIC quoting
            // Disable injection defense so +/- values are not prefixed
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<Row>create()
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
                    .writeTo(out);

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
                    .column((p, cell) -> p.name = cell.asString())
                    .skipColumn()
                    .column((p, cell) -> p.age = cell.asInt())
                    .read(new ByteArrayInputStream(csv.getBytes()), results::add);

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
                    .column((p, cell) -> p.name = cell.asString())
                    .read(new ByteArrayInputStream(csv.getBytes()), results::add);

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
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("Age", (p, cell) -> p.age = cell.asInt())
                    .read(new ByteArrayInputStream(csv.getBytes()), results::add);

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
        }

        @Test
        void csvReader_headerNotFound_throwsException() {
            String csv = "Name,Age\nAlice,30\n";
            var reader = new CsvReader<>(Person::new, null)
                    .column("NonExistent", (p, cell) -> p.name = cell.asString());

            assertThrows(ExcelKitException.class, () ->
                    reader.read(new ByteArrayInputStream(csv.getBytes()), r -> {}));
        }

        @Test
        void csvReader_emptyFile_throwsCsvReadException() {
            var reader = new CsvReader<>(Person::new, null)
                    .column("Name", (p, cell) -> p.name = cell.asString());

            assertThrows(CsvReadException.class, () ->
                    reader.read(new ByteArrayInputStream(new byte[0]), r -> {}));
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
            CsvReader.forMap()
                    .dialect(CsvDialect.TSV)
                    .charset(StandardCharsets.UTF_8)
                    .read(new ByteArrayInputStream(tsv.getBytes()), r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
        }

        @Test
        void csvMapReader_read_insufficientRowsForHeader_throwsCsvReadException() {
            String csv = "only one line";
            assertThrows(CsvReadException.class, () ->
                    CsvReader.forMap()
                            .headerRowIndex(5)
                            .read(new ByteArrayInputStream(csv.getBytes()), r -> {}));
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
