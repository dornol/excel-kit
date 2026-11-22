package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CsvWriterEdgeCaseTest {

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
