package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.UnaryOperator;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The options of {@link CsvReader} that are CSV's own: delimiter, quote and escape characters, encoding, malformed input. */
class CsvReaderOptionsTest {

    private static final Charset EUC_KR = Charset.forName("EUC-KR");

    private static List<String> read(byte[] data, UnaryOperator<CsvReader<Map<String, String>>> options) {
        List<String> rows = new ArrayList<>();
        options.apply(CsvReader.forMap()).read(new ByteArrayInputStream(data), result -> {
            assertTrue(result.success(), String.valueOf(result.messages()));
            rows.add(result.data() + "@" + result.fileRowNum());
        });
        return rows;
    }

    private static List<String> read(String csv, UnaryOperator<CsvReader<Map<String, String>>> options) {
        return read(csv.getBytes(StandardCharsets.UTF_8), options);
    }

    private static List<String> read(String csv) {
        return read(csv, reader -> reader);
    }

    // ---- quoting and escaping

    @Test
    void theQuoteCharacterCanBeChanged() {
        assertEquals(List.of("{a=x,1, b=y}@2"), read("a,b\n'x,1','y'\n", r -> r.quoteChar('\'')));
        // the default quote is then an ordinary character
        assertEquals(List.of("{a=\"x, b=1\"}@2"), read("a,b\n\"x,1\"\n", r -> r.quoteChar('\'')));
    }

    @Test
    void aQuoteCharacterOfNulMeansNoQuoting() {
        assertEquals(List.of("{a=\"x, b=1\"}@2"), read("a,b\n\"x,1\"\n", r -> r.quoteChar('\0')));
    }

    @Test
    void anEscapeCharacterEscapesTheQuoteAndTheDelimiter() {
        assertEquals(List.of("{a=x\"y, b=2}@2"), read("a,b\n\"x\\\"y\",2\n", r -> r.escapeChar('\\')));
        assertEquals(List.of("{a=x,y, b=2}@2"), read("a,b\nx\\,y,2\n", r -> r.escapeChar('\\')));
    }

    @Test
    void withoutAnEscapeCharacterABackslashIsOrdinaryAndTheQuoteMustBeDoubled() {
        assertEquals(List.of("{a=x\"y, b=2}@2"), read("a,b\n\"x\"\"y\",2\n"));
        CsvReadException error = assertThrows(CsvReadException.class, () -> read("a,b\n\"x\\\"y\",2\n"));
        assertTrue(error.getMessage().contains("Unterminated quoted CSV field in the record starting at line 2"), error.getMessage());
    }

    @Test
    void strictQuotesDropsWhatIsOutsideTheQuotes() {
        assertEquals(List.of("{a=1, b=2}@2"), read("\"a\",\"b\"\n\"1\",\"2\"\n", r -> r.strictQuotes(true)));
        assertEquals(List.of("{a=1, b=2}@2"), read("\"a\",\"b\"\n\"1\"junk,\"2\"\n", r -> r.strictQuotes(true)));
        assertEquals(List.of("{a=, b=2}@2"), read("\"a\",\"b\"\n1,\"2\"\n", r -> r.strictQuotes(true)));
        assertEquals(List.of("{a=1junk, b=2}@2"), read("a,b\n\"1\"junk,2\n"), "lenient: text after a closing quote joins the value");
    }

    @Test
    void aSpaceBeforeAQuoteIsSkippedByDefault() {
        // hand-written files often have a space after the comma: the value is the quoted text, whole
        assertEquals(List.of("{a=1, b=x,y}@2"), read("a,b\n1, \"x,y\"\n"));
        assertEquals(List.of("{a=1, b=x}@2"), read("a,b\n1, \"x\"\n"));
        assertEquals(List.of("{a=1, b=x,y, c=3}@2"), read("a,b,c\n1, \"x,y\", \"3\"\n"));
        assertEquals(List.of("{a=  x, b=  y}@2"), read("a,b\n  x,  y\n"), "unquoted values keep their spaces");
    }

    @Test
    void leadingWhiteSpaceIsIgnoredBeforeAQuotedValueOnlyWhenAsked() {
        assertEquals(List.of("{a=1, b=x,y}@2"), read("a,b\n1, \"x,y\"\n", r -> r.ignoreLeadingWhiteSpace(true)));
        assertEquals(List.of("{a=1, b=x}@2"), read("a,b\n1, \"x\"\n", r -> r.ignoreLeadingWhiteSpace(true)));
        assertEquals(List.of("{a=1, b=x,y, c= 3}@2"), read("a,b,c\n1, \"x,y\", 3\n", r -> r.ignoreLeadingWhiteSpace(true)));
        // whitespace around values that are not quoted is always kept
        assertEquals(List.of("{a=  x, b=  y}@2"), read("a,b\n  x,  y\n", r -> r.ignoreLeadingWhiteSpace(true)));
    }

    // ---- the delimiter, the quote and the escape character have to differ

    @Test
    void charactersThatCannotBeToldApartAreRefusedBeforeTheFileIsTouched() {
        record Bad(String what, UnaryOperator<CsvReader<Map<String, String>>> options, String message) {
        }
        for (Bad bad : List.of(
                new Bad("delimiter = quote", r -> r.delimiter('"'), "delimiter and the quote character must differ"),
                new Bad("delimiter = escape", r -> r.delimiter('\\').escapeChar('\\'), "delimiter and the escape character must differ"),
                new Bad("quote = escape", r -> r.escapeChar('"'), "quote and the escape character must differ"),
                new Bad("NUL delimiter", r -> r.delimiter('\0'), "cannot be the NUL character"),
                new Bad("LF delimiter", r -> r.delimiter('\n'), "delimiter cannot be a line break"),
                new Bad("CR delimiter", r -> r.delimiter('\r'), "delimiter cannot be a line break"),
                new Bad("LF quote", r -> r.quoteChar('\n'), "quote character cannot be a line break"),
                new Bad("CR escape", r -> r.escapeChar('\r'), "escape character cannot be a line break"))) {
            IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
                    () -> read("a,b\n1,2\n", bad.options()), bad.what());
            assertTrue(error.getMessage().contains(bad.message()), bad.what() + ": " + error.getMessage());
        }
    }

    @Test
    void validCombinationsAreStillTaken() {
        assertEquals(List.of("{a=1, b=2}@2"), read("a\tb\n1\t2\n", r -> r.delimiter('\t')));
        assertEquals(List.of("{a=1, b=2}@2"), read("a|b\n1|2\n", r -> r.delimiter('|')));
        assertEquals(List.of("{a=1, b=2}@2"), read("a;b\n1;2\n", r -> r.delimiter(';').quoteChar('\'').escapeChar('\\')));
        assertDoesNotThrow(() -> read("a,b\n1,2\n", r -> r.quoteChar('\0').escapeChar('\0')));
    }

    @Test
    void theDetectedDelimiterIsCheckedToo() {
        // the detector picks ';' here, which is also the escape character the caller chose
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class, () ->
                CsvReader.forMap().escapeChar(';').readDetected(new ByteArrayInputStream("a;b\n1;2\n".getBytes(StandardCharsets.UTF_8)), r -> { }));
        assertTrue(error.getMessage().contains("escape character"), error.getMessage());
    }

    // ---- line endings, shapes of files

    @Test
    void anyLineEndingEndsARowAndQuotedLineBreaksAreKept() {
        assertEquals(List.of("{a=1, b=2}@2", "{a=3, b=4}@3"), read("a,b\r1,2\r3,4\r"));
        assertEquals(List.of("{a=1, b=2}@2", "{a=3, b=4}@3"), read("a,b\r\n1,2\r\n3,4"));
        assertEquals(List.of("{a=x\r\ny, b=2}@2"), read("a,b\r\n\"x\r\ny\",2\r\n"));
        assertEquals(List.of("{a=x\ny, b=2}@2"), read("a,b\n\"x\ny\",2\n"));
        assertEquals(List.of("{a=x\ry, b=2}@2"), read("a,b\n\"x\ry\",2\n"));
    }

    @Test
    void aRecordIsAHeaderRowEvenWhenItSpansLines() {
        assertEquals(List.of("{name=ann, city=oslo}@4"), read("\"line1\nline2\"\nname,city\nann,oslo\n", r -> r.headerRowIndex(1)));
        assertEquals(List.of("{na\nme=ann, city=oslo}@3"), read("\"na\nme\",city\nann,oslo\n"));
    }

    @Test
    void rowsAreFilledOrCutToTheHeader() {
        assertEquals(List.of("{a=1, b=2, c=}@2"), read("a,b,c\n1,2\n"));
        assertEquals(List.of("{a=1, b=2}@2"), read("a,b\n1,2,3\n"));
        assertEquals(List.of("{a=1, b=2}@2"), read("a,b,\n1,2,\n"), "a trailing delimiter is an empty column with no name");
    }

    @Test
    void aByteOrderMarkIsNotPartOfTheFirstHeader() {
        assertEquals(List.of("{a=1, b=2}@2"), read("﻿a,b\n1,2\n"));
        assertEquals(List.of(), read("a,b\n"));
        assertEquals(List.of(), read("a,b"));
    }

    @Test
    void anEmptyFileOrOnlyAByteOrderMarkHasNoHeader() {
        assertTrue(assertThrows(CsvReadException.class, () -> read("")).getMessage().contains("empty or missing header"));
        assertTrue(assertThrows(CsvReadException.class, () -> read("﻿")).getMessage().contains("after BOM removal"));
    }

    // ---- what is wrong with the input is said in the message

    @Test
    void aNulCharacterIsNamedInTheMessage() {
        CsvReadException error = assertThrows(CsvReadException.class, () -> read("a,b\n1,\u00002\n"));
        assertEquals("Failed to read CSV: CSV input contains a NUL character", error.getMessage());
    }

    @Test
    void aQuoteThatNeverEndsIsNamedInTheMessageWithTheLine() {
        CsvReadException error = assertThrows(CsvReadException.class, () -> read("a,b\n1,\"2\n3,4\n"));
        assertTrue(error.getMessage().startsWith("Failed to read CSV: Unterminated quoted CSV field in the record starting at line 2"),
                error.getMessage());
    }

    @Test
    void anEncodingMismatchNamesTheCharsetAndHowToSetIt() {
        byte[] korean = "이름,도시\n홍길동,서울\n".getBytes(EUC_KR);
        CsvReadException error = assertThrows(CsvReadException.class, () -> read(korean, r -> r));
        assertTrue(error.getMessage().contains("not valid UTF-8") && error.getMessage().contains("charset(...)"), error.getMessage());
        assertEquals(List.of("{이름=홍길동, 도시=서울}@2"), read(korean, r -> r.charset(EUC_KR)));
    }

    // ---- encodings

    @Test
    void theConfiguredCharsetDecodesTheFile() {
        assertEquals(List.of("{a=1, b=2}@2"),
                read("a,b\n1,2\n".getBytes(StandardCharsets.UTF_16LE), r -> r.charset(StandardCharsets.UTF_16LE)));
        assertEquals(List.of("{a=1, b=2}@2"),
                read("﻿a,b\n1,2\n".getBytes(StandardCharsets.UTF_16BE), r -> r.charset(StandardCharsets.UTF_16BE)));
        assertEquals(List.of("{city=Zürich, n=1}@2"),
                read("city,n\nZürich,1\n".getBytes(StandardCharsets.ISO_8859_1), r -> r.charset(StandardCharsets.ISO_8859_1)));
    }

    // ---- the writer's side of it

    @Test
    void aDelimiterThatCouldNotBeReadBackIsRefusedByTheWriter() {
        for (char bad : new char[]{'"', '\n', '\r', '\0'}) {
            assertThrows(IllegalArgumentException.class, () -> CsvWriter.<String>create().delimiter(bad), "U+" + (int) bad);
        }
        for (char fine : new char[]{',', ';', '\t', '|', ' ', ':'}) {
            assertDoesNotThrow(() -> CsvWriter.<String>create().delimiter(fine), "'" + fine + "'");
        }
    }

    @Test
    void whatTheWriterWritesWithAnyAcceptedDelimiterReadsBack() throws Exception {
        for (char delimiter : new char[]{',', ';', '\t', '|', ' ', ':', 'x'}) {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            CsvWriter.<String>create().bom(false).delimiter(delimiter).column("a", s -> s).column("b", s -> "has " + delimiter + " in it")
                    .write(List.of("v")).writeTo(out);
            assertEquals(List.of("{a=v, b=has " + delimiter + " in it}@2"), read(out.toByteArray(), r -> r.delimiter(delimiter)),
                    "delimiter '" + delimiter + "'");
        }
    }

    @Test
    void readResultsKeepTheirType() {
        List<ReadResult<Map<String, String>>> results = new ArrayList<>();
        CsvReader.forMap().read(new ByteArrayInputStream("a\n1\n".getBytes(StandardCharsets.UTF_8)), results::add);
        assertEquals(1, results.size());
    }
}
