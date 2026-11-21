package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CsvLineSeparatorTest {

    private static String write(CsvWriter<String> writer) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.column("a", s -> s).bom(false).write(List.of("x", "y")).writeTo(out);
        return out.toString(StandardCharsets.UTF_8);
    }

    @Test
    void defaultsToCrlfLikeRfc4180() {
        assertEquals("a\r\nx\r\ny\r\n", write(CsvWriter.<String>create()));
    }

    @Test
    void dialectsCarryTheirOwnSeparator() {
        assertEquals("a\r\nx\r\ny\r\n", write(CsvWriter.<String>create().dialect(CsvDialect.RFC4180)));
        assertEquals("a\r\nx\r\ny\r\n", write(CsvWriter.<String>create().dialect(CsvDialect.EXCEL)));
        assertEquals("a\nx\ny\n", write(CsvWriter.<String>create().dialect(CsvDialect.TSV)));
        assertEquals("a\nx\ny\n", write(CsvWriter.<String>create().dialect(CsvDialect.PIPE)));
    }

    @Test
    void separatorSetAfterTheDialectWins() {
        assertEquals("a\nx\ny\n",
                write(CsvWriter.<String>create().dialect(CsvDialect.EXCEL).lineSeparator("\n")));
    }

    @Test
    void usesTheConfiguredSeparatorForEveryRow() {
        assertEquals("a\nx\ny\n", write(CsvWriter.<String>create().lineSeparator("\n")));
        assertEquals("a\rx\ry\r", write(CsvWriter.<String>create().lineSeparator("\r")));
    }

    @Test
    void appliesToLinesWrittenByAfterData() {
        CsvWriter<String> writer = CsvWriter.<String>create().lineSeparator("\r\n")
                .afterData(w -> w.println("total,2"));
        assertEquals("a\r\nx\r\ny\r\ntotal,2\r\n", write(writer));
    }

    @Test
    void rejectsOtherSeparators() {
        assertThrows(IllegalArgumentException.class, () -> CsvWriter.<String>create().lineSeparator(";"));
        assertThrows(NullPointerException.class, () -> CsvWriter.<String>create().lineSeparator(null));
    }
}
