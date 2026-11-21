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
    void defaultsToLineFeed() {
        assertEquals("a\nx\ny\n", write(CsvWriter.<String>create()));
    }

    @Test
    void usesTheConfiguredSeparatorForEveryRow() {
        assertEquals("a\r\nx\r\ny\r\n", write(CsvWriter.<String>create().lineSeparator("\r\n")));
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
