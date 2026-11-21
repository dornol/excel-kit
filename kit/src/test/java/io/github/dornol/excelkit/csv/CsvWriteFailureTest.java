package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class CsvWriteFailureTest {

    @Test
    void namesTheRowColumnAndCharacterThatCannotBeEncoded() {
        CsvWriter<String> writer = CsvWriter.<String>create()
                .charset(StandardCharsets.US_ASCII).bom(false)
                .column("name", s -> s);
        CsvWriteException e = assertThrows(CsvWriteException.class,
                () -> writer.write(List.of("ok", "café")));
        assertTrue(e.getMessage().contains("column 'name'"), e.getMessage());
        assertTrue(e.getMessage().contains("U+00E9"), e.getMessage());
        assertTrue(e.getMessage().contains("US-ASCII"), e.getMessage());
        assertTrue(e.getMessage().matches(".*at row \\d+.*"), e.getMessage());
    }

    @Test
    void namesTheHeaderWhenItCannotBeEncoded() {
        CsvWriter<String> writer = CsvWriter.<String>create()
                .charset(StandardCharsets.US_ASCII).bom(false)
                .column("이름", s -> s);
        CsvWriteException e = assertThrows(CsvWriteException.class, () -> writer.write(List.of("x")));
        assertTrue(e.getMessage().contains("column '이름'"), e.getMessage());
    }

    @Test
    void stopsReadingRowsAsSoonAsTheFileCannotBeWritten() throws IOException {
        // The temp file is real, so provoke the failure through an unencodable lone surrogate instead.
        int[] consumed = {0};
        CsvWriter<Integer> writer = CsvWriter.<Integer>create().bom(false).column("v", i -> {
            consumed[0]++;
            return i == 3 ? "\ud800" : "row" + i;
        });
        CsvWriteException e = assertThrows(CsvWriteException.class,
                () -> writer.write(IntStream.range(0, 1_000_000).boxed()));
        assertTrue(e.getMessage().contains("Failed to write CSV"), e.getMessage());
        assertTrue(consumed[0] < 100_000, "kept reading rows after the write failed: " + consumed[0]);
    }

    @Test
    void validRowsStillWrite() throws IOException {
        CsvWriter<String> writer = CsvWriter.<String>create()
                .charset(java.nio.charset.Charset.forName("windows-1252")).bom(false)
                .column("name", s -> s);
        java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
        writer.write(List.of("café")).writeTo(out);
        assertEquals("name\r\ncafé\r\n", out.toString("windows-1252"));
    }
}
