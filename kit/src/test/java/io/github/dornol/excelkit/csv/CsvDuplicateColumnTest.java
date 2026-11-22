package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class CsvDuplicateColumnTest {

    // ========================================================================
    // CSV duplicate column validation
    // ========================================================================
    @Test
    void csvDuplicateColumn_shouldThrow() {
        var writer = CsvWriter.<String>create()
                .column("Name", s -> s)
                .column("Name", s -> s);
        assertThrows(CsvWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void csvUniqueColumns_shouldNotThrow() {
        assertDoesNotThrow(() ->
                CsvWriter.<String>create()
                        .column("Name", s -> s)
                        .column("Age", s -> s)
                        .write(Stream.of("test"))
                        .writeTo(java.io.OutputStream.nullOutputStream()));
    }
}
