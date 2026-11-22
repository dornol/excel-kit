package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.excel.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

// ================================================================
// 3. CsvWriter.constColumnIf
// ================================================================

@DisplayName("CsvWriter.constColumnIf")
class CsvConstColumnIfTest {

    @Test
    @DisplayName("condition=true adds the constant column")
    void conditionTrue_addsColumn() throws IOException {
        var out = new ByteArrayOutputStream();
        CsvWriter.<String>create()
                .column("Name", s -> s)
                .constColumnIf("Type", true, "USER")
                .write(Stream.of("Alice"))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = csv.split("\r?\n");
        assertEquals("Name,Type", lines[0]);
        assertEquals("Alice,USER", lines[1]);
    }

    @Test
    @DisplayName("condition=false skips the constant column")
    void conditionFalse_skipsColumn() throws IOException {
        var out = new ByteArrayOutputStream();
        CsvWriter.<String>create()
                .column("Name", s -> s)
                .constColumnIf("Type", false, "USER")
                .write(Stream.of("Alice"))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = csv.split("\r?\n");
        assertEquals("Name", lines[0]);
        assertEquals("Alice", lines[1]);
    }

    @Test
    @DisplayName("constColumnIf with null value produces empty cell")
    void conditionTrue_nullValue() throws IOException {
        var out = new ByteArrayOutputStream();
        CsvWriter.<String>create()
                .column("Name", s -> s)
                .constColumnIf("Extra", true, null)
                .write(Stream.of("Alice"))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = csv.split("\r?\n");
        assertEquals("Name,Extra", lines[0]);
        assertEquals("Alice,", lines[1]);
    }
}
