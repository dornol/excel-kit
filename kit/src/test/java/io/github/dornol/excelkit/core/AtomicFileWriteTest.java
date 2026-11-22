package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** Writing to a path must never destroy an existing file or leave a partial one behind. */
class AtomicFileWriteTest {
    @TempDir
    Path dir;

    private List<String> names() throws IOException {
        try (Stream<Path> files = Files.list(dir)) {
            return files.map(path -> path.getFileName().toString()).sorted().toList();
        }
    }

    @Test
    void aSecondWriteOfAConsumedHandlerKeepsTheFirstFile() throws IOException {
        Path csv = dir.resolve("out.csv");
        var handler = CsvWriter.<String>create().column("v", s -> s).write(List.of("x", "y"));
        handler.writeTo(csv);
        byte[] first = Files.readAllBytes(csv);
        assertTrue(first.length > 0);

        assertThrows(RuntimeException.class, () -> handler.writeTo(csv));

        assertArrayEquals(first, Files.readAllBytes(csv));
        assertEquals(List.of("out.csv"), names(), "no temporary files remain");
    }

    @Test
    void aFailureWhileWritingLeavesTheExistingFileAndNoPartialFile() throws IOException {
        Path target = dir.resolve("report.bin");
        Files.writeString(target, "previous content");
        FileHandler failing = out -> {
            try {
                out.write("partial".getBytes(StandardCharsets.UTF_8));
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
            throw new IllegalStateException("boom");
        };

        assertThrows(IllegalStateException.class, () -> failing.writeTo(target));

        assertEquals("previous content", Files.readString(target));
        assertEquals(List.of("report.bin"), names());

        Path fresh = dir.resolve("fresh.bin");
        assertThrows(IllegalStateException.class, () -> failing.writeTo(fresh));
        assertFalse(Files.exists(fresh), "a failed write creates no file");
        assertEquals(List.of("report.bin"), names());
    }

    @Test
    void aSuccessfulWriteReplacesTheTargetAndCreatesNewFiles() throws IOException {
        Path target = dir.resolve("out.bin");
        Files.writeString(target, "old");
        FileHandler ok = out -> {
            try {
                out.write("new content".getBytes(StandardCharsets.UTF_8));
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        };

        ok.writeTo(target);
        ok.writeTo(dir.resolve("created.bin"));

        assertEquals("new content", Files.readString(target));
        assertEquals("new content", Files.readString(dir.resolve("created.bin")));
        assertEquals(List.of("created.bin", "out.bin"), names());
    }

    @Test
    void passwordProtectedWritesAreAtomicToo() throws IOException {
        Path file = dir.resolve("secret.xlsx");
        var handler = ExcelWriter.<String>create().column("v", s -> s).write(List.of("x"));
        handler.writeTo(file, "pw");
        byte[] first = Files.readAllBytes(file);

        var rows = ExcelReader.forMap().password("pw").readAll(file);
        assertEquals("x", rows.get(0).get("v"));

        assertThrows(RuntimeException.class, () -> handler.writeTo(file, "pw"));
        assertArrayEquals(first, Files.readAllBytes(file));

        Path chars = dir.resolve("secret2.xlsx");
        ExcelWriter.<String>create().column("v", s -> s).write(List.of("y")).writeTo(chars, "pw2");
        assertEquals("y", ExcelReader.forMap().password("pw2").readAll(chars).get(0).get("v"));
        assertEquals(List.of("secret.xlsx", "secret2.xlsx"), names());
    }

    private static final class UncheckedIOException extends RuntimeException {
        UncheckedIOException(IOException cause) {
            super(cause);
        }
    }
}
