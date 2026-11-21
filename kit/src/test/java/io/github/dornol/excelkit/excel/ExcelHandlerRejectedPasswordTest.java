package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** A null or blank password consumes the handler and releases its workbook. */
class ExcelHandlerRejectedPasswordTest {

    @TempDir
    Path dir;

    private static long tempFileCount() throws IOException {
        Path poi = Path.of(System.getProperty("java.io.tmpdir"), "poifiles");
        if (!Files.exists(poi)) {
            return 0;
        }
        try (Stream<Path> files = Files.list(poi)) {
            return files.count();
        }
    }

    private static ExcelHandler handler() {
        return ExcelWriter.<String>create().column("A", s -> s).write(List.of("x"));
    }

    private void assertRejectedPasswordReleasesTheWorkbook(Consumer<ExcelHandler> rejectedCall) throws IOException {
        long before = tempFileCount();
        ExcelHandler handler = handler();
        assertEquals(before + 1, tempFileCount(), "the handler owns one temp file before it is consumed");

        assertThrows(IllegalArgumentException.class, () -> rejectedCall.accept(handler));

        assertEquals(before, tempFileCount());
        assertThrows(ExcelWriteException.class, () -> handler.writeTo(new ByteArrayOutputStream()),
                "a rejected handler is already consumed");
    }

    @Test
    void blankStringPasswordToStream() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(new ByteArrayOutputStream(), "  "));
    }

    @Test
    void nullStringPasswordToStream() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(new ByteArrayOutputStream(), (String) null));
    }

    @Test
    void blankCharPasswordToStream() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(new ByteArrayOutputStream(), new char[]{' '}));
    }

    @Test
    void emptyCharPasswordToStream() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(new ByteArrayOutputStream(), new char[0]));
    }

    @Test
    void blankStringPasswordToPath() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(dir.resolve("a.xlsx"), ""));
    }

    @Test
    void blankCharPasswordToPath() throws IOException {
        assertRejectedPasswordReleasesTheWorkbook(h -> h.writeTo(dir.resolve("b.xlsx"), new char[]{' ', ' '}));
    }

    @Test
    void passwordConflictStillAllowsPlainWriteOnTheSameHandler() {
        ExcelHandler handler = ExcelWriter.<String>create().password("first").column("A", s -> s)
                .write(List.of("x"));
        assertThrows(IllegalStateException.class, () -> handler.writeTo(new ByteArrayOutputStream(), "second"));
        handler.writeTo(new ByteArrayOutputStream());
    }
}
