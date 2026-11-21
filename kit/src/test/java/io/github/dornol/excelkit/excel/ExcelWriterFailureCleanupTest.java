package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** A failed write must release the workbook's temp files, whichever phase fails. */
class ExcelWriterFailureCleanupTest {

    private static long tempFileCount() throws IOException {
        Path dir = Path.of(System.getProperty("java.io.tmpdir"), "poifiles");
        if (!Files.exists(dir)) {
            return 0;
        }
        try (Stream<Path> files = Files.list(dir)) {
            return files.count();
        }
    }

    @Test
    void failureWhileStartingTheSheet_releasesTempFilesAndKeepsTheOriginalException() throws IOException {
        long before = tempFileCount();
        for (int i = 0; i < 3; i++) {
            ExcelWriter<String> writer = ExcelWriter.<String>create()
                    .column("A", s -> s)
                    .beforeHeader(ctx -> {
                        throw new IllegalStateException("boom");
                    });
            IllegalStateException error = assertThrows(IllegalStateException.class,
                    () -> writer.write(List.of("x")));
            assertEquals("boom", error.getMessage());
        }
        assertEquals(before, tempFileCount());
    }

    @Test
    void failureWhileWritingRows_releasesTempFiles() throws IOException {
        long before = tempFileCount();
        for (int i = 0; i < 3; i++) {
            ExcelWriter<String> writer = ExcelWriter.<String>create()
                    .column("A", s -> {
                        throw new IllegalStateException("boom");
                    });
            assertThrows(ExcelWriteException.class, () -> writer.write(List.of("x")));
        }
        assertEquals(before, tempFileCount());
    }
}
