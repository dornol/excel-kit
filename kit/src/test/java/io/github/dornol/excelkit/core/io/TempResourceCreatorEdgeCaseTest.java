package io.github.dornol.excelkit.core.io;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class TempResourceCreatorEdgeCaseTest {

    @Test
    void createTempFile_multipleCalls_createDistinctFiles() throws IOException {
        var dir = TempResourceCreator.createTempDirectory();
        try {
            var f1 = TempResourceCreator.createTempFile(dir, "a", ".tmp");
            var f2 = TempResourceCreator.createTempFile(dir, "a", ".tmp");
            assertNotEquals(f1, f2);
            assertTrue(java.nio.file.Files.exists(f1));
            assertTrue(java.nio.file.Files.exists(f2));
        } finally {
            try (var files = java.nio.file.Files.walk(dir)) {
                files.sorted(java.util.Comparator.reverseOrder())
                        .forEach(p -> { try { java.nio.file.Files.delete(p); } catch (Exception ignored) {} });
            }
        }
    }
}
