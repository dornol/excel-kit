package io.github.dornol.excelkit.core.io;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class TempResourceContainerEdgeCaseTest {

    @Test
    void close_withOnlyTempFile_shouldNotThrow() throws IOException {
        var dir = TempResourceCreator.createTempDirectory();
        var file = TempResourceCreator.createTempFile(dir, "test", ".tmp");
        var container = new TempResourceContainer();
        TempResourceAccess.setTempFile(container, file);
        // no tempDir set
        container.close();
        assertFalse(java.nio.file.Files.exists(file));
        // cleanup dir
        java.nio.file.Files.deleteIfExists(dir);
    }

    @Test
    void close_withOnlyTempDir_shouldNotThrow() throws IOException {
        var dir = TempResourceCreator.createTempDirectory();
        var container = new TempResourceContainer();
        TempResourceAccess.setTempDir(container, dir);
        // no tempFile set
        container.close();
        assertFalse(java.nio.file.Files.exists(dir));
    }
}
