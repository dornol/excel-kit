package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermissions;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

class AtomicFileWriterTest {

    @Test
    void replacingAFileKeepsItsPermissions(@TempDir Path dir) throws Exception {
        assumeTrue(dir.getFileSystem().supportedFileAttributeViews().contains("posix"));
        Path target = dir.resolve("shared.txt");
        Files.writeString(target, "old");
        var mode = PosixFilePermissions.fromString("rw-rw-r--");
        Files.setPosixFilePermissions(target, mode);

        AtomicFileWriter.write(target, out -> out.write("new".getBytes(StandardCharsets.UTF_8)));

        assertEquals("new", Files.readString(target));
        assertEquals(mode, Files.getPosixFilePermissions(target));
    }

    @Test
    void aFailedWriteLeavesTheTargetAndNoTempFile(@TempDir Path dir) throws Exception {
        Path target = dir.resolve("keep.txt");
        Files.writeString(target, "old");

        assertThrows(java.io.IOException.class, () -> AtomicFileWriter.write(target, out -> {
            out.write(1);
            throw new java.io.IOException("boom");
        }));

        assertEquals("old", Files.readString(target));
        try (var files = Files.list(dir)) {
            assertEquals(1, files.count());
        }
    }
}
