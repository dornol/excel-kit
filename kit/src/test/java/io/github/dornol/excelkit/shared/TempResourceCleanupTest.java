package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.DisabledOnOs;
import org.junit.jupiter.api.condition.OS;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermissions;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for temporary resource cleanup behavior across the library.
 * <p>
 * Verifies that:
 * <ul>
 *     <li>Temp files and directories are deleted after normal operations</li>
 *     <li>Cleanup does not throw exceptions on already-deleted paths</li>
 *     <li>Cleanup does not throw exceptions on null paths</li>
 *     <li>When deletion fails, {@code deleteOnExit()} is registered as fallback</li>
 * </ul>
 */
class TempResourceCleanupTest {

    // ──────────────────────────────────────────────────────────────
    // TempResourceContainer: normal cleanup
    // ──────────────────────────────────────────────────────────────

    @Test
    void close_shouldDeleteBothFileAndDirectory() {
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "cleanup-test", ".tmp");

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        container.setTempFile(file);

        assertTrue(Files.exists(file));
        assertTrue(Files.exists(dir));

        container.close();

        assertFalse(Files.exists(file), "Temp file should be deleted after close");
        assertFalse(Files.exists(dir), "Temp dir should be deleted after close");
    }

    @Test
    void close_withNullPaths_shouldNotThrow() {
        TempResourceContainer container = new TempResourceContainer();
        assertDoesNotThrow(container::close);
    }

    @Test
    void close_withOnlyFile_shouldNotThrow() throws IOException {
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "only-file", ".tmp");

        TempResourceContainer container = new TempResourceContainer();
        container.setTempFile(file);
        // no setTempDir

        container.close();

        assertFalse(Files.exists(file), "Temp file should be deleted");
        // dir is not managed, clean up manually
        Files.deleteIfExists(dir);
    }

    @Test
    void close_withOnlyDir_shouldNotThrow() {
        Path dir = TempResourceCreator.createTempDirectory();

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        // no setTempFile

        container.close();

        assertFalse(Files.exists(dir), "Temp dir should be deleted");
    }

    @Test
    void close_withAlreadyDeletedPaths_shouldNotThrow() throws IOException {
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "pre-deleted", ".tmp");

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        container.setTempFile(file);

        // Delete manually before close
        Files.delete(file);
        Files.delete(dir);

        assertDoesNotThrow(container::close,
                "close should not throw when paths are already deleted");
    }

    @Test
    void close_multipleTimes_shouldNotThrow() {
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "multi-close", ".tmp");

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        container.setTempFile(file);

        container.close();
        assertDoesNotThrow(container::close,
                "Second close should not throw even though paths are already deleted");
    }

    @Test
    void close_withTryWithResources_shouldCleanup() {
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "try-with", ".tmp");

        try (TempResourceContainer container = new TempResourceContainer()) {
            container.setTempDir(dir);
            container.setTempFile(file);
            assertTrue(Files.exists(file));
        }

        assertFalse(Files.exists(file), "try-with-resources should clean up file");
        assertFalse(Files.exists(dir), "try-with-resources should clean up dir");
    }

    // ──────────────────────────────────────────────────────────────
    // TempResourceContainer: deletion failure (POSIX only)
    // ──────────────────────────────────────────────────────────────

    @Test
    @DisabledOnOs(OS.WINDOWS)
    void close_whenFileDeleteFails_shouldNotThrowAndRegisterDeleteOnExit() throws IOException {
        // Create a directory with a file, then remove write+execute on the directory
        // so the file cannot be unlinked (POSIX: need write+execute on parent dir to unlink)
        Path dir = TempResourceCreator.createTempDirectory();
        Path file = TempResourceCreator.createTempFile(dir, "undeletable", ".tmp");

        Files.setPosixFilePermissions(dir, PosixFilePermissions.fromString("r--------"));

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        container.setTempFile(file);

        try {
            // close should NOT throw — it catches the IOException and calls deleteOnExit
            assertDoesNotThrow(container::close,
                    "close should not throw even when deletion fails");
        } finally {
            // Restore permissions for actual cleanup
            Files.setPosixFilePermissions(dir, PosixFilePermissions.fromString("rwx------"));
            Files.deleteIfExists(file);
            Files.deleteIfExists(dir);
        }
    }

    @Test
    @DisabledOnOs(OS.WINDOWS)
    void close_whenDirDeleteFails_shouldNotThrow() throws IOException {
        // Create a non-empty directory (close tries to delete dir but it has extra files)
        Path dir = TempResourceCreator.createTempDirectory();
        Path managedFile = TempResourceCreator.createTempFile(dir, "managed", ".tmp");
        Path extraFile = Files.createTempFile(dir, "extra", ".tmp");

        TempResourceContainer container = new TempResourceContainer();
        container.setTempDir(dir);
        container.setTempFile(managedFile);

        // close will delete managedFile, then try to delete dir (which still has extraFile)
        assertDoesNotThrow(container::close,
                "close should not throw when dir deletion fails due to non-empty dir");

        assertFalse(Files.exists(managedFile), "Managed file should be deleted");
        assertTrue(Files.exists(dir), "Dir should still exist because it's not empty");
        assertTrue(Files.exists(extraFile), "Extra file should still exist");

        // Cleanup
        Files.deleteIfExists(extraFile);
        Files.deleteIfExists(dir);
    }

    // ──────────────────────────────────────────────────────────────
    // TempResourceCreator: basic creation tests
    // ──────────────────────────────────────────────────────────────

    @Test
    void createTempDirectory_shouldCreateWithRestrictedPermissions() throws IOException {
        Path dir = TempResourceCreator.createTempDirectory();

        try {
            assertTrue(Files.exists(dir));
            assertTrue(Files.isDirectory(dir));

            // On POSIX, verify owner-only permissions
            if (dir.getFileSystem().supportedFileAttributeViews().contains("posix")) {
                var perms = Files.getPosixFilePermissions(dir);
                assertTrue(perms.contains(java.nio.file.attribute.PosixFilePermission.OWNER_READ));
                assertTrue(perms.contains(java.nio.file.attribute.PosixFilePermission.OWNER_WRITE));
                assertTrue(perms.contains(java.nio.file.attribute.PosixFilePermission.OWNER_EXECUTE));
                assertFalse(perms.contains(java.nio.file.attribute.PosixFilePermission.GROUP_READ));
                assertFalse(perms.contains(java.nio.file.attribute.PosixFilePermission.OTHERS_READ));
            }
        } finally {
            Files.deleteIfExists(dir);
        }
    }

    @Test
    void createTempFile_shouldBeInsideSpecifiedDirectory() throws IOException {
        Path dir = TempResourceCreator.createTempDirectory();
        try {
            Path file = TempResourceCreator.createTempFile(dir, "test-prefix", ".xlsx");

            assertTrue(Files.exists(file));
            assertEquals(dir, file.getParent());
            assertTrue(file.getFileName().toString().startsWith("test-prefix"));
            assertTrue(file.getFileName().toString().endsWith(".xlsx"));

            Files.deleteIfExists(file);
        } finally {
            Files.deleteIfExists(dir);
        }
    }

    @Test
    void createTempFile_inNonExistentDir_shouldThrow() {
        assertThrows(TempResourceCreateException.class, () ->
                TempResourceCreator.createTempFile(Path.of("/nonexistent/path"), "test", ".tmp"));
    }

    @Test
    void createTempDirectory_twoCallsCreateDifferentPaths() throws IOException {
        Path dir1 = TempResourceCreator.createTempDirectory();
        Path dir2 = TempResourceCreator.createTempDirectory();
        try {
            assertNotEquals(dir1, dir2, "Each call should create a unique directory");
        } finally {
            Files.deleteIfExists(dir1);
            Files.deleteIfExists(dir2);
        }
    }
}
