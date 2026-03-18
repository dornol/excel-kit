package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link TempResourceCreator} class.
 */
class TempResourceCreatorTest {

    @Test
    void createTempDirectory_shouldCreateDirectory() {
        // Act
        Path tempDir = TempResourceCreator.createTempDirectory();

        // Assert
        assertTrue(Files.exists(tempDir), "Temp directory should exist");
        assertTrue(Files.isDirectory(tempDir), "Created path should be a directory");

        // Cleanup
        try {
            Files.delete(tempDir);
        } catch (IOException e) {
            // Ignore cleanup errors in tests
        }
    }

    @Test
    void createTempFile_shouldCreateFileInDirectory() {
        // Arrange
        Path tempDir = TempResourceCreator.createTempDirectory();
        String prefix = "test";
        String suffix = ".tmp";

        // Act
        Path tempFile = TempResourceCreator.createTempFile(tempDir, prefix, suffix);

        // Assert
        assertTrue(Files.exists(tempFile), "Temp file should exist");
        assertTrue(Files.isRegularFile(tempFile), "Created path should be a regular file");
        assertTrue(tempFile.getFileName().toString().startsWith(prefix), 
                "File name should start with the specified prefix");
        assertTrue(tempFile.getFileName().toString().endsWith(suffix), 
                "File name should end with the specified suffix");
        assertEquals(tempDir, tempFile.getParent(), 
                "File should be created in the specified directory");

        // Cleanup
        try {
            Files.delete(tempFile);
            Files.delete(tempDir);
        } catch (IOException e) {
            // Ignore cleanup errors in tests
        }
    }

    @Test
    void createTempFile_shouldThrowExceptionWhenDirectoryDoesNotExist() {
        // Arrange
        Path nonExistentDir = Path.of("non-existent-dir");
        String prefix = "test";
        String suffix = ".tmp";

        // Act & Assert
        assertThrows(TempResourceCreateException.class, () -> {
            TempResourceCreator.createTempFile(nonExistentDir, prefix, suffix);
        }, "Should throw TempResourceCreateException when directory doesn't exist");
    }

    @Nested
    class PlatformSpecificTests {

        @Test
        void createTempDirectory_onPosix_shouldHaveRestrictedPermissions() throws IOException {
            boolean isPosix = FileSystems.getDefault().supportedFileAttributeViews().contains("posix");
            if (!isPosix) {
                return; // Skip on Windows
            }
            Path tempDir = TempResourceCreator.createTempDirectory();
            try {
                Set<PosixFilePermission> perms = Files.getPosixFilePermissions(tempDir);
                assertTrue(perms.contains(PosixFilePermission.OWNER_READ));
                assertTrue(perms.contains(PosixFilePermission.OWNER_WRITE));
                assertTrue(perms.contains(PosixFilePermission.OWNER_EXECUTE));
                assertFalse(perms.contains(PosixFilePermission.GROUP_READ));
                assertFalse(perms.contains(PosixFilePermission.OTHERS_READ));
            } finally {
                Files.delete(tempDir);
            }
        }

        @Test
        void createTempDirectory_multipleCalls_createDistinctDirs() throws IOException {
            Path dir1 = TempResourceCreator.createTempDirectory();
            Path dir2 = TempResourceCreator.createTempDirectory();
            try {
                assertNotEquals(dir1, dir2);
                assertTrue(Files.exists(dir1));
                assertTrue(Files.exists(dir2));
            } finally {
                Files.delete(dir1);
                Files.delete(dir2);
            }
        }
    }
}