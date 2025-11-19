package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

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
        } catch (IOException _) {
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
        } catch (IOException _) {
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
}