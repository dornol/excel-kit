package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link TempResourceContainer} class.
 */
class TempResourceContainerTest {

    @Test
    void getTempDir_shouldReturnSetValue() {
        // Arrange
        TempResourceContainer container = new TempResourceContainer();
        Path tempDir = TempResourceCreator.createTempDirectory();
        
        // Act
        container.setTempDir(tempDir);
        Path result = container.getTempDir();
        
        // Assert
        assertEquals(tempDir, result, "getTempDir should return the value set by setTempDir");
        
        // Cleanup
        try {
            Files.delete(tempDir);
        } catch (IOException _) {
            // Ignore cleanup errors in tests
        }
    }
    
    @Test
    void getTempFile_shouldReturnSetValue() {
        // Arrange
        TempResourceContainer container = new TempResourceContainer();
        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "test", ".tmp");
        
        // Act
        container.setTempFile(tempFile);
        Path result = container.getTempFile();
        
        // Assert
        assertEquals(tempFile, result, "getTempFile should return the value set by setTempFile");
        
        // Cleanup
        try {
            Files.delete(tempFile);
            Files.delete(tempDir);
        } catch (IOException _) {
            // Ignore cleanup errors in tests
        }
    }
    
    @Test
    void close_shouldDeleteTempFileAndDirectory() {
        // Arrange
        TempResourceContainer container = new TempResourceContainer();
        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "test", ".tmp");
        
        container.setTempDir(tempDir);
        container.setTempFile(tempFile);
        
        assertTrue(Files.exists(tempFile), "Temp file should exist before close");
        assertTrue(Files.exists(tempDir), "Temp directory should exist before close");
        
        // Act
        container.close();
        
        // Assert
        assertFalse(Files.exists(tempFile), "Temp file should be deleted after close");
        assertFalse(Files.exists(tempDir), "Temp directory should be deleted after close");
    }
    
    @Test
    void close_shouldHandleNullPaths() {
        // Arrange
        TempResourceContainer container = new TempResourceContainer();
        
        // Act & Assert
        assertDoesNotThrow(container::close,
                "close should not throw exception when paths are null");
    }
    
    @Test
    void close_shouldHandleNonExistentPaths() throws IOException {
        // Arrange
        TempResourceContainer container = new TempResourceContainer();
        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "test", ".tmp");
        
        container.setTempDir(tempDir);
        container.setTempFile(tempFile);
        
        // Delete the files manually before calling close
        Files.delete(tempFile);
        Files.delete(tempDir);
        
        // Act & Assert
        assertDoesNotThrow(container::close,
                "close should not throw exception when paths don't exist");
    }
    
    @Test
    void autoCloseable_shouldWorkWithTryWithResources() {
        // Arrange
        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "test", ".tmp");
        
        // Act
        try (TempResourceContainer container = new TempResourceContainer()) {
            container.setTempDir(tempDir);
            container.setTempFile(tempFile);
            
            assertTrue(Files.exists(tempFile), "Temp file should exist inside try block");
            assertTrue(Files.exists(tempDir), "Temp directory should exist inside try block");
        }
        
        // Assert
        assertFalse(Files.exists(tempFile), "Temp file should be deleted after try-with-resources");
        assertFalse(Files.exists(tempDir), "Temp directory should be deleted after try-with-resources");
    }
}