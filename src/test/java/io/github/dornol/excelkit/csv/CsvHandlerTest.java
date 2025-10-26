package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvHandler} class.
 */
class CsvHandlerTest {

    @TempDir
    Path tempDir;

    @Test
    void constructor_shouldCreateInstanceWithValidParameters() throws IOException {
        // Arrange
        Path tempFile = Files.createFile(tempDir.resolve("test.csv"));
        String testContent = "test,content\nrow1,data1";
        Files.write(tempFile, testContent.getBytes(StandardCharsets.UTF_8));

        // Act
        CsvHandler handler = new CsvHandler(tempDir, tempFile);

        // Assert
        assertNotNull(handler, "Handler should be created with valid parameters");
    }

    @Test
    void consumeOutputStream_shouldWriteContentToOutputStream() throws IOException {
        // Arrange
        Path tempFile = Files.createFile(tempDir.resolve("test.csv"));
        String testContent = "test,content\nrow1,data1";
        Files.write(tempFile, testContent.getBytes(StandardCharsets.UTF_8));
        
        CsvHandler handler = new CsvHandler(tempDir, tempFile);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        // Act
        handler.consumeOutputStream(outputStream);

        // Assert
        String result = outputStream.toString(StandardCharsets.UTF_8);
        assertEquals(testContent, result, "Content written to output stream should match the file content");
    }

    @Test
    void consumeOutputStream_shouldThrowExceptionWhenCalledTwice() throws IOException {
        // Arrange
        Path tempFile = Files.createFile(tempDir.resolve("test.csv"));
        String testContent = "test,content\nrow1,data1";
        Files.write(tempFile, testContent.getBytes(StandardCharsets.UTF_8));
        
        CsvHandler handler = new CsvHandler(tempDir, tempFile);
        ByteArrayOutputStream outputStream1 = new ByteArrayOutputStream();
        ByteArrayOutputStream outputStream2 = new ByteArrayOutputStream();

        // Act & Assert
        handler.consumeOutputStream(outputStream1); // First call should succeed
        
        assertThrows(IllegalStateException.class, () -> {
            handler.consumeOutputStream(outputStream2); // Second call should throw exception
        }, "Second call to consumeOutputStream should throw IllegalStateException");
    }

    @Test
    void consumeOutputStream_shouldDeleteTempFileAfterConsumption() throws IOException {
        // Arrange
        Path tempFile = Files.createFile(tempDir.resolve("test.csv"));
        String testContent = "test,content\nrow1,data1";
        Files.write(tempFile, testContent.getBytes(StandardCharsets.UTF_8));
        
        CsvHandler handler = new CsvHandler(tempDir, tempFile);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        // Act
        handler.consumeOutputStream(outputStream);

        // Assert
        assertFalse(Files.exists(tempFile), "Temp file should be deleted after consumption");
    }
}