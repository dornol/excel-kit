package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvColumn} class.
 */
class CsvColumnTest {

    @Test
    void constructor_shouldThrowExceptionWhenNameIsNull() {
        // Arrange
        String name = null;
        CsvRowFunction<String, Object> function = (data, cursor) -> data;

        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            new CsvColumn<>(name, function);
        }, "Constructor should throw IllegalArgumentException when name is null");
    }

    @Test
    void constructor_shouldThrowExceptionWhenFunctionIsNull() {
        // Arrange
        String name = "Column";
        CsvRowFunction<String, Object> function = null;

        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            new CsvColumn<>(name, function);
        }, "Constructor should throw IllegalArgumentException when function is null");
    }

    @Test
    void constructor_shouldCreateInstanceWithValidParameters() {
        // Arrange
        String name = "Column";
        CsvRowFunction<String, Object> function = (data, cursor) -> data;

        // Act
        CsvColumn<String> column = new CsvColumn<>(name, function);

        // Assert
        assertNotNull(column, "Column should be created with valid parameters");
        assertEquals(name, column.getName(), "getName should return the column name");
    }

    @Test
    void applyFunction_shouldReturnFunctionResult() {
        // Arrange
        String name = "Column";
        String testData = "Test Data";
        CsvCursor cursor = new CsvCursor();
        CsvRowFunction<String, Object> function = (data, cursor1) -> data + "-processed";
        CsvColumn<String> column = new CsvColumn<>(name, function);

        // Act
        Object result = column.applyFunction(testData, cursor);

        // Assert
        assertEquals("Test Data-processed", result, "applyFunction should return the result of applying the function");
    }

    @Test
    void applyFunction_shouldReturnNullWhenFunctionThrowsException() {
        // Arrange
        String name = "Column";
        String testData = "Test Data";
        CsvCursor cursor = new CsvCursor();
        CsvRowFunction<String, Object> function = (data, cursor1) -> {
            throw new RuntimeException("Test exception");
        };
        CsvColumn<String> column = new CsvColumn<>(name, function);

        // Act
        Object result = column.applyFunction(testData, cursor);

        // Assert
        assertNull(result, "applyFunction should return null when function throws an exception");
    }

    @Test
    void getName_shouldReturnColumnName() {
        // Arrange
        String name = "Test Column";
        CsvRowFunction<String, Object> function = (data, cursor) -> data;
        CsvColumn<String> column = new CsvColumn<>(name, function);

        // Act
        String result = column.getName();

        // Assert
        assertEquals(name, result, "getName should return the column name");
    }
}