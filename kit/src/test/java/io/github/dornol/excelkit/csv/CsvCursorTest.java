package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Tests for {@link CsvCursor} class.
 */
class CsvCursorTest {

    @Test
    void constructor_shouldInitializeWithZeroValues() {
        // Act
        CsvCursor cursor = new CsvCursor();

        // Assert
        assertEquals(0, cursor.getRowOfSheet(), "Row of sheet should be initialized to 0");
        assertEquals(0, cursor.getCurrentTotal(), "Current total should be initialized to 0");
    }

    @Test
    void plusRow_shouldIncrementRowOfSheet() {
        // Arrange
        CsvCursor cursor = new CsvCursor();
        int initialRow = cursor.getRowOfSheet();

        // Act
        cursor.plusRow();

        // Assert
        assertEquals(initialRow + 1, cursor.getRowOfSheet(), "plusRow should increment row of sheet by 1");
        assertEquals(0, cursor.getCurrentTotal(), "plusRow should not affect current total");
    }

    @Test
    void plusRow_shouldIncrementMultipleTimes() {
        // Arrange
        CsvCursor cursor = new CsvCursor();

        // Act
        cursor.plusRow();
        cursor.plusRow();
        cursor.plusRow();

        // Assert
        assertEquals(3, cursor.getRowOfSheet(), "plusRow should increment row of sheet by 1 each time");
        assertEquals(0, cursor.getCurrentTotal(), "plusRow should not affect current total");
    }

    @Test
    void initRow_shouldResetRowOfSheet() {
        // Arrange
        CsvCursor cursor = new CsvCursor();
        cursor.plusRow();
        cursor.plusRow();
        
        // Act
        cursor.initRow();

        // Assert
        assertEquals(0, cursor.getRowOfSheet(), "initRow should reset row of sheet to 0");
        assertEquals(0, cursor.getCurrentTotal(), "initRow should not affect current total");
    }

    @Test
    void plusTotal_shouldIncrementCurrentTotal() {
        // Arrange
        CsvCursor cursor = new CsvCursor();
        int initialTotal = cursor.getCurrentTotal();

        // Act
        cursor.plusTotal();

        // Assert
        assertEquals(initialTotal + 1, cursor.getCurrentTotal(), "plusTotal should increment current total by 1");
        assertEquals(0, cursor.getRowOfSheet(), "plusTotal should not affect row of sheet");
    }

    @Test
    void plusTotal_shouldIncrementMultipleTimes() {
        // Arrange
        CsvCursor cursor = new CsvCursor();

        // Act
        cursor.plusTotal();
        cursor.plusTotal();
        cursor.plusTotal();

        // Assert
        assertEquals(3, cursor.getCurrentTotal(), "plusTotal should increment current total by 1 each time");
        assertEquals(0, cursor.getRowOfSheet(), "plusTotal should not affect row of sheet");
    }

    @Test
    void combinedOperations_shouldWorkCorrectly() {
        // Arrange
        CsvCursor cursor = new CsvCursor();

        // Act & Assert
        cursor.plusRow();
        assertEquals(1, cursor.getRowOfSheet(), "Row should be 1 after plusRow");
        assertEquals(0, cursor.getCurrentTotal(), "Total should still be 0");

        cursor.plusTotal();
        assertEquals(1, cursor.getRowOfSheet(), "Row should still be 1");
        assertEquals(1, cursor.getCurrentTotal(), "Total should be 1 after plusTotal");

        cursor.initRow();
        assertEquals(0, cursor.getRowOfSheet(), "Row should be 0 after initRow");
        assertEquals(1, cursor.getCurrentTotal(), "Total should still be 1");

        cursor.plusRow();
        cursor.plusTotal();
        assertEquals(1, cursor.getRowOfSheet(), "Row should be 1 after plusRow");
        assertEquals(2, cursor.getCurrentTotal(), "Total should be 2 after plusTotal");
    }
}