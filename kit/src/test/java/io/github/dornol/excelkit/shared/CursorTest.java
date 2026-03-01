package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Tests for {@link Cursor} class.
 */
class CursorTest {

    @Test
    void constructor_shouldInitializeWithZeroValues() {
        // Act
        Cursor cursor = new Cursor();

        // Assert
        assertEquals(0, cursor.getRowOfSheet(), "Row of sheet should be initialized to 0");
        assertEquals(0, cursor.getCurrentTotal(), "Current total should be initialized to 0");
    }

    @Test
    void constructorWithBaseRow_shouldInitializeWithBaseRow() {
        // Act
        Cursor cursor = new Cursor(2);

        // Assert
        assertEquals(2, cursor.getRowOfSheet(), "Row of sheet should be initialized to baseRow");
        assertEquals(0, cursor.getCurrentTotal(), "Current total should be initialized to 0");
    }

    @Test
    void plusRow_shouldIncrementRowOfSheet() {
        // Arrange
        Cursor cursor = new Cursor();
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
        Cursor cursor = new Cursor();

        // Act
        cursor.plusRow();
        cursor.plusRow();
        cursor.plusRow();

        // Assert
        assertEquals(3, cursor.getRowOfSheet(), "plusRow should increment row of sheet by 1 each time");
        assertEquals(0, cursor.getCurrentTotal(), "plusRow should not affect current total");
    }

    @Test
    void initRow_shouldResetRowOfSheetToBaseRow() {
        // Arrange
        Cursor cursor = new Cursor();
        cursor.plusRow();
        cursor.plusRow();

        // Act
        cursor.initRow();

        // Assert
        assertEquals(0, cursor.getRowOfSheet(), "initRow should reset row of sheet to baseRow (0)");
        assertEquals(0, cursor.getCurrentTotal(), "initRow should not affect current total");
    }

    @Test
    void initRow_withBaseRow_shouldResetToBaseRow() {
        // Arrange
        Cursor cursor = new Cursor(2);
        cursor.plusRow();
        cursor.plusRow();

        // Act
        cursor.initRow();

        // Assert
        assertEquals(2, cursor.getRowOfSheet(), "initRow should reset row of sheet to baseRow (2)");
    }

    @Test
    void plusTotal_shouldIncrementCurrentTotal() {
        // Arrange
        Cursor cursor = new Cursor();
        long initialTotal = cursor.getCurrentTotal();

        // Act
        cursor.plusTotal();

        // Assert
        assertEquals(initialTotal + 1, cursor.getCurrentTotal(), "plusTotal should increment current total by 1");
        assertEquals(0, cursor.getRowOfSheet(), "plusTotal should not affect row of sheet");
    }

    @Test
    void plusTotal_shouldIncrementMultipleTimes() {
        // Arrange
        Cursor cursor = new Cursor();

        // Act
        cursor.plusTotal();
        cursor.plusTotal();
        cursor.plusTotal();

        // Assert
        assertEquals(3, cursor.getCurrentTotal(), "plusTotal should increment current total by 1 each time");
        assertEquals(0, cursor.getRowOfSheet(), "plusTotal should not affect row of sheet");
    }

    @Test
    void getRowOfSheet_shouldReturnCurrentRowValue() {
        // Arrange
        Cursor cursor = new Cursor();

        // Act & Assert
        assertEquals(0, cursor.getRowOfSheet(), "Initial row value should be 0");

        cursor.plusRow();
        assertEquals(1, cursor.getRowOfSheet(), "Row value should be 1 after plusRow");

        cursor.plusRow();
        assertEquals(2, cursor.getRowOfSheet(), "Row value should be 2 after second plusRow");

        cursor.initRow();
        assertEquals(0, cursor.getRowOfSheet(), "Row value should be 0 after initRow");
    }

    @Test
    void getCurrentTotal_shouldReturnCurrentTotalValue() {
        // Arrange
        Cursor cursor = new Cursor();

        // Act & Assert
        assertEquals(0, cursor.getCurrentTotal(), "Initial total value should be 0");

        cursor.plusTotal();
        assertEquals(1, cursor.getCurrentTotal(), "Total value should be 1 after plusTotal");

        cursor.plusTotal();
        assertEquals(2, cursor.getCurrentTotal(), "Total value should be 2 after second plusTotal");
    }

    @Test
    void combinedOperations_shouldWorkCorrectly() {
        // Arrange
        Cursor cursor = new Cursor();

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

    @Test
    void simulateSheetRollover_shouldTrackCorrectly() {
        // Arrange
        Cursor cursor = new Cursor();

        // Act & Assert - First sheet
        for (int i = 0; i < 5; i++) {
            cursor.plusRow();
            cursor.plusTotal();
        }
        assertEquals(5, cursor.getRowOfSheet(), "Row should be 5 after 5 rows in first sheet");
        assertEquals(5, cursor.getCurrentTotal(), "Total should be 5 after 5 rows total");

        // Act & Assert - Sheet rollover
        cursor.initRow(); // Simulate new sheet creation
        assertEquals(0, cursor.getRowOfSheet(), "Row should be 0 after sheet rollover");
        assertEquals(5, cursor.getCurrentTotal(), "Total should still be 5 after sheet rollover");

        // Act & Assert - Second sheet
        for (int i = 0; i < 3; i++) {
            cursor.plusRow();
            cursor.plusTotal();
        }
        assertEquals(3, cursor.getRowOfSheet(), "Row should be 3 after 3 rows in second sheet");
        assertEquals(8, cursor.getCurrentTotal(), "Total should be 8 after 8 rows total");
    }
}
