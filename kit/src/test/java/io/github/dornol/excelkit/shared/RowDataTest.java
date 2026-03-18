package io.github.dornol.excelkit.shared;

import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link RowData}.
 */
class RowDataTest {

    private RowData createRowData(String[] headers, String[] values) {
        List<String> headerNames = List.of(headers);
        Map<String, Integer> headerIndex = new LinkedHashMap<>();
        for (int i = 0; i < headers.length; i++) {
            headerIndex.putIfAbsent(headers[i], i);
        }
        List<CellData> cells = new java.util.ArrayList<>();
        for (int i = 0; i < values.length; i++) {
            cells.add(new CellData(i, values[i]));
        }
        return new RowData(cells, headerNames, headerIndex);
    }

    @Test
    void get_byHeaderName_shouldReturnCorrectCell() {
        RowData row = createRowData(
                new String[]{"Name", "Age", "City"},
                new String[]{"Alice", "30", "Seoul"});

        assertEquals("Alice", row.get("Name").asString());
        assertEquals(30, row.get("Age").asInt());
        assertEquals("Seoul", row.get("City").asString());
    }

    @Test
    void get_byIndex_shouldReturnCorrectCell() {
        RowData row = createRowData(
                new String[]{"Name", "Age"},
                new String[]{"Alice", "30"});

        assertEquals("Alice", row.get(0).asString());
        assertEquals(30, row.get(1).asInt());
    }

    @Test
    void get_byHeaderName_shouldThrowForMissingHeader() {
        RowData row = createRowData(
                new String[]{"Name", "Age"},
                new String[]{"Alice", "30"});

        var ex = assertThrows(IllegalArgumentException.class, () -> row.get("NonExistent"));
        assertTrue(ex.getMessage().contains("NonExistent"));
        assertTrue(ex.getMessage().contains("Available headers"));
    }

    @Test
    void get_byIndex_shouldReturnEmptyCellForOutOfBounds() {
        RowData row = createRowData(
                new String[]{"Name"},
                new String[]{"Alice"});

        CellData cell = row.get(5);
        assertTrue(cell.isEmpty());
        assertEquals("", cell.asString());
    }

    @Test
    void get_byIndex_negativeIndex_shouldThrow() {
        RowData row = createRowData(
                new String[]{"Name"},
                new String[]{"Alice"});

        assertThrows(IllegalArgumentException.class, () -> row.get(-1));
    }

    @Test
    void get_byHeaderName_shouldReturnEmptyCellWhenRowIsShorterThanHeaders() {
        // Header has 3 columns but row only has 1 cell
        List<String> headerNames = List.of("Name", "Age", "City");
        Map<String, Integer> headerIndex = Map.of("Name", 0, "Age", 1, "City", 2);
        List<CellData> cells = List.of(new CellData(0, "Alice"));

        RowData row = new RowData(cells, headerNames, headerIndex);

        assertEquals("Alice", row.get("Name").asString());
        assertTrue(row.get("Age").isEmpty());
        assertTrue(row.get("City").isEmpty());
    }

    @Test
    void has_shouldReturnTrueForExistingHeader() {
        RowData row = createRowData(
                new String[]{"Name", "Age"},
                new String[]{"Alice", "30"});

        assertTrue(row.has("Name"));
        assertTrue(row.has("Age"));
    }

    @Test
    void has_shouldReturnFalseForMissingHeader() {
        RowData row = createRowData(
                new String[]{"Name"},
                new String[]{"Alice"});

        assertFalse(row.has("NonExistent"));
        assertFalse(row.has(""));
    }

    @Test
    void size_shouldReturnNumberOfCells() {
        RowData row = createRowData(
                new String[]{"A", "B", "C"},
                new String[]{"1", "2", "3"});

        assertEquals(3, row.size());
    }

    @Test
    void size_shouldReturnZeroForEmptyRow() {
        RowData row = new RowData(List.of(), List.of("Name"), Map.of("Name", 0));
        assertEquals(0, row.size());
    }

    @Test
    void headerNames_shouldReturnUnmodifiableList() {
        RowData row = createRowData(
                new String[]{"Name", "Age"},
                new String[]{"Alice", "30"});

        List<String> headers = row.headerNames();
        assertEquals(List.of("Name", "Age"), headers);
        assertThrows(UnsupportedOperationException.class, () -> headers.add("City"));
    }

    @Test
    void duplicateHeaders_shouldUseFirstOccurrence() {
        // Simulate duplicate header: "Name" appears at index 0 and 2
        List<String> headerNames = List.of("Name", "Age", "Name");
        Map<String, Integer> headerIndex = new LinkedHashMap<>();
        headerIndex.put("Name", 0);  // first occurrence
        headerIndex.put("Age", 1);
        // "Name" at index 2 is ignored (putIfAbsent)

        List<CellData> cells = List.of(
                new CellData(0, "First"),
                new CellData(1, "30"),
                new CellData(2, "Second"));

        RowData row = new RowData(cells, headerNames, headerIndex);

        // get("Name") returns first occurrence
        assertEquals("First", row.get("Name").asString());
        // but can still access second by index
        assertEquals("Second", row.get(2).asString());
    }

    @Test
    void cellDataConversions_shouldWorkThroughRowData() {
        RowData row = createRowData(
                new String[]{"Name", "Age", "Active", "Score", "Date"},
                new String[]{"Alice", "30", "true", "99.5", "2025-01-15"});

        assertEquals("Alice", row.get("Name").asString());
        assertEquals(30, row.get("Age").asInt());
        assertTrue(row.get("Active").asBoolean());
        assertEquals(99.5, row.get("Score").asDouble());
        assertEquals(java.time.LocalDate.of(2025, 1, 15), row.get("Date").asLocalDate());
    }

    @Test
    void nullCellValue_shouldBeEmpty() {
        List<CellData> cells = List.of(new CellData(0, null));
        RowData row = new RowData(cells, List.of("Col"), Map.of("Col", 0));

        assertTrue(row.get("Col").isEmpty());
        assertEquals("", row.get("Col").asString());
    }
}
