package io.github.dornol.excelkit.shared;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Provides access to all cell data in a single row, indexed by header name or column position.
 * <p>
 * Used with the mapping-based reader API to support immutable object (e.g., Java record)
 * construction from Excel/CSV rows:
 * <pre>{@code
 * ExcelReader.mapping(row -> new PersonRecord(
 *         row.get("Name").asString(),
 *         row.get("Age").asInt()
 * )).build(inputStream).read(result -> { ... });
 * }</pre>
 *
 * @author dhkim
 * @since 0.8.0
 */
public class RowData {
    private final List<CellData> cells;
    private final List<String> headerNames;
    private final Map<String, Integer> headerIndex;

    /**
     * Constructs a RowData instance.
     *
     * @param cells       the cell data for this row
     * @param headerNames the header names from the file
     * @param headerIndex a map from header name to column index (built once per file)
     */
    public RowData(List<CellData> cells, List<String> headerNames, Map<String, Integer> headerIndex) {
        this.cells = cells;
        this.headerNames = headerNames;
        this.headerIndex = headerIndex;
    }

    /**
     * Gets cell data by header name.
     * Returns an empty {@link CellData} if the column exists but the cell is missing in this row.
     *
     * @param headerName the column header name
     * @return the cell data for the specified column
     * @throws IllegalArgumentException if the header name is not found
     */
    public CellData get(String headerName) {
        Integer idx = headerIndex.get(headerName);
        if (idx == null) {
            throw new IllegalArgumentException(
                    "Header '" + headerName + "' not found. Available headers: " + headerNames);
        }
        return getCell(idx);
    }

    /**
     * Gets cell data by column index (0-based).
     * Returns an empty {@link CellData} if the index is out of bounds for this row.
     *
     * @param columnIndex the 0-based column index
     * @return the cell data at the specified index
     */
    public CellData get(int columnIndex) {
        return getCell(columnIndex);
    }

    /**
     * Checks whether a header name exists in this row's header definition.
     *
     * @param headerName the header name to check
     * @return {@code true} if the header exists
     */
    public boolean has(String headerName) {
        return headerIndex.containsKey(headerName);
    }

    /**
     * Returns the number of cells in this row.
     */
    public int size() {
        return cells.size();
    }

    /**
     * Returns the header names from the file.
     */
    public List<String> headerNames() {
        return Collections.unmodifiableList(headerNames);
    }

    private CellData getCell(int index) {
        if (index < 0 || index >= cells.size()) {
            return new CellData(Math.max(0, index), null);
        }
        return cells.get(index);
    }
}
