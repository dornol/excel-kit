package io.github.dornol.excelkit.excel.read;

import java.util.List;

/**
 * Aggregate result of a multi-sheet workbook read.
 *
 * @since 0.23.0
 */
public record WorkbookReadReport(List<SheetReadReport> sheets) {
    public WorkbookReadReport {
        sheets = List.copyOf(sheets);
    }

    /** Returns true when every sheet was read successfully. */
    public boolean successful() {
        return sheets.stream().allMatch(SheetReadReport::successful);
    }

    /**
     * Returns the report of the sheet registered under {@code name}; a sheet registered by index is
     * named {@code #index}.
     *
     * @param name the sheet name
     * @return that sheet's report
     * @throws IllegalArgumentException if no sheet was registered under that name
     */
    public SheetReadReport sheet(String name) {
        return sheets.stream().filter(sheet -> sheet.name().equals(name)).findFirst()
                .orElseThrow(() -> new IllegalArgumentException("Sheet not registered: " + name));
    }
}
