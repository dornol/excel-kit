package io.github.dornol.excelkit.excel;

import java.util.List;

/** Aggregate result of a multi-sheet workbook read. */
public record WorkbookReadReport(List<SheetReadReport> sheets) {
    public WorkbookReadReport {
        sheets = List.copyOf(sheets);
    }

    public boolean successful() {
        return sheets.stream().allMatch(SheetReadReport::successful);
    }

    public SheetReadReport sheet(String name) {
        return sheets.stream().filter(sheet -> sheet.name().equals(name)).findFirst()
                .orElseThrow(() -> new IllegalArgumentException("Sheet not registered: " + name));
    }
}
