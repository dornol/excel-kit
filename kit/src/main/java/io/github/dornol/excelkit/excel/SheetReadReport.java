package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadSummary;
import io.github.dornol.excelkit.core.RowError;
import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Result of processing one registered workbook sheet.
 *
 * @param name   the sheet name, or {@code #index} when it was registered by index
 * @param index  the 0-based sheet index, or {@code -1} when a sheet registered by name was not found
 */
public record SheetReadReport(String name, int index, ReadSummary summary,
                              List<RowError> errors, boolean errorsTruncated,
                              @Nullable Throwable failure) {
    public SheetReadReport {
        if (index < -1) throw new IllegalArgumentException("index must be non-negative or -1 for an unresolved sheet");
        if (name == null || name.isBlank()) throw new IllegalArgumentException("name cannot be blank");
        java.util.Objects.requireNonNull(summary, "summary cannot be null");
        errors = List.copyOf(errors);
    }

    public boolean successful() {
        return failure == null && errors.isEmpty() && !errorsTruncated;
    }
}
