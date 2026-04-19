package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * Shared reader configuration for {@link io.github.dornol.excelkit.excel.ExcelReader}
 * and {@link io.github.dornol.excelkit.csv.CsvReader}.
 * <p>
 * Contains column registration, progress, validation, and skip logic.
 * Format-specific configuration (sheet index, delimiter, charset, etc.) stays in subclasses.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete reader type, for fluent chaining
 * @author dhkim
 * @since 0.17.0
 */
@SuppressWarnings("unchecked")
public abstract class AbstractReader<T, SELF extends AbstractReader<T, SELF>> {

    protected final List<ReadColumn<T>> columns = new ArrayList<>();
    protected final @Nullable Supplier<T> instanceSupplier;
    protected final @Nullable Function<RowData, T> rowMapper;
    protected final @Nullable Validator validator;
    protected int headerRowIndex = 0;
    protected @Nullable ProgressCallback progressCallback;
    protected int progressInterval;
    protected boolean mapMode = false;

    protected AbstractReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this.instanceSupplier = java.util.Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.rowMapper = null;
        this.validator = validator;
    }

    protected AbstractReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        this.instanceSupplier = null;
        this.rowMapper = java.util.Objects.requireNonNull(rowMapper, "rowMapper cannot be null");
        this.validator = validator;
    }

    private SELF self() {
        return (SELF) this;
    }

    protected void requireNotMapMode(String method) {
        if (mapMode) {
            throw new IllegalStateException(
                    method + " cannot be called on a forMap() reader; "
                            + "map mode auto-discovers columns from the header row");
        }
    }

    /**
     * Sets the zero-based row index of the header row.
     */
    public SELF headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return self();
    }

    void addColumn(ReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Registers a positional column mapping.
     */
    public SELF column(BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(BiConsumer)");
        columns.add(new ReadColumn<>(setter));
        return self();
    }

    /**
     * Registers a name-based column mapping.
     */
    public SELF column(String headerName, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(String, BiConsumer)");
        columns.add(new ReadColumn<>(headerName, setter));
        return self();
    }

    /**
     * Registers an index-based column mapping.
     */
    public SELF columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        requireNotMapMode("columnAt(int, BiConsumer)");
        columns.add(new ReadColumn<>(null, columnIndex, setter));
        return self();
    }

    /**
     * Marks the last registered column as required.
     */
    public SELF required() {
        if (columns.isEmpty()) {
            throw new IllegalStateException("required() must be called after column()");
        }
        int lastIndex = columns.size() - 1;
        columns.set(lastIndex, columns.get(lastIndex).required());
        return self();
    }

    /**
     * Skips one positional column.
     */
    public SELF skipColumn() {
        requireNotMapMode("skipColumn()");
        columns.add(new ReadColumn<>((instance, cellData) -> {}));
        return self();
    }

    /**
     * Skips the specified number of positional columns.
     */
    public SELF skipColumns(int count) {
        requireNotMapMode("skipColumns(int)");
        if (count < 0) {
            throw new IllegalArgumentException("skipColumns count must be non-negative");
        }
        for (int i = 0; i < count; i++) {
            columns.add(new ReadColumn<>((instance, cellData) -> {}));
        }
        return self();
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     */
    public SELF onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return self();
    }
}
