package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
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
    protected boolean strictHeaders = false;
    protected DuplicateHeaderPolicy duplicateHeaderPolicy = DuplicateHeaderPolicy.FIRST;
    protected @Nullable Set<String> selectedMapColumns;
    protected @Nullable CellConversionConfig cellConversionConfig;
    protected long maxRows = -1;
    protected boolean skipBlankRows;
    protected int stopAtBlankRows;

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

    protected void selectedMapColumns(@Nullable Set<String> selectedColumns) {
        this.selectedMapColumns = selectedColumns == null ? null : new LinkedHashSet<>(selectedColumns);
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
     * Registers a name-based column mapping with header aliases.
     * The first matching alias, in list order, is used.
     */
    public SELF column(List<String> headerAliases, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(List, BiConsumer)");
        columns.add(new ReadColumn<>(headerAliases, setter));
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

    /**
     * Enables strict header validation.
     * In strict mode, positional and index-based column bindings must resolve to
     * an existing header column before any data row is processed.
     */
    public SELF strictHeaders() {
        return strictHeaders(true);
    }

    /**
     * Enables or disables strict header validation.
     */
    public SELF strictHeaders(boolean enabled) {
        this.strictHeaders = enabled;
        return self();
    }

    /**
     * Alias for {@link #strictHeaders()}.
     */
    public SELF requireHeaders() {
        return strictHeaders();
    }

    /**
     * Sets duplicate header handling. Defaults to {@link DuplicateHeaderPolicy#FIRST}.
     */
    public SELF duplicateHeaderPolicy(DuplicateHeaderPolicy policy) {
        this.duplicateHeaderPolicy = java.util.Objects.requireNonNull(policy, "policy cannot be null");
        return self();
    }

    /**
     * Sets reader-scoped conversion settings for {@link CellData}.
     *
     * @since 0.19.0
     */
    public SELF cellConversion(CellConversionConfig config) {
        this.cellConversionConfig = java.util.Objects.requireNonNull(config, "config cannot be null");
        return self();
    }

    /**
     * Configures reader-scoped {@link CellData} conversion settings.
     *
     * @since 0.19.0
     */
    public SELF cellConversion(Consumer<CellConversionConfig.Builder> configurer) {
        CellConversionConfig.Builder builder = CellConversionConfig.builder();
        configurer.accept(builder);
        return cellConversion(builder.build());
    }

    /**
     * Limits the number of non-skipped data rows emitted by this reader.
     *
     * @since 0.19.0
     */
    public SELF maxRows(long maxRows) {
        if (maxRows < 0) {
            throw new IllegalArgumentException("maxRows must be non-negative");
        }
        this.maxRows = maxRows;
        return self();
    }

    /**
     * Skips rows where every cell is blank.
     *
     * @since 0.19.0
     */
    public SELF skipBlankRows() {
        return skipBlankRows(true);
    }

    /**
     * Enables or disables skipping rows where every cell is blank.
     *
     * @since 0.19.0
     */
    public SELF skipBlankRows(boolean enabled) {
        this.skipBlankRows = enabled;
        return self();
    }

    /**
     * Stops reading after the given number of consecutive blank data rows.
     * Pass {@code 0} to disable.
     *
     * @since 0.19.0
     */
    public SELF stopAtBlankRows(int count) {
        if (count < 0) {
            throw new IllegalArgumentException("count must be non-negative");
        }
        this.stopAtBlankRows = count;
        return self();
    }
}
