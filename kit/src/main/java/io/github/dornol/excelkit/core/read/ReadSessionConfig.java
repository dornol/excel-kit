package io.github.dornol.excelkit.core.read;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * What every read handler needs to run, whatever the file format: how a row becomes an object, which columns to
 * read, where the header is, and the reader-wide options. A format adds its own settings next to it.
 *
 * @param columns           the positional/header columns of setter mode, {@code null} in mapping mode
 * @param supplier          creates the row object in setter mode ({@code null} in mapping mode)
 * @param mapper            maps a whole row in mapping mode ({@code null} in setter mode)
 * @param validator         the Bean Validation validator, or {@code null}
 * @param headerRowIndex    the 0-based header row
 * @param selectedColumns   the columns a map-mode read keeps, or {@code null} for all
 * @param countRows         whether to count the data rows in a first pass
 * @param options           the reader-wide options
 * @param <T>               the row type
 * @since 0.29.0
 */
public record ReadSessionConfig<T>(@Nullable List<ReadColumn<T>> columns,
                                   @Nullable Supplier<T> supplier, @Nullable Function<RowData, T> mapper,
                                   @Nullable Validator validator, int headerRowIndex,
                                   @Nullable Set<String> selectedColumns, boolean countRows,
                                   ReadOptions options) {
    public ReadSessionConfig {
        if ((supplier == null) == (mapper == null))
            throw new IllegalArgumentException("Exactly one of supplier or mapper is required");
        Objects.requireNonNull(options, "readOptions cannot be null");
    }
}
