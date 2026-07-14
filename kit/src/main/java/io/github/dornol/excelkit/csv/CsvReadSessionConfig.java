package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.*;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import java.nio.charset.Charset;
import java.util.List;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Supplier;

/** Immutable execution configuration consumed by the CSV engine. */
record CsvReadSessionConfig<T>(@Nullable List<ReadColumn<T>> columns,
        @Nullable Supplier<T> supplier, @Nullable Function<RowData,T> mapper,
        @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
        int progressInterval, @Nullable ProgressCallback progressCallback,
        @Nullable Set<String> selectedColumns, char quoteChar, char escapeChar,
        boolean strictQuotes, boolean ignoreLeadingWhiteSpace, ReadOptions options) {}
