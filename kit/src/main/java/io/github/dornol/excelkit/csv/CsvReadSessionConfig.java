package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.read.ReadColumn;
import io.github.dornol.excelkit.core.read.ReadOptions;
import io.github.dornol.excelkit.core.read.RowData;
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
        @Nullable Set<String> selectedColumns, char quoteChar, char escapeChar,
        boolean strictQuotes, boolean ignoreLeadingWhiteSpace, boolean countRows, ReadOptions options) {}
