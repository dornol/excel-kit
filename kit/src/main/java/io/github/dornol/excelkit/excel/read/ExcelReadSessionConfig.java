package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadColumn;
import io.github.dornol.excelkit.core.read.ReadOptions;
import io.github.dornol.excelkit.core.read.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import java.util.List;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Supplier;

/** Immutable execution configuration consumed by the Excel engine. */
record ExcelReadSessionConfig<T>(@Nullable List<ReadColumn<T>> columns,
        @Nullable Supplier<T> supplier, @Nullable Function<RowData,T> mapper,
        @Nullable Validator validator, int sheetIndex, int headerRowIndex, int headerRows,
        @Nullable String password, boolean countRows, @Nullable Set<String> selectedColumns,
        ReadOptions options) {}
