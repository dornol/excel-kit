package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Represents the result of reading a single row from an Excel file.
 *
 * @param <T>      The type of the parsed row data
 * @param data     The actual parsed object (null on failure)
 * @param success  Indicates whether the row was successfully parsed and validated
 * @param messages Any validation or processing messages (e.g. errors or warnings)
 * @param cause    The underlying exception for mapping-stage failures, if any;
 *                 {@code null} for validation-only failures
 * @param fileRowNum 1-based physical row number in the source file, or {@code -1} if unknown
 * @param cellErrors structured cell-level errors, if available
 * @param rawValues raw source row values, if available
 *
 * @author dhkim
 * @since 2025-07-19
 */
public record ReadResult<T>(
        @Nullable T data,
        boolean success,
        @Nullable List<String> messages,
        @Nullable Throwable cause,
        long fileRowNum,
        List<CellError> cellErrors,
        List<String> rawValues
) {
    public ReadResult {
        cellErrors = cellErrors == null ? List.of() : List.copyOf(cellErrors);
        rawValues = rawValues == null ? List.of() : List.copyOf(rawValues);
    }

    public ReadResult(@Nullable T data, boolean success, @Nullable List<String> messages,
                      @Nullable Throwable cause, long fileRowNum, List<CellError> cellErrors) {
        this(data, success, messages, cause, fileRowNum, cellErrors, List.of());
    }

    /**
     * Backward-compatible constructor with a file row number but without cell errors.
     */
    public ReadResult(@Nullable T data, boolean success, @Nullable List<String> messages,
                      @Nullable Throwable cause, long fileRowNum) {
        this(data, success, messages, cause, fileRowNum, List.of());
    }

    /**
     * Backward-compatible constructor without a file row number.
     *
     * @param data     parsed object (null on failure)
     * @param success  whether parsing/validation succeeded
     * @param messages validation or processing messages
     * @param cause    underlying mapping exception, if any
     */
    public ReadResult(@Nullable T data, boolean success, @Nullable List<String> messages,
                      @Nullable Throwable cause) {
        this(data, success, messages, cause, -1, List.of());
    }

    /**
     * Backward-compatible constructor without a cause.
     *
     * @param data     parsed object (null on failure)
     * @param success  whether parsing/validation succeeded
     * @param messages validation or processing messages
     */
    public ReadResult(@Nullable T data, boolean success, @Nullable List<String> messages) {
        this(data, success, messages, null, -1, List.of());
    }
}
