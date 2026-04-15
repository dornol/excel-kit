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
 *
 * @author dhkim
 * @since 2025-07-19
 */
public record ReadResult<T>(
        @Nullable T data,
        boolean success,
        @Nullable List<String> messages,
        @Nullable Throwable cause
) {
    /**
     * Backward-compatible constructor without a cause.
     *
     * @param data     parsed object (null on failure)
     * @param success  whether parsing/validation succeeded
     * @param messages validation or processing messages
     */
    public ReadResult(@Nullable T data, boolean success, @Nullable List<String> messages) {
        this(data, success, messages, null);
    }
}
