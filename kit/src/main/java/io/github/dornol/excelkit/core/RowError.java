package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Describes a row that failed to parse, map, or validate during read.
 * <p>
 * Delivered via the error callback in
 * {@link AbstractReadHandler#read(java.util.function.Consumer, java.util.function.Consumer)}.
 *
 * @param rowNum    1-based data row ordinal (excludes header rows)
 * @param fileRowNum 1-based physical row number in the source file, or {@code -1} if unknown
 * @param type      the category of failure
 * @param messages  human-readable messages (validation violations or error descriptions);
 *                  never {@code null}, may be empty
 * @param cause     the underlying exception for mapping/conversion failures;
 *                  {@code null} for validation-only failures
 *
 * @author dhkim
 * @since 0.16.12
 */
public record RowError(
        long rowNum,
        long fileRowNum,
        Type type,
        List<String> messages,
        @Nullable Throwable cause
) {

    /**
     * Backward-compatible constructor without a physical file row number.
     */
    public RowError(long rowNum, Type type, List<String> messages, @Nullable Throwable cause) {
        this(rowNum, -1, type, messages, cause);
    }

    /** Category of row-level read error. */
    public enum Type {
        /** Bean Validation (or required-column) constraint failed. */
        VALIDATION,
        /** Mapping or cell type conversion threw an exception. */
        MAPPING
    }
}
