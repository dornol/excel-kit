package io.github.dornol.excelkit.core.read;

/**
 * Controls how readers resolve duplicate header names.
 *
 * @since 0.18.0
 */
public enum DuplicateHeaderPolicy {
    /** Use the first occurrence of a duplicated header name. */
    FIRST,
    /** Use the last occurrence of a duplicated header name. */
    LAST,
    /**
     * Fail fast when a duplicated header name is detected, before any data row is read.
     * This is the default; blank headers are ignored.
     */
    FAIL
}
