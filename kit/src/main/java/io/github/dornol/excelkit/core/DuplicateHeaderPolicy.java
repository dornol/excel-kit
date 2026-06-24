package io.github.dornol.excelkit.core;

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
    /** Fail fast when a duplicated header name is detected. */
    FAIL
}
