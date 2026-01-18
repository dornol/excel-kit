package io.github.dornol.excelkit.shared;

import org.jspecify.annotations.NonNull;

/**
 * Exception thrown when a temporary resource (file or directory) cannot be created.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class TempResourceCreateException extends RuntimeException {

    /**
     * Constructs a new TempResourceCreateException with the specified cause.
     *
     * @param cause The cause of the failure
     */
    public TempResourceCreateException(@NonNull Throwable cause) {
        super(cause);
    }

}
