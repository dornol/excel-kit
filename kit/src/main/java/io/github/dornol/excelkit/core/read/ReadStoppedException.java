package io.github.dornol.excelkit.core.read;

/**
 * Internal control-flow signal used by readWhile for normal early completion.
 *
 * @since 0.20.0
 */
public final class ReadStoppedException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    /** Creates the signal; it carries no message and no stack trace. */
    public ReadStoppedException() {
        super(null, null, false, false);
    }
}
