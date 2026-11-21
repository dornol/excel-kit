package io.github.dornol.excelkit.core;

/** Internal control-flow signal used by readWhile for normal early completion. */
public final class ReadStoppedException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    public ReadStoppedException() {
        super(null, null, false, false);
    }
}
