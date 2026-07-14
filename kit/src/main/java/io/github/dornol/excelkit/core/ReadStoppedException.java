package io.github.dornol.excelkit.core;

/** Internal control-flow signal used by readWhile for normal early completion. */
public final class ReadStoppedException extends RuntimeException {
    public ReadStoppedException() {
        super(null, null, false, false);
    }
}
