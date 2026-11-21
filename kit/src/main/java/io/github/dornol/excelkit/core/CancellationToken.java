package io.github.dornol.excelkit.core;

/** Cooperative cancellation signal checked between emitted rows. */
@FunctionalInterface
public interface CancellationToken {
    /** A token that never cancels; the default of a reader. */
    CancellationToken NONE = () -> false;

    /**
     * Whether the read should stop. Called from the reading thread before every row, so it should be
     * cheap and thread-safe (typically a read of a {@code volatile} flag).
     *
     * @return {@code true} to stop reading before the next row
     */
    boolean isCancellationRequested();
}
