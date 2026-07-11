package io.github.dornol.excelkit.core;

/** Cooperative cancellation signal checked between emitted rows. */
@FunctionalInterface
public interface CancellationToken {
    CancellationToken NONE = () -> false;
    boolean isCancellationRequested();
}
