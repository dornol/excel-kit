package io.github.dornol.excelkit.core;

import java.lang.ref.Cleaner;

/**
 * Safety net for temp resources whose owner is dropped without being consumed or closed.
 * <p>
 * A handler owns its temp files from the moment it is created until it is consumed. If the handler
 * is garbage collected before that, the registered action releases them. The timing depends on the
 * garbage collector, so callers should still consume or discard handlers explicitly.
 * <p>
 * Internal helper shared by the Excel and CSV handlers; it is not part of the supported API.
 */
public final class TempResourceCleaner {
    private static final Cleaner CLEANER = Cleaner.create();

    private TempResourceCleaner() {
    }

    /**
     * Registers an action that runs at most once: when {@code owner} becomes unreachable, or when the
     * returned handle is cleaned, whichever comes first.
     * <p>
     * The action must not reference {@code owner}, otherwise the owner is never collected.
     *
     * @param owner  the object whose reachability triggers the action
     * @param action the release action
     * @return a handle that runs the action immediately and unregisters it
     */
    public static Cleaner.Cleanable register(Object owner, Runnable action) {
        return CLEANER.register(owner, action);
    }
}
