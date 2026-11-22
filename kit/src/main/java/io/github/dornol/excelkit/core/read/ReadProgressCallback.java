package io.github.dornol.excelkit.core.read;

/**
 * Receives {@link ReadProgress} snapshots while a reader runs; see {@code onReadProgress(...)}.
 *
 * @since 0.21.0
 */
@FunctionalInterface
public interface ReadProgressCallback {
    /**
     * Called on the reading thread every configured interval and once more when the read ends.
     *
     * @param progress an immutable snapshot of the read so far
     */
    void onProgress(ReadProgress progress);
}
