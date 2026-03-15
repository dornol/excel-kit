package io.github.dornol.excelkit.shared;

/**
 * Callback for reporting progress during large Excel/CSV writes.
 * <p>
 * Invoked every N rows as configured via {@code onProgress(interval, callback)}.
 *
 * @author dhkim
 */
@FunctionalInterface
public interface ProgressCallback {

    /**
     * Called when the specified number of rows have been processed.
     *
     * @param processedRows total number of rows written so far (across all sheets)
     * @param cursor        the current cursor position
     */
    void onProgress(long processedRows, Cursor cursor);
}
