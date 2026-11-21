package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/**
 * Abstract container class for managing temporary files and directories.
 * <p>
 * This class is typically used by file-based exporters such as CSV and Excel handlers,
 * providing automatic cleanup of temporary resources via {@link AutoCloseable}.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class TempResourceContainer implements AutoCloseable {
    /** Creates a new empty temp resource container. */
    public TempResourceContainer() {}

    private static final Logger log = LoggerFactory.getLogger(TempResourceContainer.class);
    private @Nullable Path tempDir;
    private @Nullable Path tempFile;
    private boolean ownsFile = true;

    /**
     * Returns the path to the temporary directory (if set).
     *
     * @return the temporary directory path, or null
     */
    protected @Nullable Path getTempDir() {
        return tempDir;
    }

    /**
     * Sets the path to the temporary directory.
     *
     * @param tempDir The directory to store the temporary file
     */
    protected void setTempDir(@Nullable Path tempDir) {
        this.tempDir = tempDir;
    }

    /**
     * Returns the path to the temporary file (if set).
     *
     * @return the temporary file path, or null
     */
    protected @Nullable Path getTempFile() {
        return tempFile;
    }

    /**
     * Sets the path to the temporary file.
     *
     * @param tempFile The file that will be eventually streamed
     */
    protected void setTempFile(@Nullable Path tempFile) {
        this.tempFile = tempFile;
        this.ownsFile = true;
    }

    /** Replaces owned temporary resources with a caller-owned file that must never be deleted. */
    protected void useExternalFile(Path file) {
        close();
        this.tempDir = null;
        this.tempFile = Objects.requireNonNull(file, "file cannot be null");
        this.ownsFile = false;
    }

    /**
     * Attempts to delete the temporary file and directory (if they exist).
     * <p>
     * Called automatically at the end of file export operations.
     */
    @Override
    public void close() {
        releaseQuietly(ownsFile ? tempFile : null, tempDir);
    }

    /**
     * Deletes a temp file and its directory, falling back to delete-on-exit when deletion fails.
     * Static so that a cleanup action can run without holding a reference to the container.
     *
     * @param file the file to delete, or null
     * @param dir  the directory to delete, or null
     */
    protected static void releaseQuietly(@Nullable Path file, @Nullable Path dir) {
        if (file != null) {
            try {
                Files.deleteIfExists(file);
            } catch (IOException e) {
                log.warn("Failed to delete temp file: {}", file, e);
                file.toFile().deleteOnExit();
            }
        }
        if (dir != null) {
            try {
                Files.deleteIfExists(dir);
            } catch (IOException e) {
                log.warn("Failed to delete temp dir: {}", dir, e);
                dir.toFile().deleteOnExit();
            }
        }
    }
}
