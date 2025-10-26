package io.github.dornol.excelkit.shared;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

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
    private static final Logger log = LoggerFactory.getLogger(TempResourceContainer.class);
    private Path tempDir;
    private Path tempFile;

    /**
     * Returns the path to the temporary directory (if set).
     */
    protected Path getTempDir() {
        return tempDir;
    }

    /**
     * Sets the path to the temporary directory.
     *
     * @param tempDir The directory to store the temporary file
     */
    protected void setTempDir(Path tempDir) {
        this.tempDir = tempDir;
    }

    /**
     * Returns the path to the temporary file (if set).
     */
    protected Path getTempFile() {
        return tempFile;
    }

    /**
     * Sets the path to the temporary file.
     *
     * @param tempFile The file that will be eventually streamed
     */
    protected void setTempFile(Path tempFile) {
        this.tempFile = tempFile;
    }

    /**
     * Attempts to delete the temporary file and directory (if they exist).
     * <p>
     * Called automatically at the end of file export operations.
     */
    @Override
    public void close() {
        if (tempFile != null) {
            try {
                Files.deleteIfExists(tempFile);
            } catch (IOException e) {
                log.warn("Failed to delete temp file: {}", tempFile, e);
            }
        }
        if (tempDir != null) {
            try {
                Files.deleteIfExists(tempDir);
            } catch (IOException e) {
                log.warn("Failed to delete temp dir: {}", tempDir, e);
            }
        }
    }
}
