package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.TempResourceContainer;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Handles the output stage of a CSV export.
 * <p>
 * This class holds a temporary CSV file and writes its content to a provided {@link OutputStream}.
 * It ensures the file is only consumed once, and automatically cleans up temporary files afterward.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvHandler extends TempResourceContainer {
    private boolean consumed = false;

    /**
     * Creates a new CsvHandler wrapping the given temp file and directory.
     *
     * @param tempDir  The temporary directory containing the CSV file
     * @param tempFile The path to the CSV file to be output
     */
    CsvHandler(Path tempDir, Path tempFile) {
        setTempFile(tempFile);
        setTempDir(tempDir);
    }

    /**
     * Writes the content of the CSV file to the given OutputStream.
     * <p>
     * This method can be called only once. Subsequent calls will throw {@link IllegalStateException}.
     * <p>
     * The temporary file and directory will be deleted automatically after writing.
     *
     * @param outputStream The stream to which the CSV content will be written
     * @throws IllegalStateException If this method has already been called
     */
    public void consumeOutputStream(OutputStream outputStream) {
        if (consumed) {
            throw new IllegalStateException("Already consumed");
        }
        try {
            try (InputStream is = Files.newInputStream(getTempFile())) {
                is.transferTo(outputStream);
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        } finally {
            consumed = true;
            super.close();
        }
    }
}
