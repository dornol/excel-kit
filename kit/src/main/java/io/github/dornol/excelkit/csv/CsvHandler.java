package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.core.io.TempResourceCleaner;
import io.github.dornol.excelkit.core.io.TempResourceContainer;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.ref.Cleaner;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Handles the output stage of a CSV export.
 * <p>
 * This class holds a temporary CSV file and writes its content to a provided {@link OutputStream}.
 * It ensures the file is only consumed once, and automatically cleans up temporary files afterward.
 * <p>
 * A handler that will not be written must be released with {@link #discard()}. As a safety net, a handler that is
 * garbage collected without being consumed deletes its temp files then; that happens at a time of the garbage
 * collector's choosing, so do not rely on it.
 * <p>
 * It is not {@link AutoCloseable}, like {@link io.github.dornol.excelkit.excel.ExcelHandler}: the handler is usually
 * written after the method that made it has returned (a {@code StreamingResponseBody}), and {@code try (handler)}
 * would release the file before that.
 *
 * @author dhkim
 * @since 0.1.2
 */
public final class CsvHandler implements FileHandler {
    /** Deletes the staged file and its directory; the one place that owns that knowledge. */
    private static final class Staging extends TempResourceContainer {
        Staging(Path dir, Path file) {
            setTempDir(dir);
            setTempFile(file);
        }

        Path file() {
            return getTempFile();
        }

        static void delete(Path file, Path dir) {
            releaseQuietly(file, dir);
        }
    }

    /** What must be released. Holds paths only, never the handler, so the handler can be collected. */
    private static final class Release implements Runnable {
        private final AtomicBoolean consumed = new AtomicBoolean(false);
        private final Path tempDir;
        private final Path tempFile;

        Release(Path tempDir, Path tempFile) {
            this.tempDir = tempDir;
            this.tempFile = tempFile;
        }

        boolean claim() {
            return consumed.compareAndSet(false, true);
        }

        @Override
        public void run() {
            if (claim()) {
                Staging.delete(tempFile, tempDir);
            }
        }
    }

    private final Staging staging;
    private final Release release;
    private final Cleaner.Cleanable cleanable;

    /**
     * Creates a new CsvHandler wrapping the given temp file and directory.
     *
     * @param tempDir  The temporary directory containing the CSV file
     * @param tempFile The path to the CSV file to be output
     */
    CsvHandler(Path tempDir, Path tempFile) {
        this.staging = new Staging(tempDir, tempFile);
        this.release = new Release(tempDir, tempFile);
        this.cleanable = TempResourceCleaner.register(this, release);
    }

    /**
     * Releases the temp file without writing it. Does nothing if the handler was already consumed or
     * discarded. Afterwards {@link #writeTo(OutputStream)} throws {@link CsvWriteException}.
     *
     * @since 0.26.0
     */
    @Override
    public void discard() {
        cleanable.clean();
    }

    /**
     * Writes the content of the CSV file to the given OutputStream.
     * <p>
     * This method can be called only once. Subsequent calls will throw {@link CsvWriteException}.
     * <p>
     * The temporary file and directory will be deleted automatically after writing.
     * <p>
     * Note: I/O errors are wrapped as {@link CsvWriteException} (unchecked) rather than
     * propagated as {@link IOException}, so callers do not need a try-catch for checked exceptions.
     *
     * @param outputStream The stream to which the CSV content will be written
     * @throws CsvWriteException If this method has already been called or if an I/O error occurs
     */
    @Override
    public void writeTo(OutputStream outputStream) {
        if (!release.claim()) {
            throw new CsvWriteException("Already consumed");
        }
        try {
            try (InputStream is = Files.newInputStream(staging.file())) {
                is.transferTo(outputStream);
            }
        } catch (IOException e) {
            throw new CsvWriteException("Failed to write CSV", e);
        } finally {
            staging.close();
            cleanable.clean();
        }
    }
}
