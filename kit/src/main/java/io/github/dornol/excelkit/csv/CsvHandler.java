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
 * A handler that will not be written must be released with {@link #discard()} (or {@link #close()},
 * which does the same). As a safety net, a handler that is garbage collected without being consumed
 * deletes its temp files then; that happens at a time of the garbage collector's choosing, so do not
 * rely on it.
 *
 * @author dhkim
 * @since 0.1.2
 */
public final class CsvHandler extends TempResourceContainer implements FileHandler {
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
                releaseQuietly(tempFile, tempDir);
            }
        }
    }

    private final Release release;
    private final Cleaner.Cleanable cleanable;

    /**
     * Creates a new CsvHandler wrapping the given temp file and directory.
     *
     * @param tempDir  The temporary directory containing the CSV file
     * @param tempFile The path to the CSV file to be output
     */
    CsvHandler(Path tempDir, Path tempFile) {
        setTempFile(tempFile);
        setTempDir(tempDir);
        this.release = new Release(tempDir, tempFile);
        this.cleanable = TempResourceCleaner.register(this, release);
    }

    /**
     * Releases the temp file without writing it. Does nothing if the handler was already consumed or
     * discarded. Afterwards {@link #writeTo(OutputStream)} throws {@link CsvWriteException}.
     *
     * @since 0.26.0
     */
    public void discard() {
        cleanable.clean();
    }

    /** Same as {@link #discard()}. */
    @Override
    public void close() {
        discard();
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
            try (InputStream is = Files.newInputStream(getTempFile())) {
                is.transferTo(outputStream);
            }
        } catch (IOException e) {
            throw new CsvWriteException("Failed to write CSV", e);
        } finally {
            super.close();
            cleanable.clean();
        }
    }
}
