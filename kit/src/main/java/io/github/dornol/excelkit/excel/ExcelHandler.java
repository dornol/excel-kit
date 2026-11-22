package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.io.AtomicFileWriter;
import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.core.io.TempResourceCleaner;
import io.github.dornol.excelkit.core.io.TempResourceCreator;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.ref.Cleaner;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Handles the final output stage of an Excel export.
 * <p>
 * This class is responsible for writing the {@link SXSSFWorkbook} to an {@link OutputStream},
 * with optional support for Excel password encryption.
 * It ensures that the workbook is consumed only once.
 *
 * <p><b>Why not AutoCloseable?</b>
 * The primary usage pattern is {@code ResponseEntity<StreamingResponseBody>},
 * where the handler is captured by a lambda and consumed asynchronously.
 * Implementing {@link AutoCloseable} would cause IDE "resource not closed" warnings
 * in this pattern, requiring {@code @SuppressWarnings("resource")} on every call site.
 * <p>
 * The workbook is closed inside the {@link #writeTo(OutputStream)} overloads. A handler that will not
 * be written must be released with {@link #discard()}. As a safety net, a handler that is garbage
 * collected without being consumed closes its workbook then; that happens at a time of the garbage
 * collector's choosing, so do not rely on it.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public final class ExcelHandler implements FileHandler {
    private static final Logger log = LoggerFactory.getLogger(ExcelHandler.class);
    private final SXSSFWorkbook wb;
    private final char @Nullable [] password;

    /**
     * What must be released. It references the workbook but never the handler, so the handler can be
     * garbage collected and then triggers this through {@link TempResourceCleaner}.
     */
    private static final class WorkbookRelease implements Runnable {
        private final SXSSFWorkbook workbook;
        private final AtomicBoolean consumed = new AtomicBoolean(false);

        WorkbookRelease(SXSSFWorkbook workbook) {
            this.workbook = workbook;
        }

        /** Marks the workbook as consumed; false if it already was. */
        boolean claim() {
            return consumed.compareAndSet(false, true);
        }

        @Override
        public void run() {
            if (claim()) {
                try {
                    ExcelWorkbookSupport.close(workbook);
                } catch (IOException | RuntimeException e) {
                    log.warn("Failed to close an unconsumed workbook", e);
                }
            }
        }
    }

    private final WorkbookRelease release;
    private final Cleaner.Cleanable cleanable;

    /**
     * Constructs an ExcelHandler wrapping the given workbook.
     *
     * @param wb The SXSSFWorkbook to be written
     */
    ExcelHandler(SXSSFWorkbook wb) {
        this.wb = wb;
        this.password = null;
        this.release = new WorkbookRelease(wb);
        this.cleanable = TempResourceCleaner.register(this, release);
    }

    /**
     * Constructs an ExcelHandler wrapping the given workbook with an optional encryption password.
     *
     * @param wb       The SXSSFWorkbook to be written
     * @param password The password for file encryption, or null for no encryption
     */
    ExcelHandler(SXSSFWorkbook wb, @Nullable String password) {
        this.wb = wb;
        this.password = password != null ? password.toCharArray() : null;
        this.release = new WorkbookRelease(wb);
        this.cleanable = TempResourceCleaner.register(this, release);
    }

    /**
     * Constructs an ExcelHandler wrapping the given workbook with a char[] encryption password.
     * <p>
     * The array is copied internally; the caller may zero the original after this call.
     *
     * @param wb       The SXSSFWorkbook to be written
     * @param password The password as a char array (copied internally), or null for no encryption
     */
    ExcelHandler(SXSSFWorkbook wb, char @Nullable [] password) {
        this.wb = wb;
        this.password = password != null ? password.clone() : null;
        this.release = new WorkbookRelease(wb);
        this.cleanable = TempResourceCleaner.register(this, release);
    }

    /**
     * Releases the workbook and its temp files without writing it.
     * <p>
     * Use this when the handler will not be written, for example when the export is cancelled or
     * the handler is dropped after an error. It does nothing if the handler was already consumed or
     * discarded; afterwards the {@code writeTo} methods throw {@link ExcelWriteException}.
     *
     * @since 0.26.0
     */
    public void discard() {
        cleanable.clean();
    }

    /**
     * Writes the workbook to the given OutputStream.
     * <p>
     * If a password was set via {@link ExcelWriter#password(String)} or {@link ExcelWorkbook#password(String)},
     * the output is automatically encrypted using the "agile" encryption mode.
     * <p>
     * This method can only be called once; subsequent calls will throw an exception.
     *
     * @param outputStream The OutputStream to write the Excel file to
     * @throws ExcelWriteException If this method has already been called or if an I/O error occurs
     */
    @Override
    public void writeTo(OutputStream outputStream) {
        try {
            if (password != null) {
                try {
                    encryptAndWrite(outputStream, new String(password));
                } finally {
                    Arrays.fill(password, '\0');
                }
            } else {
                writePlain(outputStream);
            }
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write Excel", e);
        }
    }

    /**
     * Writes the workbook to the given OutputStream with Excel-compatible password encryption.
     * <p>
     * This overload encrypts the file using the "agile" encryption mode supported by modern Excel versions.
     * Cannot be used when a password was already set via {@link ExcelWriter#password(String)}
     * or {@link ExcelWorkbook#password(String)} — use {@link #writeTo(OutputStream)} instead.
     *
     * <p><b>Memory note:</b> POI's encryption API does not support file-backed buffering
     * for the OLE container — the encrypted payload is held in heap (≈ size of the
     * resulting XLSX) before being streamed out. For very large exports (hundreds of MB)
     * plan heap accordingly, or prefer the plain {@link #writeTo(OutputStream)} path.
     *
     * @param outputStream The OutputStream to write the encrypted Excel file to
     * @param password     The password to protect the Excel file with
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalArgumentException If the password is null or blank; the handler is then consumed and its workbook closed
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.5
     */
    public void writeTo(OutputStream outputStream, String password) {
        if (password == null || password.isBlank()) {
            throw rejectBlankPassword();
        }
        try {
            rejectIfPasswordAlreadySet();
            encryptAndWrite(outputStream, password);
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write encrypted Excel", e);
        }
    }

    /**
     * Writes the workbook to the given OutputStream with Excel-compatible password encryption.
     * <p>
     * This overload accepts a {@code char[]} to allow callers to clear the password from memory after use.
     * The array is zeroed out after encryption completes (or on failure).
     * Cannot be used when a password was already set via {@link ExcelWriter#password(String)}
     * or {@link ExcelWorkbook#password(String)} — use {@link #writeTo(OutputStream)} instead.
     *
     * <p><b>Memory note:</b> same as {@link #writeTo(OutputStream, String)} — POI buffers
     * the encrypted OLE container in heap.
     *
     * @param outputStream The OutputStream to write the encrypted Excel file to
     * @param password     The password as a char array (will be zeroed after use)
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalArgumentException If the password is null or blank; the handler is then consumed and its workbook closed
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.5
     */
    public void writeTo(OutputStream outputStream, char[] password) {
        if (password == null || password.length == 0 || isBlank(password)) {
            throw rejectBlankPassword();
        }
        try {
            rejectIfPasswordAlreadySet();
            encryptAndWrite(outputStream, new String(password));
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write encrypted Excel", e);
        } finally {
            Arrays.fill(password, '\0');
        }
    }

    /**
     * Writes the workbook directly to a file path with Excel-compatible password encryption.
     * <p>
     * Convenience overload that opens a buffered {@link OutputStream} to the given path
     * and delegates to {@link #writeTo(OutputStream, String)}.
     * Cannot be used when a password was already set via {@link ExcelWriter#password(String)}
     * or {@link ExcelWorkbook#password(String)} — use {@link #writeTo(Path)} instead.
     *
     * @param path     The destination file path
     * @param password The password to protect the Excel file with
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalArgumentException If the password is null or blank; the handler is then consumed and its workbook closed
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.6
     */
    public void writeTo(Path path, String password) {
        if (password == null || password.isBlank()) {
            throw rejectBlankPassword();
        }
        rejectIfPasswordAlreadySet();
        try {
            AtomicFileWriter.write(path, out -> writeBuffered(out, password));
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write encrypted Excel to: " + path, e);
        }
    }

    /**
     * Writes the workbook directly to a file path with Excel-compatible password encryption.
     * <p>
     * This overload accepts a {@code char[]} to allow callers to clear the password from memory after use.
     * The array is zeroed out after encryption completes (or on failure).
     * Cannot be used when a password was already set via {@link ExcelWriter#password(String)}
     * or {@link ExcelWorkbook#password(String)} — use {@link #writeTo(Path)} instead.
     *
     * @param path     The destination file path
     * @param password The password as a char array (will be zeroed after use)
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalArgumentException If the password is null or blank; the handler is then consumed and its workbook closed
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.6
     */
    public void writeTo(Path path, char[] password) {
        if (password == null || password.length == 0 || isBlank(password)) {
            throw rejectBlankPassword();
        }
        try {
            rejectIfPasswordAlreadySet();
            try {
                AtomicFileWriter.write(path, out -> writeBuffered(out, password));
            } catch (IOException e) {
                throw new ExcelWriteException("Failed to write encrypted Excel to: " + path, e);
            }
        } finally {
            // writeTo(OutputStream, char[]) zeroes the array on its own path;
            // zero again here to cover the case where we threw before reaching it
            // (e.g., rejectIfPasswordAlreadySet or Files.newOutputStream failure).
            // Second zero on already-zeroed array is harmless.
            Arrays.fill(password, '\0');
        }
    }

    private void writeBuffered(OutputStream out, String password) {
        // Closing the buffer flushes it; closing the underlying stream again is harmless.
        try (OutputStream buffered = new BufferedOutputStream(out)) {
            writeTo(buffered, password);
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write encrypted Excel", e);
        }
    }

    private void writeBuffered(OutputStream out, char[] password) {
        try (OutputStream buffered = new BufferedOutputStream(out)) {
            writeTo(buffered, password);
        } catch (IOException e) {
            throw new ExcelWriteException("Failed to write encrypted Excel", e);
        }
    }

    /**
     * A null or blank password is an argument error, not a recoverable state: the handler is
     * consumed and its workbook closed, so a handler that is abandoned after the failure (for
     * example inside a streaming response) does not keep temp files.
     */
    private IllegalArgumentException rejectBlankPassword() {
        discard();
        return new IllegalArgumentException("Password cannot be null or blank");
    }

    private void rejectIfPasswordAlreadySet() {
        if (this.password != null) {
            throw new IllegalStateException(
                    "Password is already set via ExcelWriter.password() or ExcelWorkbook.password(). "
                            + "Use writeTo(OutputStream) instead, or remove the password() call to pass the password to writeTo(OutputStream, ...).");
        }
    }

    private static boolean isBlank(char[] chars) {
        for (char c : chars) {
            if (!Character.isWhitespace(c)) {
                return false;
            }
        }
        return true;
    }

    private void writePlain(OutputStream outputStream) throws IOException {
        markConsumed();
        try {
            wb.write(outputStream);
        } finally {
            closeConsumedWorkbook();
        }
    }

    private void markConsumed() {
        if (!release.claim()) {
            throw new ExcelWriteException("Already consumed");
        }
    }

    /** Closes the workbook this handler consumed, then unregisters the safety net. */
    private void closeConsumedWorkbook() throws IOException {
        try {
            ExcelWorkbookSupport.close(wb);
        } finally {
            cleanable.clean();
        }
    }

    private void encryptAndWrite(OutputStream outputStream, String password) throws IOException {
        markConsumed();

        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "excel-enc", ".tmp");
        try {
            // Write workbook to temp file first to free SXSSFWorkbook memory
            try (OutputStream tempOut = Files.newOutputStream(tempFile)) {
                wb.write(tempOut);
            } finally {
                closeConsumedWorkbook();
            }

            // Encrypt the plain temp file into an in-memory POIFS, then stream it out.
            // POI's encryption API buffers the OLE container in heap — there is no
            // file-backed variant. The two-step design (temp file first, then close
            // the SXSSF workbook, THEN encrypt) prevents the SXSSF window memory
            // and the encryption buffer from coexisting as peak memory.
            try (POIFSFileSystem fs = new POIFSFileSystem()) {
                EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
                Encryptor enc = info.getEncryptor();
                enc.confirmPassword(password);

                try (OutputStream encOut = enc.getDataStream(fs)) {
                    try (InputStream tempIn = Files.newInputStream(tempFile)) {
                        tempIn.transferTo(encOut);
                    }
                } catch (GeneralSecurityException e) {
                    throw new ExcelWriteException("Failed to encrypt Excel file", e);
                }

                fs.writeFilesystem(outputStream);
            }
        } finally {
            // Clean up temp resources
            try {
                Files.deleteIfExists(tempFile);
            } catch (IOException e) {
                log.warn("Failed to delete temp file: {}", tempFile, e);
                tempFile.toFile().deleteOnExit();
            }
            try {
                Files.deleteIfExists(tempDir);
            } catch (IOException e) {
                log.warn("Failed to delete temp dir: {}", tempDir, e);
                tempDir.toFile().deleteOnExit();
            }
        }
    }

}
