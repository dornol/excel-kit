package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.core.TempResourceCreator;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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
 * The workbook is always closed inside {@link #writeTo(OutputStream)} overloads,
 * and there is no realistic code path where the handler is obtained but never consumed,
 * since callers either invoke it immediately or pass it to a {@code StreamingResponseBody} lambda.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public final class ExcelHandler implements FileHandler {
    private static final Logger log = LoggerFactory.getLogger(ExcelHandler.class);
    private final SXSSFWorkbook wb;
    private final @Nullable String password;
    private final AtomicBoolean consumed = new AtomicBoolean(false);

    /**
     * Constructs an ExcelHandler wrapping the given workbook.
     *
     * @param wb The SXSSFWorkbook to be written
     */
    ExcelHandler(SXSSFWorkbook wb) {
        this(wb, null);
    }

    /**
     * Constructs an ExcelHandler wrapping the given workbook with an optional encryption password.
     * <p>
     * When a non-null password is provided, {@link #writeTo(OutputStream)} will
     * automatically encrypt the output using the "agile" encryption mode.
     *
     * @param wb       The SXSSFWorkbook to be written
     * @param password The password for file encryption, or null for no encryption
     */
    ExcelHandler(SXSSFWorkbook wb, @Nullable String password) {
        this.wb = wb;
        this.password = password;
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
                encryptAndWrite(outputStream, password);
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
     * @param outputStream The OutputStream to write the encrypted Excel file to
     * @param password     The password to protect the Excel file with
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.5
     */
    public void writeTo(OutputStream outputStream, String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
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
     * @param outputStream The OutputStream to write the encrypted Excel file to
     * @param password     The password as a char array (will be zeroed after use)
     * @throws ExcelWriteException If an I/O or encryption error occurs during writing
     * @throws IllegalStateException If a password was already set at the writer level, or if already consumed
     * @since 0.16.5
     */
    public void writeTo(OutputStream outputStream, char[] password) {
        if (password == null || password.length == 0 || isBlank(password)) {
            throw new IllegalArgumentException("Password cannot be null or blank");
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
            wb.close();
        }
    }

    private void markConsumed() {
        if (!consumed.compareAndSet(false, true)) {
            throw new ExcelWriteException("Already consumed");
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
                wb.close();
            }

            // Encrypt from temp file using file-based POIFSFileSystem (low memory)
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
