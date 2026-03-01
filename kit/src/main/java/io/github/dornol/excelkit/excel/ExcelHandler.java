package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.TempResourceCreator;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.NonNull;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;

/**
 * Handles the final output stage of an Excel export.
 * <p>
 * This class is responsible for writing the {@link SXSSFWorkbook} to an {@link OutputStream},
 * with optional support for Excel password encryption.
 * It ensures that the workbook is consumed only once.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelHandler {
    private final SXSSFWorkbook wb;
    private boolean consumed = false;

    /**
     * Constructs an ExcelHandler wrapping the given workbook.
     *
     * @param wb The SXSSFWorkbook to be written
     */
    ExcelHandler(SXSSFWorkbook wb) {
        this.wb = wb;
    }

    /**
     * Writes the workbook to the given OutputStream.
     * <p>
     * This method can only be called once; subsequent calls will throw an exception.
     *
     * @param outputStream The OutputStream to write the Excel file to
     * @throws IOException If an I/O error occurs during writing
     * @throws IllegalStateException If this method has already been called
     */
    public void consumeOutputStream(@NonNull OutputStream outputStream) throws IOException {
        if (consumed) {
            throw new ExcelWriteException("Already consumed");
        }
        try {
            wb.write(outputStream);
        } finally {
            consumed = true;
            wb.close();
        }
    }

    /**
     * Writes the workbook to the given OutputStream with Excel-compatible password encryption.
     * <p>
     * This method encrypts the file using the "agile" encryption mode supported by modern Excel versions.
     *
     * @param outputStream The OutputStream to write the encrypted Excel file to
     * @param password     The password to protect the Excel file with
     * @throws IOException If an I/O or encryption error occurs during writing
     * @throws IllegalStateException If this method has already been called
     */
    public void consumeOutputStreamWithPassword(@NonNull OutputStream outputStream, @NonNull String password) throws IOException {
        if (consumed) {
            throw new ExcelWriteException("Already consumed");
        }
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }

        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, "excel-enc", ".tmp");
        try {
            // Write workbook to temp file first to free SXSSFWorkbook memory
            try (OutputStream tempOut = Files.newOutputStream(tempFile)) {
                wb.write(tempOut);
            } finally {
                consumed = true;
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
            try { Files.deleteIfExists(tempFile); } catch (IOException ignored) { }
            try { Files.deleteIfExists(tempDir); } catch (IOException ignored) { }
        }
    }

}
