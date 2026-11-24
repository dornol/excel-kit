package io.github.dornol.excelkit.excel.read;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Tells what a file is before POI is asked to open it as an {@code .xlsx}, so that a file that is not one fails with a
 * message that says what it is instead of with a ZIP or an OOXML exception: a workbook of Excel 97-2003 ({@code .xls}),
 * a password-protected {@code .xlsx} read without its password, or something else altogether.
 */
final class ExcelFileFormat {

    /** What the content of a file is, judged by its first bytes and, for an OLE2 container, by its streams. */
    enum Kind {
        /** A ZIP archive: an {@code .xlsx} as far as the first bytes can tell (the security scan looks further). */
        ZIP,
        /** An Excel 97-2003 workbook ({@code .xls}). */
        XLS,
        /** A password-protected {@code .xlsx}: an OLE2 container around the encrypted package. */
        ENCRYPTED_XLSX,
        /** An OLE2 container that is neither of the two (a Word document, for example). */
        OTHER_OLE2,
        /** Not empty, not a ZIP and not OLE2: text, an image, a CSV... */
        NOT_OFFICE,
        /** No bytes at all. */
        EMPTY
    }

    private static final byte[] OLE2 = {(byte) 0xD0, (byte) 0xCF, 0x11, (byte) 0xE0, (byte) 0xA1, (byte) 0xB1, 0x1A, (byte) 0xE1};

    private ExcelFileFormat() {
    }

    static Kind detect(Path file) throws IOException {
        byte[] head = new byte[OLE2.length];
        int read;
        try (InputStream in = Files.newInputStream(file)) {
            read = in.readNBytes(head, 0, head.length);
        }
        if (read == 0) return Kind.EMPTY;
        if (read >= 2 && head[0] == 'P' && head[1] == 'K') return Kind.ZIP;
        if (read == OLE2.length && java.util.Arrays.equals(head, OLE2)) return oleKind(file);
        return Kind.NOT_OFFICE;
    }

    private static Kind oleKind(Path file) {
        try (POIFSFileSystem fs = new POIFSFileSystem(file.toFile(), true)) {
            var root = fs.getRoot();
            if (root.hasEntry("EncryptedPackage")) return Kind.ENCRYPTED_XLSX;
            if (root.hasEntry("Workbook") || root.hasEntry("Book")) return Kind.XLS;
            return Kind.OTHER_OLE2;
        } catch (IOException | RuntimeException e) {
            return Kind.OTHER_OLE2;
        }
    }

    /**
     * Fails with a message that says what the file is when it cannot be read as an {@code .xlsx}.
     *
     * @param encryptedAllowed whether a password-protected workbook is acceptable here: the caller decrypts it
     */
    static void requireXlsx(Path file, boolean encryptedAllowed) {
        Kind kind;
        try {
            kind = detect(file);
        } catch (IOException e) {
            // Not ours to explain: the read that follows reports the failure with its own message.
            return;
        }
        switch (kind) {
            case ZIP -> { }
            case ENCRYPTED_XLSX -> {
                if (!encryptedAllowed) {
                    throw new ExcelReadException("The file is a password-protected Excel workbook: give its password "
                            + "with password(...) to read it");
                }
            }
            case XLS -> throw new ExcelReadException("The file is an Excel 97-2003 workbook (.xls), which excel-kit does not "
                    + "read: save it as .xlsx (in Excel: File > Save As > Excel Workbook) and read that file");
            case OTHER_OLE2 -> throw new ExcelReadException("The file is not an Excel workbook: it is an OLE2 document of "
                    + "another kind (a Word or PowerPoint file, for example)");
            case NOT_OFFICE -> throw new ExcelReadException("The file is not an Excel workbook: an .xlsx file is a ZIP "
                    + "archive, and this one is not (is it a CSV or another text file?)");
            case EMPTY -> throw new ExcelReadException("The file is empty: there is no Excel workbook in it");
        }
    }
}
