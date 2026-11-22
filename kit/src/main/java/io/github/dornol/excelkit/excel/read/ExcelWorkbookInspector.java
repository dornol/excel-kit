package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.core.read.ReadLimitExceededException;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.core.io.TempResourceCreator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.jspecify.annotations.Nullable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.UUID;

/**
 * Looks at a workbook without reading its rows: which sheets it has and what a sheet's header row says.
 * <p>
 * Configure what the file needs, then ask; the same inspector can be used for any number of files:
 *
 * <pre>{@code
 * ExcelWorkbookInspector inspector = ExcelWorkbookInspector.create()
 *         .password("secret")             // only for an encrypted workbook
 *         .limits(ReadLimits.UPLOAD);     // for a file from an untrusted source
 * List<ExcelSheetInfo> sheets = inspector.sheetNames(upload);
 * List<String> headers = inspector.headers(upload, 0, 0);
 * }</pre>
 * <p>
 * A stream is copied to a temporary file first (and stays caller-owned); a {@link Path} is opened in place.
 * The {@link ReadLimits} that apply are {@link ReadLimits#maxInputBytes()} (the copy of a stream stops as soon as
 * it passes it, and a file larger than it is refused), {@link ReadLimits#maxSheets()} for {@link #sheetNames} and
 * {@link ReadLimits#maxColumns()} for {@link #headers}. Without limits nothing is bounded.
 * <p>
 * An inspector is a plain holder of settings: it is not thread-safe while it is being configured.
 *
 * @since 0.28.0
 */
public final class ExcelWorkbookInspector {
    private @Nullable String password;
    private ReadLimits limits = ReadLimits.UNLIMITED;

    private ExcelWorkbookInspector() {
    }

    /**
     * Creates an inspector for plain workbooks with no limits.
     *
     * @return a new inspector
     */
    public static ExcelWorkbookInspector create() {
        return new ExcelWorkbookInspector();
    }

    /**
     * Sets the password of an encrypted workbook; a wrong one fails with {@link ExcelReadException}.
     *
     * @param password the password, or {@code null} for a workbook that is not encrypted (the default)
     * @return this inspector for chaining
     */
    public ExcelWorkbookInspector password(@Nullable String password) {
        this.password = password;
        return this;
    }

    /**
     * Sets the limits to enforce, for example {@link ReadLimits#UPLOAD}.
     *
     * @param limits the limits
     * @return this inspector for chaining
     */
    public ExcelWorkbookInspector limits(ReadLimits limits) {
        this.limits = Objects.requireNonNull(limits, "limits cannot be null");
        return this;
    }

    /**
     * Returns the sheets of a workbook, in order.
     *
     * @param inputStream the workbook; it is consumed but not closed
     * @return the sheet names with their 0-based indices
     * @throws ExcelReadException         if the workbook cannot be read, or the password is wrong
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the number of sheets passes a limit
     */
    public List<ExcelSheetInfo> sheetNames(InputStream inputStream) {
        Objects.requireNonNull(inputStream, "inputStream cannot be null");
        return checkSheets(withWorkbook(inputStream, "Failed to read sheet names", ExcelWorkbookInspector::readSheetNames));
    }

    /**
     * Returns the sheets of a workbook file, in order.
     *
     * @param path the workbook
     * @return the sheet names with their 0-based indices
     * @throws ExcelReadException         if the file cannot be opened or read, or the password is wrong
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the file or the number of sheets passes a limit
     */
    public List<ExcelSheetInfo> sheetNames(Path path) {
        Objects.requireNonNull(path, "path cannot be null");
        return checkSheets(withWorkbook(path, "Failed to read sheet names", ExcelWorkbookInspector::readSheetNames));
    }

    /**
     * Like {@link #sheetNames(InputStream)}, for a source whose stream is opened and closed here.
     *
     * @param source the workbook
     * @return the sheet names with their 0-based indices
     */
    public List<ExcelSheetInfo> sheetNames(InputStreamSource source) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return sheetNames(input);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    /**
     * Returns the header names of one sheet.
     *
     * @param inputStream    the workbook; it is consumed but not closed
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @return the header names
     * @throws ExcelReadException         if the workbook or the sheet cannot be read, or the password is wrong
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the input or the header row passes a limit
     */
    public List<String> headers(InputStream inputStream, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(inputStream, "inputStream cannot be null");
        validateIndexes(sheetIndex, headerRowIndex);
        return checkColumns(withWorkbook(inputStream, "Failed to read sheet headers",
                path -> readHeaders(path, sheetIndex, headerRowIndex)));
    }

    /**
     * Returns the header names of one sheet of a workbook file.
     *
     * @param path           the workbook
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @return the header names
     * @throws ExcelReadException         if the file or the sheet cannot be read, or the password is wrong
     * @throws io.github.dornol.excelkit.core.read.ReadLimitExceededException if the file or the header row passes a limit
     */
    public List<String> headers(Path path, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(path, "path cannot be null");
        validateIndexes(sheetIndex, headerRowIndex);
        return checkColumns(withWorkbook(path, "Failed to read sheet headers",
                file -> readHeaders(file, sheetIndex, headerRowIndex)));
    }

    /**
     * Like {@link #headers(InputStream, int, int)}, for a source whose stream is opened and closed here.
     *
     * @param source         the workbook
     * @param sheetIndex     the 0-based sheet index
     * @param headerRowIndex the 0-based header row index
     * @return the header names
     */
    public List<String> headers(InputStreamSource source, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(source, "source cannot be null");
        validateIndexes(sheetIndex, headerRowIndex);
        try (InputStream input = source.openStream()) {
            return headers(input, sheetIndex, headerRowIndex);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    private List<ExcelSheetInfo> checkSheets(List<ExcelSheetInfo> sheets) {
        if (limits.maxSheets() >= 0 && sheets.size() > limits.maxSheets()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.SHEETS,
                    limits.maxSheets(), sheets.size());
        }
        return sheets;
    }

    private List<String> checkColumns(List<String> headers) {
        if (limits.maxColumns() >= 0 && headers.size() > limits.maxColumns()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.COLUMNS,
                    limits.maxColumns(), headers.size());
        }
        return headers;
    }

    // Each sheet stream is opened only to advance the iterator and is closed right away.
    @SuppressWarnings("try")
    private static List<ExcelSheetInfo> readSheetNames(Path path) {
        try (OPCPackage pkg = OPCPackage.open(path.toFile(), PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            List<ExcelSheetInfo> result = new ArrayList<>();
            int index = 0;
            while (sheets.hasNext()) {
                try (InputStream ignored = sheets.next()) {
                    result.add(new ExcelSheetInfo(index, sheets.getSheetName()));
                }
                index++;
            }
            return result;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read sheet names", e);
        }
    }

    private static List<String> readHeaders(Path path, int sheetIndex, int headerRowIndex) {
        try (OPCPackage pkg = OPCPackage.open(path.toFile(), PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sharedStrings = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XMLReader parser = XMLHelper.newXMLReader();
            HeaderExtractor extractor = new HeaderExtractor(headerRowIndex);
            parser.setContentHandler(new ExcelReadSupport.SheetXmlHandler(styles, sharedStrings, extractor,
                    ExcelReadSupport.dataFormatter(reader)));

            Iterator<InputStream> sheets = reader.getSheetsData();
            int currentIndex = 0;
            while (sheets.hasNext()) {
                try (InputStream sheet = sheets.next()) {
                    if (currentIndex == sheetIndex) {
                        parser.parse(new InputSource(sheet));
                        return extractor.headers();
                    }
                }
                currentIndex++;
            }
            throw new ExcelReadException("Sheet index " + sheetIndex + " not found. File has "
                    + currentIndex + " sheet(s).");
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read sheet headers", e);
        }
    }

    /** Copies the stream to a temp file (within the byte limit), decrypts if needed, and runs the operation. */
    private <R> R withWorkbook(InputStream inputStream, String message, WorkbookOperation<R> operation) {
        Path tempDir = null;
        Path tempFile = null;
        Path decrypted = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            ExcelInputCopy.copy(inputStream, tempFile, limits.maxInputBytes());
            if (password == null) return operation.apply(tempFile);
            decrypted = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            ExcelReadSupport.decrypt(tempFile, password, decrypted);
            return operation.apply(decrypted);
        } catch (ExcelReadException | ReadLimitExceededException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException(message, e);
        } finally {
            cleanup(tempDir, decrypted, tempFile);
        }
    }

    /** Opens the file in place; only an encrypted one needs a decrypted temp copy. */
    private <R> R withWorkbook(Path path, String message, WorkbookOperation<R> operation) {
        long size;
        try {
            size = Files.size(path);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
        if (limits.maxInputBytes() >= 0 && size > limits.maxInputBytes()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES,
                    limits.maxInputBytes(), size);
        }
        if (password == null) {
            try {
                return operation.apply(path);
            } catch (ExcelReadException | ReadLimitExceededException e) {
                throw e;
            } catch (Exception e) {
                throw new ExcelReadException(message, e);
            }
        }
        Path tempDir = null;
        Path decrypted = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            decrypted = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            ExcelReadSupport.decrypt(path, password, decrypted);
            return operation.apply(decrypted);
        } catch (ExcelReadException | ReadLimitExceededException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException(message, e);
        } finally {
            cleanup(tempDir, decrypted);
        }
    }

    private static void validateIndexes(int sheetIndex, int headerRowIndex) {
        if (sheetIndex < 0) throw new IllegalArgumentException("sheetIndex must be non-negative");
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex must be non-negative");
    }

    /** Deletes the files, then the directory that held them. */
    private static void cleanup(Path tempDir, Path... files) {
        for (Path file : files) {
            if (file == null) continue;
            try {
                Files.deleteIfExists(file);
            } catch (IOException e) {
                file.toFile().deleteOnExit();
            }
        }
        try {
            if (tempDir != null) Files.deleteIfExists(tempDir);
        } catch (IOException e) {
            tempDir.toFile().deleteOnExit();
        }
    }

    @FunctionalInterface
    private interface WorkbookOperation<R> {
        R apply(Path path) throws Exception;
    }

    private static final class HeaderExtractor extends DefaultHandler
            implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final int headerRowIndex;
        private final List<String> headers = new ArrayList<>();
        private final List<String> currentRow = new ArrayList<>();
        private boolean done;

        private HeaderExtractor(int headerRowIndex) {
            this.headerRowIndex = headerRowIndex;
        }

        @Override
        public void startRow(int rowNum) {
            currentRow.clear();
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum == headerRowIndex) {
                headers.addAll(currentRow);
                done = true;
            }
        }

        // Deprecated in newer POI, but it is the callback POI 5 runtimes invoke.
        @SuppressWarnings("deprecation")
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            cell(cellReference, formattedValue, (Comment) comment);
        }

        // POI 6 callback; keep the XSSFComment overload for POI 5 runtimes.
        // No @Override so this source also compiles against POI 5.
        public void cell(String cellReference, String formattedValue, Comment comment) {
            if (!done) {
                int columnIndex = ExcelReadSupport.getColumnIndex(cellReference);
                while (currentRow.size() < columnIndex) currentRow.add("");
                currentRow.add(formattedValue == null ? "" : formattedValue);
            }
        }

        private List<String> headers() {
            return headers;
        }
    }
}
