package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.AbstractReadHandler;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.DuplicateHeaderPolicy;
import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.ReadColumn;
import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.CellConversionConfig;
import io.github.dornol.excelkit.core.ReadAbortException;
import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowData;
import io.github.dornol.excelkit.core.TempResourceCreator;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.jspecify.annotations.Nullable;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.UUID;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.zip.ZipFile;

/**
 * Reads Excel (.xlsx) files using Apache POI's event-based streaming API.
 * <p>
 * This handler parses sheet data row by row, maps values to Java objects, and performs optional validation.
 * It is optimized for large files and avoids loading the entire workbook into memory.
 *
 * <h2>Resource management</h2>
 * On construction, the handler copies the input stream to a temporary file on disk so that
 * the underlying POI API can read it. These temp resources (and, if applicable, the decrypted
 * copy of an encrypted file) are released when {@link #read(Consumer)},
 * {@link #readWhile(java.util.function.Predicate)}, or {@link #readStrict(Consumer)} returns or throws.
 *
 * <h2>Large file tuning</h2>
 * For large or complex Excel files, you may need to adjust POI's internal limits via
 * {@link ExcelKitConfig#configureLargeFileSupport()} before reading. This adjusts:
 * <ul>
 *     <li>{@code ZipSecureFile.setMaxFileCount} — maximum number of internal zip entries</li>
 *     <li>{@code IOUtils.setByteArrayMaxOverride} — maximum in-memory byte array size</li>
 * </ul>
 *
 * @param <T> The target row data type to map each row into
 * @author dhkim
 * @since 2025-07-19
 */
final class ExcelReadHandler<T> extends AbstractReadHandler<T> {
    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(ExcelReadHandler.class);

    static <T> ExcelReadHandler<T> forPath(Path path, List<ReadColumn<T>> columns, Supplier<T> supplier,
            @Nullable Validator validator, int sheetIndex, int headerRowIndex, int headerRows,
            int progressInterval, @Nullable ProgressCallback progressCallback, @Nullable String password,
            boolean countRows, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
            @Nullable CellConversionConfig conversion, long maxRows, boolean skipBlankRows,
            int stopAtBlankRows) {
        var handler = new ExcelReadHandler<>(InputStream.nullInputStream(), columns, supplier, validator,
                sheetIndex, headerRowIndex, headerRows, progressInterval, progressCallback, password,
                countRows, strictHeaders, duplicateHeaderPolicy, conversion, maxRows, skipBlankRows,
                stopAtBlankRows);
        handler.useExternalInput(path);
        return handler;
    }

    static <T> ExcelReadHandler<T> forPath(Path path, Function<RowData, T> mapper,
            @Nullable Validator validator, int sheetIndex, int headerRowIndex, int headerRows,
            int progressInterval, @Nullable ProgressCallback progressCallback, @Nullable String password,
            boolean countRows, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
            @Nullable Set<String> selectedColumns, @Nullable CellConversionConfig conversion,
            long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        var handler = new ExcelReadHandler<>(InputStream.nullInputStream(), mapper, validator, sheetIndex,
                headerRowIndex, headerRows, progressInterval, progressCallback, password, countRows,
                strictHeaders, duplicateHeaderPolicy, selectedColumns, conversion, maxRows, skipBlankRows,
                stopAtBlankRows);
        handler.useExternalInput(path);
        return handler;
    }

    private final @Nullable List<ReadColumn<T>> columns;
    private final int sheetIndex;
    private final int headerRowIndex;
    private final int headerRows;
    private final int progressInterval;
    private final @Nullable ProgressCallback progressCallback;
    private final @Nullable String password;
    private final boolean countRows;

    /**
     * Constructs a handler for reading the first sheet of an Excel file.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     */
    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, 0, 1, 0, null, null, false);
    }

    /**
     * Constructs a handler for reading a specific sheet of an Excel file.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     * @param sheetIndex       The zero-based index of the sheet to read
     */
    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int sheetIndex) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, 0, 1, 0, null, null, false);
    }

    /**
     * Constructs a handler for reading a specific sheet with a custom header row index.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     * @param sheetIndex       The zero-based index of the sheet to read
     * @param headerRowIndex   The zero-based index of the header row (rows before this are skipped)
     */
    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int sheetIndex, int headerRowIndex) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, 1, 0, null, null, false);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int progressInterval, @Nullable ProgressCallback progressCallback) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, 1, progressInterval, progressCallback, (String) null, false);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, headerRows,
                progressInterval, progressCallback, password, countRows, false, DuplicateHeaderPolicy.FIRST);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, headerRows,
                progressInterval, progressCallback, password, countRows, strictHeaders, duplicateHeaderPolicy, null);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                     @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, headerRows,
                progressInterval, progressCallback, password, countRows, strictHeaders, duplicateHeaderPolicy,
                cellConversionConfig, -1, false, 0);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                     @Nullable CellConversionConfig cellConversionConfig,
                     long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        super(inputStream, instanceSupplier, validator, ".xlsx", strictHeaders, duplicateHeaderPolicy,
                null, cellConversionConfig, maxRows, skipBlankRows, stopAtBlankRows);
        validateColumns(columns);
        validateSheetIndex(sheetIndex);
        validateHeaderRowIndex(headerRowIndex);
        validateHeaderRows(headerRows);
        this.columns = columns;
        this.sheetIndex = sheetIndex;
        this.headerRowIndex = headerRowIndex;
        this.headerRows = headerRows;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
        this.password = password;
        this.countRows = countRows;
    }

    /**
     * Constructs a handler in mapping mode for immutable object construction.
     */
    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int progressInterval, @Nullable ProgressCallback progressCallback) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, 1, progressInterval, progressCallback, (String) null, false);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, headerRows, progressInterval,
                progressCallback, password, countRows, false, DuplicateHeaderPolicy.FIRST);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, headerRows, progressInterval,
                progressCallback, password, countRows, strictHeaders, duplicateHeaderPolicy, null);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                     @Nullable Set<String> selectedMapColumns) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, headerRows, progressInterval,
                progressCallback, password, countRows, strictHeaders, duplicateHeaderPolicy, selectedMapColumns, null);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                     @Nullable Set<String> selectedMapColumns,
                     @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, headerRows, progressInterval,
                progressCallback, password, countRows, strictHeaders, duplicateHeaderPolicy, selectedMapColumns,
                cellConversionConfig, -1, false, 0);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password, boolean countRows,
                     boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                     @Nullable Set<String> selectedMapColumns,
                     @Nullable CellConversionConfig cellConversionConfig,
                     long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        super(inputStream, rowMapper, validator, ".xlsx", strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, cellConversionConfig, maxRows, skipBlankRows, stopAtBlankRows);
        validateSheetIndex(sheetIndex);
        validateHeaderRowIndex(headerRowIndex);
        validateHeaderRows(headerRows);
        this.columns = null;
        this.sheetIndex = sheetIndex;
        this.headerRowIndex = headerRowIndex;
        this.headerRows = headerRows;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
        this.password = password;
        this.countRows = countRows;
    }

    private static void validateHeaderRows(int headerRows) {
        if (headerRows < 1) {
            throw new IllegalArgumentException("headerRows must be >= 1");
        }
    }

    private static void validateSheetIndex(int sheetIndex) {
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("sheetIndex must be non-negative");
        }
    }

    /**
     * Starts parsing the Excel file and invokes the given consumer for each row result.
     * <p>
     * Each row is converted into a target object via the configured column setters.
     * Validation (if enabled) is performed after mapping.
     *
     * @param consumer Callback to receive parsed and validated row results
     */
    @Override
    public void read(Consumer<ReadResult<T>> consumer) {
        markConsumed();
        try {
            readInternal(guardedConsumer(consumer));
        } catch (io.github.dornol.excelkit.core.ReadStoppedException ignored) {
            stoppedEarly = true;
        } catch (io.github.dornol.excelkit.core.ReadLimitExceededException |
                 io.github.dornol.excelkit.core.ReadSecurityException e) {
            throw e;
        } catch (ExcelReadException | ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read excel", e);
        } finally {
            notifyReadCompletion(sheetIndex, -1);
            close();
        }
    }

    @Override
    public void readWhile(Predicate<ReadResult<T>> predicate) {
        markConsumed();
        java.util.concurrent.atomic.AtomicReference<RuntimeException> failure = new java.util.concurrent.atomic.AtomicReference<>();
        try {
            Consumer<ReadResult<T>> guarded = guardedConsumer(result -> {
                boolean proceed;
                try {
                    proceed = predicate.test(result);
                } catch (RuntimeException e) {
                    failure.set(e);
                    throw new io.github.dornol.excelkit.core.ReadStoppedException();
                }
                if (!proceed) throw new io.github.dornol.excelkit.core.ReadStoppedException();
            });
            readInternal(guarded);
        } catch (io.github.dornol.excelkit.core.ReadStoppedException ignored) {
            // Normal early completion requested by the caller.
            stoppedEarly = true;
        } catch (io.github.dornol.excelkit.core.ReadLimitExceededException |
                 io.github.dornol.excelkit.core.ReadSecurityException e) {
            throw e;
        } catch (ExcelReadException | ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read excel", e);
        } finally {
            notifyReadCompletion(sheetIndex, -1);
            close();
        }
        if (failure.get() != null) throw failure.get();
    }

    boolean wasStoppedEarly() { return stoppedEarly; }

    private void readInternal(Consumer<ReadResult<T>> consumer) throws Exception {
        Path fileToRead = getTempFile();
        Path decryptedFile = null;
        if (password != null) {
            decryptedFile = decryptFile(getTempFile(), password);
            fileToRead = decryptedFile;
        }
        enforceSecurityPolicy(fileToRead);
        try (OPCPackage pkg = OPCPackage.open(fileToRead.toFile())) {
            long totalRows = -1;
            if (countRows) {
                totalRows = preScanRowCount(pkg);
            }

            XSSFReader reader = new XSSFReader(pkg);
            if (limits.maxSheets() >= 0) {
                int sheetCount = 0;
                Iterator<InputStream> countIterator = reader.getSheetsData();
                while (countIterator.hasNext()) {
                    try (InputStream ignored = countIterator.next()) { sheetCount++; }
                    if (sheetCount > limits.maxSheets()) {
                        throw new io.github.dornol.excelkit.core.ReadLimitExceededException(
                                io.github.dornol.excelkit.core.ReadLimitExceededException.Limit.SHEETS,
                                limits.maxSheets(), sheetCount);
                    }
                }
            }

            SharedStrings ss = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            XMLReader parser = XMLHelper.newXMLReader();
            SheetHandler sheetHandler = new SheetHandler(consumer, totalRows);
            XSSFSheetXMLHandler sheetParser = new XSSFSheetXMLHandler(styles, ss, sheetHandler, false);
            parser.setContentHandler(sheetParser);

            Iterator<InputStream> sheetsData = reader.getSheetsData();
            int currentIndex = 0;
            while (sheetsData.hasNext()) {
                try (InputStream sheet = sheetsData.next()) {
                    if (currentIndex == sheetIndex) {
                        try {
                            parser.parse(new InputSource(sheet));
                        } catch (StopReadingException ignored) {
                            // Reader limit or blank-row stop condition reached.
                        }
                        break;
                    }
                }
                currentIndex++;
            }
            if (currentIndex < sheetIndex) {
                throw new ExcelReadException("Sheet index " + sheetIndex + " not found. File has " + (currentIndex + 1) + " sheet(s).");
            }
        } finally {
            if (decryptedFile != null) {
                try {
                    Files.deleteIfExists(decryptedFile);
                } catch (IOException e) {
                    log.warn("Failed to delete decrypted temp file: {}", decryptedFile, e);
                    decryptedFile.toFile().deleteOnExit();
                }
            }
        }
    }

    private void enforceSecurityPolicy(Path file) throws IOException {
        if (securityPolicy.allowFormulas() && securityPolicy.allowExternalLinks()) return;
        long totalScanned = 0;
        try (ZipFile zip = new ZipFile(file.toFile())) {
            var entries = zip.entries();
            while (entries.hasMoreElements()) {
                var entry = entries.nextElement();
                String name = entry.getName();
                if (!securityPolicy.allowExternalLinks() && name.startsWith("xl/externalLinks/")) {
                    throw new io.github.dornol.excelkit.core.ReadSecurityException(
                            io.github.dornol.excelkit.core.ReadSecurityException.Reason.EXTERNAL_LINK,
                            "External workbook links are not allowed");
                }
                if (!securityPolicy.allowFormulas() && name.startsWith("xl/worksheets/") && name.endsWith(".xml")) {
                    validateZipEntry(entry);
                    try (InputStream input = zip.getInputStream(entry)) {
                        ScanResult scan = scanFormulaTag(input, securityPolicy.maxScannedEntryBytes());
                        totalScanned += scan.bytes();
                        if (totalScanned > securityPolicy.maxTotalScannedBytes()) {
                            throw new io.github.dornol.excelkit.core.ReadSecurityException(
                                    io.github.dornol.excelkit.core.ReadSecurityException.Reason.TOTAL_SCAN_SIZE,
                                    "Workbook security scan exceeds total byte limit");
                        }
                        if (scan.formula()) throw new io.github.dornol.excelkit.core.ReadSecurityException(
                                io.github.dornol.excelkit.core.ReadSecurityException.Reason.FORMULA,
                                "Excel formulas are not allowed");
                    }
                }
            }
        }
    }

    private void validateZipEntry(java.util.zip.ZipEntry entry) {
        long size = entry.getSize();
        long compressed = entry.getCompressedSize();
        if (size > securityPolicy.maxScannedEntryBytes()) throw new io.github.dornol.excelkit.core.ReadSecurityException(
                io.github.dornol.excelkit.core.ReadSecurityException.Reason.ENTRY_SIZE,
                "Worksheet XML exceeds security scan entry limit: " + size);
        if (size > 0 && compressed > 0 && (double) size / compressed > securityPolicy.maxCompressionRatio())
            throw new io.github.dornol.excelkit.core.ReadSecurityException(
                    io.github.dornol.excelkit.core.ReadSecurityException.Reason.COMPRESSION_RATIO,
                    "Worksheet XML exceeds compression ratio limit");
    }

    private static ScanResult scanFormulaTag(InputStream input, long maximum) throws IOException {
        int state = 0;
        long bytes = 0;
        for (int value; (value = input.read()) >= 0;) {
            if (++bytes > maximum) throw new io.github.dornol.excelkit.core.ReadSecurityException(
                    io.github.dornol.excelkit.core.ReadSecurityException.Reason.ENTRY_SIZE,
                    "Worksheet XML exceeds security scan entry limit");
            if (state == 0) state = value == '<' ? 1 : 0;
            else if (state == 1) state = value == 'f' ? 2 : (value == '<' ? 1 : 0);
            else {
                if (value == '>' || Character.isWhitespace(value)) return new ScanResult(true, bytes);
                state = value == '<' ? 1 : 0;
            }
        }
        return new ScanResult(false, bytes);
    }

    private record ScanResult(boolean formula, long bytes) {}

    /**
     * Performs a lightweight SAX pre-scan to count data rows (excluding header rows).
     */
    private long preScanRowCount(OPCPackage pkg) throws Exception {
        XSSFReader reader = new XSSFReader(pkg);
        Iterator<InputStream> sheetsData = reader.getSheetsData();
        int currentIndex = 0;
        while (sheetsData.hasNext()) {
            try (InputStream sheet = sheetsData.next()) {
                if (currentIndex == sheetIndex) {
                    XMLReader xmlReader = XMLHelper.newXMLReader();
                    RowCountHandler counter = new RowCountHandler(headerRowIndex);
                    xmlReader.setContentHandler(counter);
                    xmlReader.parse(new InputSource(sheet));
                    return counter.getDataRowCount();
                }
            }
            currentIndex++;
        }
        return -1;
    }

    /**
     * Lightweight SAX handler that counts only data rows (after the header row).
     */
    private static class RowCountHandler extends DefaultHandler {
        private final int headerRowIndex;
        private long dataRowCount;
        private int currentRow = -1;

        RowCountHandler(int headerRowIndex) {
            this.headerRowIndex = headerRowIndex;
        }

        @Override
        public void startElement(String uri, String localName, String qName,
                                 org.xml.sax.Attributes attributes) {
            if ("row".equals(qName) || "row".equals(localName)) {
                currentRow++;
                if (currentRow > headerRowIndex) {
                    dataRowCount++;
                }
            }
        }

        long getDataRowCount() {
            return dataRowCount;
        }
    }

    private static final class StopReadingException extends RuntimeException {
    }


    private Path decryptFile(Path encryptedFile, String pwd) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(encryptedFile.toFile())) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor dec = Decryptor.getInstance(info);
            if (!dec.verifyPassword(pwd)) {
                throw new ExcelReadException("Invalid password for encrypted Excel file");
            }
            Path decryptedFile = TempResourceCreator.createTempFile(
                    getTempDir(), UUID.randomUUID().toString(), ".xlsx");
            try (InputStream decryptedStream = dec.getDataStream(fs);
                 OutputStream out = Files.newOutputStream(decryptedFile)) {
                decryptedStream.transferTo(out);
            }
            return decryptedFile;
        }
    }

    /**
     * Internal handler for row-by-row Excel parsing.
     */
    private class SheetHandler extends DefaultHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private @Nullable T currentInstance;
        private final List<CellData> currentRow = new ArrayList<>();
        private final List<String> headerNames = new ArrayList<>();
        /** Accumulates bottom-most non-blank header value per column across multi-row headers. */
        private final List<@Nullable String> headerAccumulator = new ArrayList<>();
        private final Consumer<ReadResult<T>> consumer;
        private @Nullable List<String> messages;
        private @Nullable List<io.github.dornol.excelkit.core.CellError> cellErrors;
        private int @Nullable [] resolvedIndices;
        private @Nullable Map<String, Integer> headerIndexMap;
        private long dataRowCount;
        private long emittedRowCount;
        private int consecutiveBlankRows;
        private final @Nullable Cursor cursor;

        public SheetHandler(Consumer<ReadResult<T>> consumer, long totalRows) {
            this.consumer = consumer;
            if (progressCallback != null) {
                this.cursor = new Cursor();
                if (totalRows >= 0) {
                    this.cursor.setTotalRows(totalRows);
                }
            } else {
                this.cursor = null;
            }
        }

        /**
         * Called at the start of each row. Resets the instance and message buffer.
         */
        @Override
        public void startRow(int rowNum) {
            if (cancellationToken.isCancellationRequested()) {
                throw new io.github.dornol.excelkit.core.ReadStoppedException();
            }
            if (instanceSupplier != null) {
                currentInstance = instanceSupplier.get();
            }
            currentRow.clear();
            messages = null;
            cellErrors = null;
        }

        /**
         * Called at the end of each row.
         * <p>
         * - Row at headerRowIndex is treated as the header.
         * - Later rows are mapped to the target object, validated (if applicable), and passed to consumer.
         */
        @Override
        public void endRow(int rowNum) {
            int firstHeaderRow = headerRowIndex - headerRows + 1;
            if (rowNum < firstHeaderRow) {
                return;
            }
            if (rowNum >= firstHeaderRow && rowNum <= headerRowIndex) {
                accumulateHeaderRow();
                if (rowNum == headerRowIndex) {
                    finalizeHeaderNames();
                    if (rowMapper != null) {
                        buildHeaderIndex();
                    } else {
                        resolveColumnIndices();
                    }
                }
                return;
            }

            List<String> rawValues = rawValues(currentRow);
            if (isBlankValues(rawValues)) {
                consecutiveBlankRows++;
                if (stopAtBlankRows > 0 && consecutiveBlankRows >= stopAtBlankRows) {
                    throw new StopReadingException();
                }
                if (skipBlankRows) {
                    return;
                }
            } else {
                consecutiveBlankRows = 0;
            }
            if (maxRows >= 0 && emittedRowCount >= maxRows) {
                throw new StopReadingException();
            }

            ReadResult<T> result;
            if (rowMapper != null) {
                RowData rowData = new RowData(new ArrayList<>(currentRow), headerNames, headerIndexMap, headerNormalizer);
                result = mapWithRowMapper(rowData, rowNum + 1L, rawValues);
            } else {
                boolean mappingSuccess = mapValuesToInstance();
                boolean validationSuccess = mappingSuccess && validateIfNeeded(currentInstance, getOrCreateMessages());
                result = new ReadResult<>(currentInstance, validationSuccess, messages, null, rowNum + 1L,
                        cellErrors == null ? List.of() : cellErrors, rawValues);
            }

            consumer.accept(result);

            dataRowCount++;
            emittedRowCount++;
            if (progressCallback != null && progressInterval > 0
                    && dataRowCount % progressInterval == 0) {
                progressCallback.onProgress(dataRowCount, cursor);
            }
            if (progressInterval > 0 && dataRowCount % progressInterval == 0) {
                notifyReadProgress(dataRowCount, sheetIndex, cursor == null ? -1 : cursor.getTotalRows());
            }
        }

        /**
         * Called for each cell in the current row.
         */
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int colIndex = getColumnIndex(cellReference);
            while (currentRow.size() < colIndex) {
                currentRow.add(cellData(currentRow.size(), null));
            }
            currentRow.add(cellData(colIndex, formattedValue));
        }

        /**
         * Merges the current header row's cells into {@link #headerAccumulator}, keeping
         * the bottom-most non-blank value per column (so a row below overrides a row above
         * only when it carries a value — preserving group labels whose column header cell
         * is blank due to a vertical merge in the source file).
         */
        private void accumulateHeaderRow() {
            for (CellData cell : currentRow) {
                int idx = cell.columnIndex();
                while (headerAccumulator.size() <= idx) {
                    headerAccumulator.add(null);
                }
                String v = cell.formattedValue();
                if (headerRows == 1) {
                    // Single-row: preserve legacy behavior — record every value (including "").
                    if (v != null) {
                        headerAccumulator.set(idx, v);
                    }
                } else if (v != null && !v.isEmpty()) {
                    // Multi-row: non-blank overrides only, so vertically merged group labels
                    // survive when the column header cell below is emitted as blank.
                    headerAccumulator.set(idx, v);
                }
            }
        }

        /**
         * Finalizes the accumulated header names.
         */
        private void finalizeHeaderNames() {
            headerNames.addAll(headerAccumulator);
        }

        /**
         * Resolves named columns to their actual indices based on header names (setter mode).
         */
        private void resolveColumnIndices() {
            if (columns == null) {
                throw new IllegalStateException("columns must not be null in setter mode");
            }
            resolvedIndices = ExcelReadHandler.this.resolveColumnIndices(
                    columns.size(),
                    i -> columns.get(i).headerAliases(),
                    i -> columns.get(i).columnIndex(),
                    headerNames, "sheet"
            );
        }

        /**
         * Builds header name to index map (mapping mode).
         */
        private void buildHeaderIndex() {
            headerIndexMap = buildHeaderIndexMap(headerNames, "sheet");
            validateSelectedMapColumns(headerIndexMap, headerNames, "sheet");
        }

        /**
         * Applies all column setters to the current row data (setter mode).
         *
         * @return true if all setters succeeded, false if any failed
         */
        private boolean mapValuesToInstance() {
            if (columns == null || resolvedIndices == null) {
                throw new IllegalStateException("columns and resolvedIndices must not be null in setter mode");
            }
            boolean success = true;

            for (int i = 0; i < columns.size(); i++) {
                int actualIndex = resolvedIndices[i];
                if (actualIndex >= currentRow.size()) {
                    if (columns.get(i).isRequired()) {
                        String header = (actualIndex < headerNames.size()) ? headerNames.get(actualIndex) : "column#" + actualIndex;
                        String message = "Required column '" + header + "' is empty";
                        getOrCreateMessages().add(message);
                        getOrCreateCellErrors().add(new io.github.dornol.excelkit.core.CellError(actualIndex, header, null, message));
                        success = false;
                    }
                    continue;
                }

                if (!mapColumn(columns.get(i), currentInstance, currentRow.get(actualIndex),
                        actualIndex, headerNames, getOrCreateMessages(), getOrCreateCellErrors())) {
                    success = false;
                }
            }

            return success;
        }

        private List<String> getOrCreateMessages() {
            if (messages == null) {
                messages = new ArrayList<>();
            }
            return messages;
        }

        private List<io.github.dornol.excelkit.core.CellError> getOrCreateCellErrors() {
            if (cellErrors == null) {
                cellErrors = new ArrayList<>();
            }
            return cellErrors;
        }

        private int getColumnIndex(String cellReference) {
            return ExcelReadSupport.getColumnIndex(cellReference);
        }

    }
}
