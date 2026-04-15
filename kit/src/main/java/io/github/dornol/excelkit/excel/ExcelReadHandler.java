package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.AbstractReadHandler;
import io.github.dornol.excelkit.core.ReadColumn;
import io.github.dornol.excelkit.core.CellData;
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
import java.util.HashSet;
import java.util.Iterator;
import java.util.UUID;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Reads Excel (.xlsx) files using Apache POI's event-based streaming API.
 * <p>
 * This handler parses sheet data row by row, maps values to Java objects, and performs optional validation.
 * It is optimized for large files and avoids loading the entire workbook into memory.
 *
 * <h2>Resource management</h2>
 * On construction, the handler copies the input stream to a temporary file on disk so that
 * the underlying POI API can read it. These temp resources (and, if applicable, the decrypted
 * copy of an encrypted file) are released when:
 * <ul>
 *     <li>{@link #read(Consumer)} / {@link #readStrict(Consumer)} returns or throws — cleanup is automatic.</li>
 *     <li>The stream returned by {@link #readAsStream()} is closed — <strong>always use
 *         try-with-resources</strong>, since this stream also holds a background producer thread:
 *         <pre>{@code
 * try (Stream<ReadResult<T>> stream = handler.readAsStream()) {
 *     stream.forEach(result -> ...);
 * }
 *         }</pre>
 *         Abandoning the stream without closing it leaks the temp file until the JVM exits
 *         (the producer thread is a daemon and will eventually self-terminate).</li>
 * </ul>
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
public class ExcelReadHandler<T> extends AbstractReadHandler<T> {
    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(ExcelReadHandler.class);

    private final @Nullable List<ReadColumn<T>> columns;
    private final int sheetIndex;
    private final int headerRowIndex;
    private final int headerRows;
    private final int progressInterval;
    private final @Nullable ProgressCallback progressCallback;
    private final @Nullable String password;

    /**
     * Constructs a handler for reading the first sheet of an Excel file.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     */
    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, 0, 1, 0, null, null);
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
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, 0, 1, 0, null, null);
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
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, 1, 0, null, null);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int progressInterval, @Nullable ProgressCallback progressCallback) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, headerRowIndex, 1, progressInterval, progressCallback, (String) null);
    }

    ExcelReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                     Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password) {
        super(inputStream, instanceSupplier, validator, ".xlsx");
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
    }

    /**
     * Constructs a handler in mapping mode for immutable object construction.
     */
    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int progressInterval, @Nullable ProgressCallback progressCallback) {
        this(inputStream, rowMapper, validator, sheetIndex, headerRowIndex, 1, progressInterval, progressCallback, (String) null);
    }

    ExcelReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                     @Nullable Validator validator, int sheetIndex, int headerRowIndex,
                     int headerRows,
                     int progressInterval, @Nullable ProgressCallback progressCallback,
                     @Nullable String password) {
        super(inputStream, rowMapper, validator, ".xlsx");
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
    }

    private static void validateHeaderRows(int headerRows) {
        if (headerRows < 1) {
            throw new IllegalArgumentException("headerRows must be >= 1");
        }
    }

    private static void validateSheetIndex(int sheetIndex) {
        if (sheetIndex < 0 || sheetIndex > 255) {
            throw new IllegalArgumentException("sheetIndex must be between 0 and 255");
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
        try {
            readInternal(consumer);
        } catch (ExcelReadException | ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read excel", e);
        } finally {
            close();
        }
    }

    /**
     * Reads the file as a stream of row results using a background producer thread.
     * <p>
     * <strong>Important:</strong> The returned stream holds file and thread resources.
     * Always use try-with-resources to ensure proper cleanup:
     * <pre>{@code
     * try (Stream<ReadResult<T>> stream = handler.readAsStream()) {
     *     stream.forEach(result -> ...);
     * }
     * }</pre>
     *
     * @return A stream of parsed and validated row results
     */
    @Override
    public Stream<ReadResult<T>> readAsStream() {
        int bufferSize = 1024;
        BlockingQueue<Object> queue = new ArrayBlockingQueue<>(bufferSize);
        Object sentinel = new Object();
        AtomicReference<Throwable> producerError = new AtomicReference<>();

        Thread producer = new Thread(() -> {
            try {
                readInternal(result -> {
                    if (Thread.currentThread().isInterrupted()) {
                        throw new ExcelReadException("Producer thread interrupted");
                    }
                    try {
                        // Use offer with timeout to avoid permanent block when consumer closes early.
                        // If the queue is full and consumer stopped draining, the offer will time out
                        // and the interrupt check above will catch it on the next row.
                        while (!queue.offer(result, 100, java.util.concurrent.TimeUnit.MILLISECONDS)) {
                            if (Thread.currentThread().isInterrupted()) {
                                throw new ExcelReadException("Producer thread interrupted");
                            }
                        }
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                        throw new ExcelReadException("Producer thread interrupted", e);
                    }
                });
            } catch (Throwable t) {
                producerError.set(t);
            } finally {
                // Use offer to avoid blocking forever if queue is full and consumer is gone.
                // Sentinel delivery is best-effort; consumer also checks producerError on interrupt.
                queue.offer(sentinel);
            }
        });
        // Daemon thread: if the caller abandons the stream without close(),
        // the producer eventually exits via interrupt check or offer timeout.
        // Normal cleanup path is stream.onClose() → producer.interrupt().
        producer.setDaemon(true);
        producer.setName("excel-kit-reader");
        producer.start();

        Spliterator<ReadResult<T>> spliterator = new Spliterators.AbstractSpliterator<>(
                Long.MAX_VALUE, Spliterator.ORDERED | Spliterator.NONNULL) {
            @SuppressWarnings("unchecked")
            @Override
            public boolean tryAdvance(Consumer<? super ReadResult<T>> action) {
                try {
                    Object item = queue.take();
                    if (item == sentinel) {
                        Throwable error = producerError.get();
                        if (error != null) {
                            if (error instanceof ExcelReadException e) throw e;
                            if (error instanceof ReadAbortException e) throw e;
                            throw new ExcelReadException("Failed to read excel", error);
                        }
                        return false;
                    }
                    action.accept((ReadResult<T>) item);
                    return true;
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    throw new ExcelReadException("Consumer thread interrupted", e);
                }
            }
        };

        return StreamSupport.stream(spliterator, false)
                .onClose(() -> {
                    producer.interrupt();
                    close();
                });
    }

    private void readInternal(Consumer<ReadResult<T>> consumer) throws Exception {
        Path fileToRead = getTempFile();
        Path decryptedFile = null;
        if (password != null) {
            decryptedFile = decryptFile(getTempFile(), password);
            fileToRead = decryptedFile;
        }
        try (OPCPackage pkg = OPCPackage.open(fileToRead.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);

            SharedStrings ss = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            XMLReader parser = XMLHelper.newXMLReader();
            SheetHandler sheetHandler = new SheetHandler(consumer);
            XSSFSheetXMLHandler sheetParser = new XSSFSheetXMLHandler(styles, ss, sheetHandler, false);
            parser.setContentHandler(sheetParser);

            Iterator<InputStream> sheetsData = reader.getSheetsData();
            int currentIndex = 0;
            while (sheetsData.hasNext()) {
                try (InputStream sheet = sheetsData.next()) {
                    if (currentIndex == sheetIndex) {
                        parser.parse(new InputSource(sheet));
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
        private int @Nullable [] resolvedIndices;
        private @Nullable Map<String, Integer> headerIndexMap;
        private long dataRowCount;

        public SheetHandler(Consumer<ReadResult<T>> consumer) {
            this.consumer = consumer;
        }

        /**
         * Called at the start of each row. Resets the instance and message buffer.
         */
        @Override
        public void startRow(int rowNum) {
            if (instanceSupplier != null) {
                currentInstance = instanceSupplier.get();
            }
            currentRow.clear();
            messages = null;
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

            ReadResult<T> result;
            if (rowMapper != null) {
                RowData rowData = new RowData(new ArrayList<>(currentRow), headerNames, headerIndexMap);
                result = mapWithRowMapper(rowData);
            } else {
                boolean mappingSuccess = mapValuesToInstance();
                boolean validationSuccess = mappingSuccess && validateIfNeeded(currentInstance, getOrCreateMessages());
                result = new ReadResult<>(currentInstance, validationSuccess, messages);
            }

            consumer.accept(result);

            dataRowCount++;
            if (progressCallback != null && progressInterval > 0
                    && dataRowCount % progressInterval == 0) {
                progressCallback.onProgress(dataRowCount, null);
            }
        }

        /**
         * Called for each cell in the current row.
         */
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int colIndex = getColumnIndex(cellReference);
            while (currentRow.size() < colIndex) {
                currentRow.add(new CellData(currentRow.size(), null));
            }
            currentRow.add(new CellData(colIndex, formattedValue));
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
         * Finalizes the accumulated header names and warns about duplicates.
         */
        private void finalizeHeaderNames() {
            headerNames.addAll(headerAccumulator);

            Set<String> seen = new HashSet<>();
            for (String name : headerNames) {
                if (name != null && !seen.add(name)) {
                    log.warn("Duplicate header name '{}' found in sheet (headerRowIndex={}). "
                            + "Only the first occurrence will be used in mapping mode.", name, headerRowIndex);
                }
            }
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
                    i -> columns.get(i).headerName(),
                    i -> columns.get(i).columnIndex(),
                    headerNames, "sheet"
            );
        }

        /**
         * Builds header name to index map (mapping mode).
         */
        private void buildHeaderIndex() {
            headerIndexMap = new LinkedHashMap<>();
            for (int i = 0; i < headerNames.size(); i++) {
                headerIndexMap.putIfAbsent(headerNames.get(i), i);
            }
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
                        getOrCreateMessages().add("Required column '" + header + "' is empty");
                        success = false;
                    }
                    continue;
                }

                if (!mapColumn(columns.get(i), currentInstance, currentRow.get(actualIndex),
                        actualIndex, headerNames, getOrCreateMessages())) {
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

        private int getColumnIndex(String cellReference) {
            return ExcelReadSupport.getColumnIndex(cellReference);
        }

    }
}
