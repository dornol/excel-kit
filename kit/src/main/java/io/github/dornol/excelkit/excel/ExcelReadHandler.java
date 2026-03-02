package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.AbstractReadHandler;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.jspecify.annotations.NonNull;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Reads Excel (.xlsx) files using Apache POI's event-based streaming API.
 * <p>
 * This handler parses sheet data row by row, maps values to Java objects, and performs optional validation.
 * It is optimized for large files and avoids loading the entire workbook into memory.
 * <p>
 * For large or complex Excel files, you may need to adjust POI's internal limits via
 * {@link ExcelReader#configureLargeFileSupport()} before reading. This adjusts:
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
    private final List<ExcelReadColumn<T>> columns;
    private final int sheetIndex;
    private final int headerRowIndex;

    /**
     * Constructs a handler for reading the first sheet of an Excel file.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     */
    ExcelReadHandler(InputStream inputStream, List<ExcelReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, 0);
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
    ExcelReadHandler(InputStream inputStream, List<ExcelReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int sheetIndex) {
        this(inputStream, columns, instanceSupplier, validator, sheetIndex, 0);
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
    ExcelReadHandler(InputStream inputStream, List<ExcelReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int sheetIndex, int headerRowIndex) {
        super(inputStream, instanceSupplier, validator, ".xlsx");
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
        if (sheetIndex < 0 || sheetIndex > 255) {
            throw new IllegalArgumentException("sheetIndex must be between 0 and 255");
        }
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
        this.columns = columns;
        this.sheetIndex = sheetIndex;
        this.headerRowIndex = headerRowIndex;
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
    public void read(@NonNull Consumer<ReadResult<T>> consumer) {
        try {
            readInternal(consumer);
        } catch (ExcelReadException e) {
            throw e;
        } catch (ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("Failed to read excel", e);
        } finally {
            close();
        }
    }

    @Override
    public Stream<ReadResult<T>> readAsStream() {
        int bufferSize = 1024;
        BlockingQueue<Object> queue = new ArrayBlockingQueue<>(bufferSize);
        Object sentinel = new Object();
        AtomicReference<Throwable> producerError = new AtomicReference<>();

        Thread producer = new Thread(() -> {
            try {
                readInternal(result -> {
                    try {
                        queue.put(result);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                        throw new ExcelReadException("Producer thread interrupted", e);
                    }
                });
            } catch (Throwable t) {
                producerError.set(t);
            } finally {
                try {
                    queue.put(sentinel);
                } catch (InterruptedException ignored) {
                    Thread.currentThread().interrupt();
                }
            }
        });
        producer.setDaemon(true);
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
        try (OPCPackage pkg = OPCPackage.open(getTempFile().toFile())) {
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
        }
    }


    /**
     * Internal handler for row-by-row Excel parsing.
     */
    private class SheetHandler extends DefaultHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private T currentInstance;
        private final List<CellData> currentRow = new ArrayList<>();
        private final List<String> headerNames = new ArrayList<>();
        private final Consumer<ReadResult<T>> consumer;
        private List<String> messages;

        public SheetHandler(Consumer<ReadResult<T>> consumer) {
            this.consumer = consumer;
        }

        /**
         * Called at the start of each row. Resets the instance and message buffer.
         */
        @Override
        public void startRow(int rowNum) {
            currentInstance = instanceSupplier.get();
            currentRow.clear();
            messages = null;
        }

        /**
         * Called at the end of each row.
         * <p>
         * - Row 0 is treated as the header.
         * - Later rows are mapped to the target object, validated (if applicable), and passed to consumer.
         */
        @Override
        public void endRow(int rowNum) {
            if (rowNum < headerRowIndex) {
                return;
            }
            if (rowNum == headerRowIndex) {
                extractHeaderNames();
                return;
            }

            boolean mappingSuccess = mapValuesToInstance();
            boolean validationSuccess = mappingSuccess && validateIfNeeded(currentInstance, getOrCreateMessages());

            consumer.accept(new ReadResult<>(currentInstance, validationSuccess, messages));
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
         * Extracts header names from the first row.
         */
        private void extractHeaderNames() {
            headerNames.addAll(currentRow.stream()
                    .map(CellData::formattedValue)
                    .toList());
        }

        /**
         * Applies all column setters to the current row data.
         *
         * @return true if all setters succeeded, false if any failed
         */
        private boolean mapValuesToInstance() {
            boolean success = true;

            for (int i = 0; i < columns.size(); i++) {
                if (i >= currentRow.size()) continue;

                if (!mapColumn(columns.get(i).setter(), currentInstance, currentRow.get(i),
                        i, headerNames, getOrCreateMessages())) {
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

        /**
         * Converts an Excel cell reference (e.g., "C5", "AA12") to a zero-based column index.
         * <p>
         * Only the alphabetic part (column letters) is used. For example:
         * <ul>
         *   <li>"A1"  -> 0</li>
         *   <li>"B3"  -> 1</li>
         *   <li>"C5"  -> 2</li>
         *   <li>"AA10"-> 26</li>
         * </ul>
         *
         * @param cellReference The Excel cell reference (e.g., "C5", "AA10")
         * @return The zero-based column index
         */
        private int getColumnIndex(String cellReference) {
            int colIdx = 0;
            for (char c : cellReference.toCharArray()) {
                if (!Character.isLetter(c)) break;
                colIdx = colIdx * 26 + (Character.toUpperCase(c) - 'A' + 1);
                if (colIdx > 16_384) { // Excel max column: XFD = 16,384
                    throw new ExcelReadException("Column index exceeds Excel maximum (XFD): " + cellReference);
                }
            }
            return colIdx - 1;
        }

    }
}
