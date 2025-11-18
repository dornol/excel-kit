package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.TempResourceContainer;
import io.github.dornol.excelkit.shared.TempResourceCreator;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.UUID;
import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * Reads Excel (.xlsx) files using Apache POI's event-based streaming API.
 * <p>
 * This handler parses sheet data row by row, maps values to Java objects, and performs optional validation.
 * It is optimized for large files and avoids loading the entire workbook into memory.
 *
 * <p>
 * For large or complex Excel files, the following POI internal limits are adjusted:
 * <ul>
 *     <li>{@code ZipSecureFile.setMaxFileCount(10_000_000)} — Increases the maximum number of internal file entries to avoid security exceptions for large files.</li>
 *     <li>{@code IOUtils.setByteArrayMaxOverride(2_000_000_000)} — Increases the maximum allowable in-memory byte array size to support large embedded binary data.</li>
 * </ul>
 * Be cautious when adjusting these values, as it may affect application memory usage and security.
 * </p>
 *
 * @param <T> The target row data type to map each row into
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelReadHandler<T> extends TempResourceContainer {
    private static final Logger log = LoggerFactory.getLogger(ExcelReadHandler.class);
    private final List<ExcelReadColumn<T>> columns;
    private final Supplier<T> instanceSupplier;
    private final Validator validator;

    /**
     * Constructs a handler for reading Excel files.
     *
     * @param inputStream      The input stream of the uploaded Excel file
     * @param columns          The list of column setters to apply per row
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     */
    ExcelReadHandler(InputStream inputStream, List<ExcelReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator) {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
        if (instanceSupplier == null) {
            throw new IllegalArgumentException("Instance supplier cannot be null");
        }
        this.columns = columns;
        this.instanceSupplier = instanceSupplier;
        this.validator = validator;
        try {
            setTempDir(TempResourceCreator.createTempDirectory());
            setTempFile(TempResourceCreator.createTempFile(getTempDir(), UUID.randomUUID().toString(), ".xlsx"));
            try (InputStream is = inputStream) {
                Files.copy(is, getTempFile(), StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
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
    public void read(Consumer<ExcelReadResult<T>> consumer) {
        try (OPCPackage pkg = OPCPackage.open(getTempFile().toFile())) {
            XSSFReader reader = new XSSFReader(pkg);

            SharedStrings ss = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            XMLReader parser = XMLHelper.newXMLReader();
            SheetHandler sheetHandler = new SheetHandler(consumer);
            XSSFSheetXMLHandler sheetParser = new XSSFSheetXMLHandler(styles, ss, sheetHandler, false);
            parser.setContentHandler(sheetParser);

            try (InputStream sheet = reader.getSheetsData().next()) {
                parser.parse(new InputSource(sheet));
            }

        } catch (Exception e) {
            throw new IllegalStateException("Failed to read excel", e);
        } finally {
            close();
        }
    }


    /**
     * Internal handler for row-by-row Excel parsing.
     */
    private class SheetHandler extends DefaultHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private T currentInstance;
        private final List<ExcelCellData> currentRow = new ArrayList<>();
        private final List<String> headerNames = new ArrayList<>();
        private final Consumer<ExcelReadResult<T>> consumer;
        private List<String> messages;

        public SheetHandler(Consumer<ExcelReadResult<T>> consumer) {
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
            if (rowNum == 0) {
                extractHeaderNames();
                return;
            }

            boolean mappingSuccess = mapValuesToInstance();
            boolean validationSuccess = mappingSuccess && validateIfNeeded();

            consumer.accept(new ExcelReadResult<>(currentInstance, validationSuccess, messages));
        }

        /**
         * Called for each cell in the current row.
         */
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int colIndex = getColumnIndex(cellReference);
            while (currentRow.size() < colIndex) {
                currentRow.add(new ExcelCellData(currentRow.size(), null));
            }
            currentRow.add(new ExcelCellData(colIndex, formattedValue));
        }

        /**
         * Extracts header names from the first row.
         */
        private void extractHeaderNames() {
            headerNames.addAll(currentRow.stream()
                    .map(ExcelCellData::formattedValue)
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

                try {
                    columns.get(i).setter().accept(currentInstance, currentRow.get(i));
                } catch (Exception e) {
                    success = false;
                    if (messages == null) {
                        messages = new ArrayList<>();
                    }
                    String header = (i < headerNames.size()) ? headerNames.get(i) : "column#" + i;
                    messages.add("Failed to set column: " + header);
                    log.warn("Column mapping failed", e);
                }
            }

            return success;
        }

        /**
         * Validates the current instance using Bean Validation (if enabled).
         *
         * @return true if valid, false if any constraint violations occurred
         */
        private boolean validateIfNeeded() {
            if (validator == null) {
                return true;
            }

            Set<ConstraintViolation<T>> violations = validator.validate(currentInstance);
            if (violations.isEmpty()) return true;

            if (messages == null) {
                messages = new ArrayList<>();
            }
            violations.stream()
                    .map(ConstraintViolation::getMessage)
                    .forEach(messages::add);

            return false;
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
            // 예: "C5" => 2 (0-based)
            StringBuilder sb = new StringBuilder();
            for (char c : cellReference.toCharArray()) {
                if (Character.isLetter(c)) sb.append(c);
                else break;
            }
            String col = sb.toString();
            int colIdx = 0;
            for (char c : col.toCharArray()) {
                colIdx = colIdx * 26 + (c - 'A' + 1);
            }
            return colIdx - 1;
        }

    }
}
