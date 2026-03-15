package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.InputStream;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Stream;

/**
 * Convenience reader for parsing Excel files into {@code Map<String, String>} rows.
 * <p>
 * Automatically maps all columns found in the header row to map entries.
 * Useful when the column structure is not known at compile time.
 *
 * <pre>{@code
 * new ExcelMapReader()
 *     .build(inputStream)
 *     .read(result -> {
 *         Map<String, String> row = result.data();
 *         String name = row.get("Name");
 *     });
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelMapReader {

    private int sheetIndex = 0;
    private int headerRowIndex = 0;

    /**
     * Sets the sheet index to read (0-based).
     */
    public ExcelMapReader sheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    /**
     * Sets the header row index (0-based).
     */
    public ExcelMapReader headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    /**
     * Builds a handler for reading the Excel file.
     *
     * @param inputStream the Excel file input stream
     * @return a handler to execute reading
     */
    public ExcelMapReadHandler build(InputStream inputStream) {
        return new ExcelMapReadHandler(inputStream, sheetIndex, headerRowIndex);
    }

    /**
     * Handler for reading Excel data into maps.
     * All columns discovered in the header row are automatically mapped.
     */
    public static class ExcelMapReadHandler extends TempResourceContainer {
        private final int sheetIndex;
        private final int headerRowIndex;

        ExcelMapReadHandler(InputStream inputStream, int sheetIndex, int headerRowIndex) {
            this.sheetIndex = sheetIndex;
            this.headerRowIndex = headerRowIndex;
            initTempFile(inputStream);
        }

        private void initTempFile(InputStream inputStream) {
            try {
                setTempDir(TempResourceCreator.createTempDirectory());
                setTempFile(TempResourceCreator.createTempFile(getTempDir(),
                        UUID.randomUUID().toString(), ".xlsx"));
                try (InputStream is = inputStream) {
                    java.nio.file.Files.copy(is, getTempFile(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                }
            } catch (java.io.IOException e) {
                throw new ExcelKitException("Failed to initialize temporary file", e);
            }
        }

        /**
         * Reads the Excel file, invoking the consumer for each row.
         */
        public void read(Consumer<ReadResult<Map<String, String>>> consumer) {
            try {
                readInternal(consumer);
            } catch (ExcelReadException e) {
                throw e;
            } catch (Exception e) {
                throw new ExcelReadException("Failed to read excel", e);
            } finally {
                close();
            }
        }

        /**
         * Reads the Excel file and returns a stream of map results.
         */
        public Stream<ReadResult<Map<String, String>>> readAsStream() {
            List<ReadResult<Map<String, String>>> results = new ArrayList<>();
            read(results::add);
            return results.stream();
        }

        private void readInternal(Consumer<ReadResult<Map<String, String>>> consumer) throws Exception {
            try (OPCPackage pkg = OPCPackage.open(getTempFile().toFile())) {
                XSSFReader reader = new XSSFReader(pkg);
                SharedStrings ss = reader.getSharedStringsTable();
                StylesTable styles = reader.getStylesTable();

                XMLReader parser = XMLHelper.newXMLReader();
                MapSheetHandler handler = new MapSheetHandler(consumer);
                XSSFSheetXMLHandler sheetParser = new XSSFSheetXMLHandler(styles, ss, handler, false);
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
            }
        }

        private class MapSheetHandler extends DefaultHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
            private final Consumer<ReadResult<Map<String, String>>> consumer;
            private final List<String> headerNames = new ArrayList<>();
            private final List<CellData> currentRow = new ArrayList<>();

            MapSheetHandler(Consumer<ReadResult<Map<String, String>>> consumer) {
                this.consumer = consumer;
            }

            @Override
            public void startRow(int rowNum) {
                currentRow.clear();
            }

            @Override
            public void endRow(int rowNum) {
                if (rowNum < headerRowIndex) return;

                if (rowNum == headerRowIndex) {
                    headerNames.addAll(currentRow.stream()
                            .map(CellData::formattedValue)
                            .filter(Objects::nonNull)
                            .toList());
                    return;
                }

                Map<String, String> map = new LinkedHashMap<>();
                for (int i = 0; i < headerNames.size() && i < currentRow.size(); i++) {
                    CellData cell = currentRow.get(i);
                    map.put(headerNames.get(i), cell != null ? cell.formattedValue() : null);
                }
                consumer.accept(new ReadResult<>(map, true, null));
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                int colIndex = getColumnIndex(cellReference);
                while (currentRow.size() < colIndex) {
                    currentRow.add(new CellData(currentRow.size(), null));
                }
                currentRow.add(new CellData(colIndex, formattedValue));
            }

            private int getColumnIndex(String cellReference) {
                int colIdx = 0;
                for (char c : cellReference.toCharArray()) {
                    if (!Character.isLetter(c)) break;
                    colIdx = colIdx * 26 + (Character.toUpperCase(c) - 'A' + 1);
                }
                return colIdx - 1;
            }
        }
    }
}
