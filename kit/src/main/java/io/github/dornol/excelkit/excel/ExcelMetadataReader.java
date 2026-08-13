package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.InputStreamSource;
import io.github.dornol.excelkit.core.TempResourceCreator;
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

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.UUID;

/** Reads workbook metadata without coupling it to row reader configuration. */
final class ExcelMetadataReader {
    private ExcelMetadataReader() {
    }

    static List<ExcelSheetInfo> sheetNames(InputStream inputStream) {
        return withTempWorkbook(inputStream, "Failed to read sheet names", ExcelMetadataReader::readSheetNames);
    }

    static List<ExcelSheetInfo> sheetNames(InputStreamSource source) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return sheetNames(input);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    static List<String> headers(InputStream inputStream, int sheetIndex, int headerRowIndex) {
        validateIndexes(sheetIndex, headerRowIndex);
        return withTempWorkbook(inputStream, "Failed to read sheet headers",
                path -> readHeaders(path, sheetIndex, headerRowIndex));
    }

    static List<String> headers(InputStreamSource source, int sheetIndex, int headerRowIndex) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return headers(input, sheetIndex, headerRowIndex);
        } catch (IOException e) {
            throw new ExcelReadException("Failed to open Excel input", e);
        }
    }

    private static List<ExcelSheetInfo> readSheetNames(Path path) {
        try (OPCPackage pkg = OPCPackage.open(path.toFile())) {
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
        try (OPCPackage pkg = OPCPackage.open(path.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sharedStrings = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();
            XMLReader parser = XMLHelper.newXMLReader();
            HeaderExtractor extractor = new HeaderExtractor(headerRowIndex);
            parser.setContentHandler(new XSSFSheetXMLHandler(styles, sharedStrings, extractor, false));

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

    private static <R> R withTempWorkbook(InputStream inputStream, String message,
                                          WorkbookOperation<R> operation) {
        Objects.requireNonNull(inputStream, "inputStream cannot be null");
        Path tempDir = null;
        Path tempFile = null;
        try {
            tempDir = TempResourceCreator.createTempDirectory();
            tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".xlsx");
            Files.copy(inputStream, tempFile, StandardCopyOption.REPLACE_EXISTING);
            return operation.apply(tempFile);
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException(message, e);
        } finally {
            cleanup(tempDir, tempFile);
        }
    }

    private static void validateIndexes(int sheetIndex, int headerRowIndex) {
        if (sheetIndex < 0) throw new IllegalArgumentException("sheetIndex must be non-negative");
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex must be non-negative");
    }

    private static void cleanup(Path tempDir, Path tempFile) {
        try {
            if (tempFile != null) Files.deleteIfExists(tempFile);
        } catch (IOException e) {
            tempFile.toFile().deleteOnExit();
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

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
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
