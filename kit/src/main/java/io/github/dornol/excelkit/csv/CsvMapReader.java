package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ProgressCallback;
import io.github.dornol.excelkit.shared.ReadResult;
import io.github.dornol.excelkit.shared.TempResourceContainer;
import io.github.dornol.excelkit.shared.TempResourceCreator;
import io.github.dornol.excelkit.shared.ExcelKitException;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.UUID;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Convenience reader for parsing CSV files into {@code Map<String, String>} rows.
 * <p>
 * Automatically maps all columns found in the header row to map entries.
 * Useful when the column structure is not known at compile time.
 *
 * <pre>{@code
 * new CsvMapReader()
 *     .build(inputStream)
 *     .read(result -> {
 *         Map<String, String> row = result.data();
 *         String name = row.get("Name");
 *     });
 * }</pre>
 *
 * @author dhkim
 * @since 0.9.3
 */
public class CsvMapReader {

    private int headerRowIndex = 0;
    private char delimiter = ',';
    private Charset charset = StandardCharsets.UTF_8;
    private @Nullable ProgressCallback progressCallback;
    private int progressInterval;

    /**
     * Applies a predefined CSV dialect configuration.
     *
     * @param dialect the dialect to apply
     * @return this instance for chaining
     */
    public CsvMapReader dialect(CsvDialect dialect) {
        this.delimiter = dialect.getDelimiter();
        this.charset = dialect.getCharset();
        return this;
    }

    /**
     * Sets the header row index (0-based).
     */
    public CsvMapReader headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    /**
     * Sets the delimiter character.
     */
    public CsvMapReader delimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    /**
     * Sets the character encoding.
     */
    public CsvMapReader charset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     *
     * @param interval the number of rows between each callback invocation (must be positive)
     * @param callback the callback to invoke
     * @return this instance for chaining
     */
    public CsvMapReader onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Builds a handler for reading the CSV file.
     *
     * @param inputStream the CSV file input stream
     * @return a handler to execute reading
     */
    public CsvMapReadHandler build(InputStream inputStream) {
        return new CsvMapReadHandler(inputStream, headerRowIndex, delimiter, charset,
                progressInterval, progressCallback);
    }

    /**
     * Handler for reading CSV data into maps.
     * All columns discovered in the header row are automatically mapped.
     */
    public static class CsvMapReadHandler extends TempResourceContainer {
        private static final Logger log = LoggerFactory.getLogger(CsvMapReadHandler.class);

        private final int headerRowIndex;
        private final char delimiter;
        private final Charset charset;
        private final int progressInterval;
        private final @Nullable ProgressCallback progressCallback;

        CsvMapReadHandler(InputStream inputStream, int headerRowIndex, char delimiter, Charset charset,
                          int progressInterval, @Nullable ProgressCallback progressCallback) {
            if (inputStream == null) {
                throw new IllegalArgumentException("InputStream cannot be null");
            }
            if (headerRowIndex < 0) {
                throw new IllegalArgumentException("headerRowIndex must be non-negative");
            }
            this.headerRowIndex = headerRowIndex;
            this.delimiter = delimiter;
            this.charset = charset;
            this.progressInterval = progressInterval;
            this.progressCallback = progressCallback;
            initTempFile(inputStream);
        }

        private void initTempFile(InputStream inputStream) {
            try {
                setTempDir(TempResourceCreator.createTempDirectory());
                setTempFile(TempResourceCreator.createTempFile(getTempDir(),
                        UUID.randomUUID().toString(), ".csv"));
                try (InputStream is = inputStream) {
                    Files.copy(is, getTempFile(), StandardCopyOption.REPLACE_EXISTING);
                }
            } catch (IOException e) {
                throw new ExcelKitException("Failed to initialize temporary file", e);
            }
        }

        /**
         * Reads the CSV file, invoking the consumer for each row.
         */
        public void read(Consumer<ReadResult<Map<String, String>>> consumer) {
            try (CSVReader reader = buildCsvReader()) {
                List<String> headerNames = readHeaders(reader);

                String[] line;
                long rowCount = 0;
                while ((line = reader.readNext()) != null) {
                    Map<String, String> map = new LinkedHashMap<>();
                    for (int i = 0; i < headerNames.size() && i < line.length; i++) {
                        map.put(headerNames.get(i), line[i]);
                    }
                    consumer.accept(new ReadResult<>(map, true, null));
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0
                            && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                }
            } catch (CsvReadException e) {
                throw e;
            } catch (Exception e) {
                throw new CsvReadException("Failed to read CSV", e);
            } finally {
                close();
            }
        }

        /**
         * Reads the CSV file as a stream of map results.
         * <p>
         * <strong>Important:</strong> The returned stream holds file resources.
         * Always use try-with-resources to ensure proper cleanup:
         * <pre>{@code
         * try (Stream<ReadResult<Map<String, String>>> stream = handler.readAsStream()) {
         *     stream.forEach(result -> ...);
         * }
         * }</pre>
         *
         * @return A stream of parsed row results
         */
        public Stream<ReadResult<Map<String, String>>> readAsStream() {
            try {
                CSVReader reader = buildCsvReader();
                List<String> headerNames = readHeaders(reader);

                long[] streamRowCount = {0};
                Spliterator<ReadResult<Map<String, String>>> spliterator = new Spliterators.AbstractSpliterator<>(
                        Long.MAX_VALUE, Spliterator.ORDERED | Spliterator.NONNULL) {
                    @Override
                    public boolean tryAdvance(Consumer<? super ReadResult<Map<String, String>>> action) {
                        try {
                            String[] line = reader.readNext();
                            if (line == null) {
                                closeQuietly(reader);
                                close();
                                return false;
                            }
                            Map<String, String> map = new LinkedHashMap<>();
                            for (int i = 0; i < headerNames.size() && i < line.length; i++) {
                                map.put(headerNames.get(i), line[i]);
                            }
                            action.accept(new ReadResult<>(map, true, null));
                            streamRowCount[0]++;
                            if (progressCallback != null && progressInterval > 0
                                    && streamRowCount[0] % progressInterval == 0) {
                                progressCallback.onProgress(streamRowCount[0], null);
                            }
                            return true;
                        } catch (Exception e) {
                            closeQuietly(reader);
                            close();
                            throw new CsvReadException("Failed to read CSV row", e);
                        }
                    }
                };

                return StreamSupport.stream(spliterator, false)
                        .onClose(() -> {
                            closeQuietly(reader);
                            close();
                        });
            } catch (CsvReadException e) {
                close();
                throw e;
            } catch (Exception e) {
                close();
                throw new CsvReadException("Failed to initialize CSV reading", e);
            }
        }

        private List<String> readHeaders(CSVReader reader) throws Exception {
            for (int i = 0; i < headerRowIndex; i++) {
                if (reader.readNext() == null) {
                    throw new CsvReadException("CSV file has insufficient rows for headerRowIndex=" + headerRowIndex);
                }
            }
            String[] headerLine = reader.readNext();
            if (headerLine == null) {
                throw new CsvReadException("CSV file is empty or missing header row");
            }
            if (headerLine.length > 0 && headerLine[0] != null && headerLine[0].startsWith("\uFEFF")) {
                headerLine[0] = headerLine[0].substring(1);
                if (headerLine[0].isEmpty()) {
                    throw new CsvReadException("First header column is empty after BOM removal");
                }
            }
            List<String> headers = new ArrayList<>();
            Collections.addAll(headers, headerLine);
            return headers;
        }

        private CSVReader buildCsvReader() throws Exception {
            CSVParser csvParser = new CSVParserBuilder().withSeparator(delimiter).build();
            return new CSVReaderBuilder(new InputStreamReader(Files.newInputStream(getTempFile()), charset))
                    .withCSVParser(csvParser).build();
        }

        private void closeQuietly(CSVReader reader) {
            try {
                reader.close();
            } catch (Exception e) {
                if (log.isDebugEnabled()) {
                    log.debug("Failed to close CSVReader", e);
                }
            }
        }
    }
}
