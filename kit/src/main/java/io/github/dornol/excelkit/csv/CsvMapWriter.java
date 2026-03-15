package io.github.dornol.excelkit.csv;

import java.util.Map;
import java.util.stream.Stream;

/**
 * Convenience writer for generating CSV files from {@code Map<String, Object>} data.
 *
 * <pre>{@code
 * CsvMapWriter writer = new CsvMapWriter("Name", "Age", "Email");
 * writer.write(Stream.of(
 *     Map.of("Name", "Alice", "Age", 30, "Email", "alice@example.com")
 * )).consumeOutputStream(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class CsvMapWriter {

    private final CsvWriter<Map<String, Object>> writer;

    /**
     * Creates a CSV map writer with the given column names.
     *
     * @param columnNames column names (used as headers and map keys)
     */
    public CsvMapWriter(String... columnNames) {
        this(new CsvWriter<>(), columnNames);
    }

    /**
     * Creates a CSV map writer with a pre-configured CsvWriter.
     *
     * @param writer      the base CsvWriter
     * @param columnNames column names
     */
    public CsvMapWriter(CsvWriter<Map<String, Object>> writer, String... columnNames) {
        this.writer = writer;
        for (String name : columnNames) {
            this.writer.column(name, map -> map.get(name));
        }
    }

    /**
     * Returns the underlying CsvWriter for additional configuration.
     */
    public CsvWriter<Map<String, Object>> writer() {
        return writer;
    }

    /**
     * Writes the data stream and returns a CsvHandler.
     */
    public CsvHandler write(Stream<Map<String, Object>> stream) {
        return writer.write(stream);
    }
}
