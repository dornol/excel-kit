package io.github.dornol.excelkit.excel;

import java.util.Map;
import java.util.function.Consumer;

/**
 * Convenience writer for generating Excel files from {@code Map<String, Object>} data.
 * <p>
 * Dynamically creates columns from the specified column names, extracting values
 * from each map by key.
 *
 * <pre>{@code
 * ExcelMapWriter writer = new ExcelMapWriter("Name", "Age", "Email");
 * writer.write(Stream.of(
 *     Map.of("Name", "Alice", "Age", 30, "Email", "alice@example.com"),
 *     Map.of("Name", "Bob", "Age", 25, "Email", "bob@example.com")
 * )).consumeOutputStream(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelMapWriter {

    private final ExcelWriter<Map<String, Object>> writer;

    /**
     * Creates a map writer with the given column names and default settings.
     *
     * @param columnNames the column names (used as header labels and map keys)
     */
    public ExcelMapWriter(String... columnNames) {
        this(new ExcelWriter<>(), columnNames);
    }

    /**
     * Creates a map writer with a pre-configured ExcelWriter and column names.
     *
     * @param writer      the base ExcelWriter to use
     * @param columnNames the column names
     */
    public ExcelMapWriter(ExcelWriter<Map<String, Object>> writer, String... columnNames) {
        this.writer = writer;
        for (String name : columnNames) {
            this.writer.addColumn(name, map -> map.get(name));
        }
    }

    /**
     * Creates a map writer with column names and per-column configuration.
     *
     * @param writer      the base ExcelWriter to use
     * @param columnNames the column names
     * @param configurers per-column configurers (array length must match columnNames)
     */
    @SafeVarargs
    public ExcelMapWriter(ExcelWriter<Map<String, Object>> writer, String[] columnNames,
                          Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>... configurers) {
        this.writer = writer;
        for (int i = 0; i < columnNames.length; i++) {
            String name = columnNames[i];
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>> cfg =
                    (i < configurers.length) ? configurers[i] : null;
            this.writer.addColumn(name, map -> map.get(name), cfg);
        }
    }

    /**
     * Returns the underlying ExcelWriter for additional configuration.
     */
    public ExcelWriter<Map<String, Object>> writer() {
        return writer;
    }

    /**
     * Writes the data stream and returns an ExcelHandler.
     *
     * @param stream the data stream of maps
     * @return ExcelHandler for output
     */
    public ExcelHandler write(java.util.stream.Stream<Map<String, Object>> stream) {
        return writer.write(stream);
    }
}
