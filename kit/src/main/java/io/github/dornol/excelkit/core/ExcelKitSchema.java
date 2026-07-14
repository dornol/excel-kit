package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelColumn;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriteErrorPolicy;
import io.github.dornol.excelkit.excel.ExcelWriter;
import jakarta.validation.Validator;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * Unified schema that defines read/write column mappings for a single entity type.
 * <p>
 * Define columns once and use them for both Excel/CSV reading and writing:
 * <pre>{@code
 * ExcelKitSchema<Person> schema = ExcelKitSchema.<Person>builder()
 *     .column("Name", Person::getName, (p, cell) -> p.setName(cell.asString()))
 *     .column("Age", Person::getAge, (p, cell) -> p.setAge(cell.asInt()),
 *             c -> c.type(ExcelDataType.INTEGER))
 *     .build();
 *
 * // Write Excel (with column type/format applied)
 * ExcelHandler handler = schema.excelWriter().write(dataStream);
 *
 * // Read Excel (matched by header name, column order doesn't matter)
 * schema.excelReader(Person::new, validator).read(inputStream, consumer);
 * }</pre>
 *
 * @param <T> The row data type
 * @author dhkim
 */
public class ExcelKitSchema<T> {

    private final List<SchemaColumn<T>> columns;
    private final boolean strictHeaders;
    private final DuplicateHeaderPolicy duplicateHeaderPolicy;
    private final @Nullable CellConversionConfig cellConversionConfig;
    private final @Nullable ExcelWriteErrorPolicy writeErrorPolicy;
    private final long maxRows;
    private final boolean skipBlankRows;
    private final int stopAtBlankRows;

    private ExcelKitSchema(List<SchemaColumn<T>> columns, Builder<T> builder) {
        this.columns = Collections.unmodifiableList(columns);
        this.strictHeaders = builder.strictHeaders;
        this.duplicateHeaderPolicy = builder.duplicateHeaderPolicy;
        this.cellConversionConfig = builder.cellConversionConfig;
        this.writeErrorPolicy = builder.writeErrorPolicy;
        this.maxRows = builder.maxRows;
        this.skipBlankRows = builder.skipBlankRows;
        this.stopAtBlankRows = builder.stopAtBlankRows;
    }

    /**
     * Creates a new schema builder.
     *
     * @param <T> The row data type
     * @return A new builder instance
     */
    public static <T> Builder<T> builder() {
        return new Builder<>();
    }

    /**
     * Creates a new {@link ExcelWriter} pre-configured with this schema's columns.
     * <p>
     * If columns have write configurers (type, format, etc.), they are applied automatically.
     * Additional options (autoFilter, freezePane, etc.) and extra columns can be chained.
     *
     * @return A configured ExcelWriter instance
     */
    public ExcelWriter<T> excelWriter() {
        return excelWriter(opts -> {});
    }

    /**
     * Creates a new {@link ExcelWriter} with initialization options and this schema's columns.
     *
     * @since 0.19.0
     */
    public ExcelWriter<T> excelWriter(Consumer<ExcelWriter.InitOptions> configurer) {
        ExcelWriter<T> writer = ExcelWriter.<T>create(configurer);
        if (writeErrorPolicy != null) {
            writer.writeErrorPolicy(writeErrorPolicy);
        }
        for (SchemaColumn<T> col : columns) {
            if (col.writeConfigurer() != null) {
                writer.column(col.name(), col.writeFunction(), col.writeConfigurer());
            } else {
                writer.column(col.name(), col.writeFunction());
            }
        }
        return writer;
    }

    /**
     * Creates a new {@link CsvWriter} pre-configured with this schema's columns.
     * Additional options (delimiter, charset, etc.) and extra columns can be chained.
     *
     * @return A configured CsvWriter instance
     */
    public CsvWriter<T> csvWriter() {
        CsvWriter<T> writer = CsvWriter.create();
        for (SchemaColumn<T> col : columns) {
            writer.column(col.name(), col.writeFunction());
        }
        return writer;
    }

    /**
     * Creates a new {@link ExcelReader} pre-configured with this schema's columns (setter mode).
     * <p>
     * Columns are matched by header name (not positional index), so the column order
     * in the Excel file does not need to match the schema definition order.
     * Additional options (sheetIndex, headerRowIndex, etc.) can be chained.
     *
     * @param supplier  A supplier to create new instances of {@code T} for each row
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured ExcelReader instance
     */
    public ExcelReader<T> excelReader(Supplier<T> supplier, @Nullable Validator validator) {
        ExcelReader<T> reader = new ExcelReader<>(supplier, validator);
        for (SchemaColumn<T> col : columns) {
            reader.column(col.readHeaderNames(), col.readSetter());
            if (col.required()) {
                reader.required();
            }
        }
        applyReaderDefaults(reader);
        return reader;
    }

    /**
     * Creates a schema-backed Excel reader without Bean Validation.
     *
     * @since 0.19.0
     */
    public ExcelReader<T> excelReader(Supplier<T> supplier) {
        return excelReader(supplier, null);
    }

    /**
     * Creates a new {@link ExcelReader} in mapping mode for immutable object construction.
     * <p>
     * The mapping function receives a {@link RowData} and creates the target object in a single step.
     * Column definitions from this schema are not used for reading in this mode.
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured ExcelReader instance in mapping mode
     */
    public ExcelReader<T> excelReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        ExcelReader<T> reader = ExcelReader.mapping(rowMapper, validator);
        applyReaderDefaults(reader);
        return reader;
    }

    /**
     * Creates a new {@link CsvReader} pre-configured with this schema's columns (setter mode).
     * <p>
     * Columns are matched by header name (not positional index), so the column order
     * in the CSV file does not need to match the schema definition order.
     * Additional options (delimiter, charset, etc.) can be chained.
     *
     * @param supplier  A supplier to create new instances of {@code T} for each row
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured CsvReader instance
     */
    public CsvReader<T> csvReader(Supplier<T> supplier, @Nullable Validator validator) {
        CsvReader<T> reader = new CsvReader<>(supplier, validator);
        for (SchemaColumn<T> col : columns) {
            reader.column(col.readHeaderNames(), col.readSetter());
            if (col.required()) {
                reader.required();
            }
        }
        applyReaderDefaults(reader);
        return reader;
    }

    /**
     * Creates a schema-backed CSV reader without Bean Validation.
     *
     * @since 0.19.0
     */
    public CsvReader<T> csvReader(Supplier<T> supplier) {
        return csvReader(supplier, null);
    }

    /**
     * Creates a new {@link CsvReader} in mapping mode for immutable object construction.
     *
     * @param rowMapper A function that creates an instance of {@code T} from a {@link RowData}
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured CsvReader instance in mapping mode
     */
    public CsvReader<T> csvReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        CsvReader<T> reader = CsvReader.mapping(rowMapper, validator);
        applyReaderDefaults(reader);
        return reader;
    }

    private <R extends AbstractReader<T, R>> void applyReaderDefaults(R reader) {
        reader.duplicateHeaderPolicy(duplicateHeaderPolicy);
        if (strictHeaders) {
            reader.strictHeaders();
        }
        if (cellConversionConfig != null) {
            reader.cellConversion(cellConversionConfig);
        }
        if (maxRows >= 0) {
            reader.maxRows(maxRows);
        }
        if (skipBlankRows) {
            reader.skipBlankRows();
        }
        if (stopAtBlankRows > 0) {
            reader.stopAtBlankRows(stopAtBlankRows);
        }
    }

    /**
     * Returns the unmodifiable list of schema columns.
     *
     * @return Schema columns
     */
    public List<SchemaColumn<T>> getColumns() {
        return columns;
    }

    /**
     * Represents a single column definition in the schema.
     *
     * @param name             Column header name
     * @param writeFunction    Function to extract the cell value from a row object
     * @param readSetter       BiConsumer to set the cell value into a row object
     * @param writeConfigurer  Optional consumer to configure Excel column properties (type, format, etc.)
     * @param readHeaderNames  Header names accepted when reading, in priority order
     * @param required         Whether blank/empty cells should fail row validation
     * @param <T>              The row data type
     */
    public record SchemaColumn<T>(
            String name,
            Function<T, @Nullable Object> writeFunction,
            BiConsumer<T, CellData> readSetter,
            @Nullable Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer,
            List<String> readHeaderNames,
            boolean required
    ) {
        /**
         * Creates a schema column.
         */
        public SchemaColumn {
            java.util.Objects.requireNonNull(name, "name cannot be null");
            java.util.Objects.requireNonNull(writeFunction, "writeFunction cannot be null");
            java.util.Objects.requireNonNull(readSetter, "readSetter cannot be null");
            readHeaderNames = normalizeReadHeaderNames(name, readHeaderNames);
        }

        /**
         * Creates a schema column without write configuration.
         *
         * @param name the column header name
         * @param writeFunction function to extract the cell value
         * @param readSetter consumer to set the cell value
         */
        public SchemaColumn(String name, Function<T, @Nullable Object> writeFunction, BiConsumer<T, CellData> readSetter) {
            this(name, writeFunction, readSetter, null, List.of(name), false);
        }

        /**
         * Creates a schema column with Excel write configuration.
         *
         * @param name the column header name
         * @param writeFunction function to extract the cell value
         * @param readSetter consumer to set the cell value
         * @param writeConfigurer optional Excel column write configuration
         */
        public SchemaColumn(String name, Function<T, @Nullable Object> writeFunction,
                            BiConsumer<T, CellData> readSetter,
                            @Nullable Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer) {
            this(name, writeFunction, readSetter, writeConfigurer, List.of(name), false);
        }

        private static List<String> normalizeReadHeaderNames(String name, List<String> aliases) {
            java.util.Objects.requireNonNull(aliases, "readHeaderNames cannot be null");
            List<String> normalized = new ArrayList<>();
            normalized.add(name);
            for (String alias : aliases) {
                if (alias == null) {
                    throw new IllegalArgumentException("readHeaderNames cannot contain null");
                }
                if (!normalized.contains(alias)) {
                    normalized.add(alias);
                }
            }
            return List.copyOf(normalized);
        }
    }

    /**
     * Builder for constructing {@link ExcelKitSchema} instances.
     *
     * @param <T> The row data type
     */
    public static class Builder<T> {
        private final List<SchemaColumn<T>> columns = new ArrayList<>();
        private boolean strictHeaders;
        private DuplicateHeaderPolicy duplicateHeaderPolicy = DuplicateHeaderPolicy.FIRST;
        private @Nullable CellConversionConfig cellConversionConfig;
        private @Nullable ExcelWriteErrorPolicy writeErrorPolicy;
        private long maxRows = -1;
        private boolean skipBlankRows;
        private int stopAtBlankRows;

        private Builder() {}

        /**
         * Adds a column definition to the schema.
         *
         * @param name          Column header name
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter    BiConsumer to set the cell value into a row object for reading
         * @return This builder for chaining
         */
        public Builder<T> column(String name, Function<T, @Nullable Object> writeFunction, BiConsumer<T, CellData> readSetter) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter));
            return this;
        }

        /**
         * Adds a column definition with read header aliases.
         *
         * @param name          Column header name used for writing
         * @param readAliases   Additional header names accepted when reading
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter    BiConsumer to set the cell value into a row object for reading
         * @return This builder for chaining
         */
        public Builder<T> column(String name, List<String> readAliases,
                                 Function<T, @Nullable Object> writeFunction,
                                 BiConsumer<T, CellData> readSetter) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, null, readAliases, false));
            return this;
        }

        /**
         * Adds a required column definition.
         *
         * @param name Column header name
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter BiConsumer to set the cell value into a row object for reading
         * @return This builder for chaining
         */
        public Builder<T> requiredColumn(String name, Function<T, @Nullable Object> writeFunction,
                                         BiConsumer<T, CellData> readSetter) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, null, List.of(name), true));
            return this;
        }

        /**
         * Adds a required column definition with read header aliases.
         *
         * @param name Column header name used for writing
         * @param readAliases Additional header names accepted when reading
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter BiConsumer to set the cell value into a row object for reading
         * @return This builder for chaining
         */
        public Builder<T> requiredColumn(String name, List<String> readAliases,
                                         Function<T, @Nullable Object> writeFunction,
                                         BiConsumer<T, CellData> readSetter) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, null, readAliases, true));
            return this;
        }

        /**
         * Adds a column definition with Excel write configuration.
         * <p>
         * The configurer receives an {@link ExcelColumn.ExcelColumnBuilder} to set column properties
         * such as type, format, alignment, width, etc. Only call configuration methods
         * (type, format, alignment, backgroundColor, bold, fontSize, width, minWidth, maxWidth, dropdown)
         * inside the configurer.
         *
         * <pre>{@code
         * .column("Price", Book::getPrice, (b, cell) -> b.setPrice(cell.asInt()),
         *         c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
         * }</pre>
         *
         * @param name            Column header name
         * @param writeFunction   Function to extract the cell value from a row object for writing
         * @param readSetter      BiConsumer to set the cell value into a row object for reading
         * @param writeConfigurer Consumer to configure Excel column properties
         * @return This builder for chaining
         */
        public Builder<T> column(String name, Function<T, @Nullable Object> writeFunction, BiConsumer<T, CellData> readSetter,
                                  Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, writeConfigurer));
            return this;
        }

        /**
         * Adds a column definition with read header aliases and Excel write configuration.
         *
         * @param name Column header name used for writing
         * @param readAliases Additional header names accepted when reading
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter BiConsumer to set the cell value into a row object for reading
         * @param writeConfigurer Consumer to configure Excel column properties
         * @return This builder for chaining
         */
        public Builder<T> column(String name, List<String> readAliases,
                                 Function<T, @Nullable Object> writeFunction,
                                 BiConsumer<T, CellData> readSetter,
                                 Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, writeConfigurer, readAliases, false));
            return this;
        }

        /**
         * Adds a required column definition with Excel write configuration.
         *
         * @param name Column header name
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter BiConsumer to set the cell value into a row object for reading
         * @param writeConfigurer Consumer to configure Excel column properties
         * @return This builder for chaining
         */
        public Builder<T> requiredColumn(String name, Function<T, @Nullable Object> writeFunction,
                                         BiConsumer<T, CellData> readSetter,
                                         Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, writeConfigurer, List.of(name), true));
            return this;
        }

        /**
         * Adds a required column definition with read header aliases and Excel write configuration.
         *
         * @param name Column header name used for writing
         * @param readAliases Additional header names accepted when reading
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter BiConsumer to set the cell value into a row object for reading
         * @param writeConfigurer Consumer to configure Excel column properties
         * @return This builder for chaining
         */
        public Builder<T> requiredColumn(String name, List<String> readAliases,
                                         Function<T, @Nullable Object> writeFunction,
                                         BiConsumer<T, CellData> readSetter,
                                         Consumer<ExcelColumn.ExcelColumnBuilder<T>> writeConfigurer) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter, writeConfigurer, readAliases, true));
            return this;
        }

        /**
         * Builds the schema.
         *
         * @return A new ExcelKitSchema instance
         * @throws IllegalArgumentException if no columns are defined
         */
        public ExcelKitSchema<T> build() {
            if (columns.isEmpty()) {
                throw new IllegalArgumentException("At least one column must be defined");
            }
            return new ExcelKitSchema<>(new ArrayList<>(columns), this);
        }

        public Builder<T> strictHeaders() {
            this.strictHeaders = true;
            return this;
        }

        public Builder<T> duplicateHeaderPolicy(DuplicateHeaderPolicy policy) {
            this.duplicateHeaderPolicy = java.util.Objects.requireNonNull(policy, "policy cannot be null");
            return this;
        }

        public Builder<T> cellConversion(CellConversionConfig config) {
            this.cellConversionConfig = java.util.Objects.requireNonNull(config, "config cannot be null");
            return this;
        }

        public Builder<T> writeErrorPolicy(ExcelWriteErrorPolicy policy) {
            this.writeErrorPolicy = java.util.Objects.requireNonNull(policy, "policy cannot be null");
            return this;
        }

        public Builder<T> maxRows(long maxRows) {
            if (maxRows < 0) {
                throw new IllegalArgumentException("maxRows must be non-negative");
            }
            this.maxRows = maxRows;
            return this;
        }

        public Builder<T> skipBlankRows() {
            this.skipBlankRows = true;
            return this;
        }

        public Builder<T> stopAtBlankRows(int count) {
            if (count < 0) {
                throw new IllegalArgumentException("count must be non-negative");
            }
            this.stopAtBlankRows = count;
            return this;
        }
    }
}
