package io.github.dornol.excelkit.shared;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import jakarta.validation.Validator;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * Unified schema that defines read/write column mappings for a single entity type.
 * <p>
 * Define columns once and use them for both Excel/CSV reading and writing:
 * <pre>{@code
 * ExcelKitSchema<Person> schema = ExcelKitSchema.<Person>builder()
 *     .column("이름", Person::getName, (p, cell) -> p.setName(cell.asString()))
 *     .column("나이", Person::getAge, (p, cell) -> p.setAge(cell.asInt()))
 *     .build();
 *
 * // Write Excel
 * ExcelHandler handler = schema.excelWriter().write(dataStream);
 *
 * // Read Excel
 * schema.excelReader(Person::new, validator).build(inputStream).read(consumer);
 * }</pre>
 *
 * @param <T> The row data type
 * @author dhkim
 */
public class ExcelKitSchema<T> {

    private final List<SchemaColumn<T>> columns;

    private ExcelKitSchema(List<SchemaColumn<T>> columns) {
        this.columns = Collections.unmodifiableList(columns);
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
     * Additional options (title, autoFilter, etc.) and extra columns can be chained.
     *
     * @return A configured ExcelWriter instance
     */
    public ExcelWriter<T> excelWriter() {
        ExcelWriter<T> writer = new ExcelWriter<>();
        for (SchemaColumn<T> col : columns) {
            writer.addColumn(col.name(), col.writeFunction());
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
        CsvWriter<T> writer = new CsvWriter<>();
        for (SchemaColumn<T> col : columns) {
            writer.column(col.name(), col.writeFunction());
        }
        return writer;
    }

    /**
     * Creates a new {@link ExcelReader} pre-configured with this schema's columns.
     * Additional options (sheetIndex, headerRowIndex, etc.) can be chained.
     *
     * @param supplier  A supplier to create new instances of {@code T} for each row
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured ExcelReader instance
     */
    public ExcelReader<T> excelReader(Supplier<T> supplier, Validator validator) {
        ExcelReader<T> reader = new ExcelReader<>(supplier, validator);
        for (SchemaColumn<T> col : columns) {
            reader.addColumn(col.readSetter());
        }
        return reader;
    }

    /**
     * Creates a new {@link CsvReader} pre-configured with this schema's columns.
     * Additional options (delimiter, charset, etc.) can be chained.
     *
     * @param supplier  A supplier to create new instances of {@code T} for each row
     * @param validator Optional Bean Validation validator (nullable)
     * @return A configured CsvReader instance
     */
    public CsvReader<T> csvReader(Supplier<T> supplier, Validator validator) {
        CsvReader<T> reader = new CsvReader<>(supplier, validator);
        for (SchemaColumn<T> col : columns) {
            reader.addColumn(col.readSetter());
        }
        return reader;
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
     * @param name          Column header name
     * @param writeFunction Function to extract the cell value from a row object
     * @param readSetter    BiConsumer to set the cell value into a row object
     * @param <T>           The row data type
     */
    public record SchemaColumn<T>(
            String name,
            Function<T, Object> writeFunction,
            BiConsumer<T, CellData> readSetter
    ) {}

    /**
     * Builder for constructing {@link ExcelKitSchema} instances.
     *
     * @param <T> The row data type
     */
    public static class Builder<T> {
        private final List<SchemaColumn<T>> columns = new ArrayList<>();

        private Builder() {}

        /**
         * Adds a column definition to the schema.
         *
         * @param name          Column header name
         * @param writeFunction Function to extract the cell value from a row object for writing
         * @param readSetter    BiConsumer to set the cell value into a row object for reading
         * @return This builder for chaining
         */
        public Builder<T> column(String name, Function<T, Object> writeFunction, BiConsumer<T, CellData> readSetter) {
            columns.add(new SchemaColumn<>(name, writeFunction, readSetter));
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
            return new ExcelKitSchema<>(new ArrayList<>(columns));
        }
    }
}
