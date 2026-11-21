package io.github.dornol.excelkit.schema;

import io.github.dornol.excelkit.core.AbstractReader;
import io.github.dornol.excelkit.core.ReaderOptions;
import io.github.dornol.excelkit.core.RowData;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.List;
import java.util.function.Function;
import java.util.function.Supplier;

/** Builds readers from the read side of a unified schema. */
final class SchemaReaderFactory {
    private SchemaReaderFactory() {
    }

    static <T> ExcelReader<T> excelSetter(List<ExcelKitSchema.SchemaColumn<T>> columns,
                                          ReaderOptions options, Supplier<T> supplier,
                                          @Nullable Validator validator) {
        ExcelReader<T> reader = new ExcelReader<>(supplier, validator);
        configureColumns(reader, columns);
        configure(reader, options, validator);
        return reader;
    }

    static <T> ExcelReader<T> excelMapping(ReaderOptions options,
                                           Function<RowData, T> mapper,
                                           @Nullable Validator validator) {
        ExcelReader<T> reader = ExcelReader.mapping(mapper, validator);
        configure(reader, options, validator);
        return reader;
    }

    static <T> CsvReader<T> csvSetter(List<ExcelKitSchema.SchemaColumn<T>> columns,
                                      ReaderOptions options, Supplier<T> supplier,
                                      @Nullable Validator validator) {
        CsvReader<T> reader = new CsvReader<>(supplier, validator);
        configureColumns(reader, columns);
        configure(reader, options, validator);
        return reader;
    }

    static <T> CsvReader<T> csvMapping(ReaderOptions options,
                                       Function<RowData, T> mapper,
                                       @Nullable Validator validator) {
        CsvReader<T> reader = CsvReader.mapping(mapper, validator);
        configure(reader, options, validator);
        return reader;
    }

    private static <T, R extends AbstractReader<T, R>> void configureColumns(
            R reader, List<ExcelKitSchema.SchemaColumn<T>> columns) {
        for (ExcelKitSchema.SchemaColumn<T> column : columns) {
            reader.column(column.readHeaderNames(), column.readSetter());
            if (column.required()) reader.required();
        }
    }

    /**
     * Applies the schema's common settings. {@code options(...)} replaces every common setting,
     * including the validator, so the validator given for this reader goes into the options first.
     */
    private static <T, R extends AbstractReader<T, R>> void configure(
            R reader, ReaderOptions options, @Nullable Validator validator) {
        reader.options(options.toBuilder().validator(validator).build());
    }
}
