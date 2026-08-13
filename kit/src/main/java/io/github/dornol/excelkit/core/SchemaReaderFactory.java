package io.github.dornol.excelkit.core;

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
                                          SchemaReadOptions options, Supplier<T> supplier,
                                          @Nullable Validator validator) {
        ExcelReader<T> reader = new ExcelReader<>(supplier, validator);
        configureColumns(reader, columns);
        configure(reader, options);
        return reader;
    }

    static <T> ExcelReader<T> excelMapping(SchemaReadOptions options,
                                           Function<RowData, T> mapper,
                                           @Nullable Validator validator) {
        ExcelReader<T> reader = ExcelReader.mapping(mapper, validator);
        configure(reader, options);
        return reader;
    }

    static <T> CsvReader<T> csvSetter(List<ExcelKitSchema.SchemaColumn<T>> columns,
                                      SchemaReadOptions options, Supplier<T> supplier,
                                      @Nullable Validator validator) {
        CsvReader<T> reader = new CsvReader<>(supplier, validator);
        configureColumns(reader, columns);
        configure(reader, options);
        return reader;
    }

    static <T> CsvReader<T> csvMapping(SchemaReadOptions options,
                                       Function<RowData, T> mapper,
                                       @Nullable Validator validator) {
        CsvReader<T> reader = CsvReader.mapping(mapper, validator);
        configure(reader, options);
        return reader;
    }

    private static <T, R extends AbstractReader<T, R>> void configureColumns(
            R reader, List<ExcelKitSchema.SchemaColumn<T>> columns) {
        for (ExcelKitSchema.SchemaColumn<T> column : columns) {
            reader.column(column.readHeaderNames(), column.readSetter());
            if (column.required()) reader.required();
        }
    }

    private static <T, R extends AbstractReader<T, R>> void configure(
            R reader, SchemaReadOptions options) {
        reader.duplicateHeaderPolicy(options.duplicateHeaderPolicy());
        if (options.strictHeaders()) reader.strictHeaders();
        if (options.cellConversionConfig() != null) {
            reader.cellConversion(options.cellConversionConfig());
        }
        if (options.maxRows() >= 0) reader.maxRows(options.maxRows());
        if (options.skipBlankRows()) reader.skipBlankRows();
        if (options.stopAtBlankRows() > 0) reader.stopAtBlankRows(options.stopAtBlankRows());
    }

    record SchemaReadOptions(boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                             @Nullable CellConversionConfig cellConversionConfig, long maxRows,
                             boolean skipBlankRows, int stopAtBlankRows) {
    }
}
