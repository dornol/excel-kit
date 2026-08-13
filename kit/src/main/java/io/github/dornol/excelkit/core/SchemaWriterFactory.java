package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelWriteErrorPolicy;
import io.github.dornol.excelkit.excel.ExcelWriter;

import java.util.List;
import java.util.function.Consumer;

/** Builds writers from the write side of a unified schema. */
final class SchemaWriterFactory {
    private SchemaWriterFactory() {
    }

    static <T> ExcelWriter<T> excel(List<ExcelKitSchema.SchemaColumn<T>> columns,
                                    ExcelWriteErrorPolicy writeErrorPolicy,
                                    Consumer<ExcelWriter.InitOptions> configurer) {
        ExcelWriter<T> writer = ExcelWriter.create(configurer);
        if (writeErrorPolicy != null) writer.writeErrorPolicy(writeErrorPolicy);
        for (ExcelKitSchema.SchemaColumn<T> column : columns) {
            if (column.writeConfigurer() != null) {
                writer.column(column.name(), column.writeFunction(), column.writeConfigurer());
            } else {
                writer.column(column.name(), column.writeFunction());
            }
        }
        return writer;
    }

    static <T> CsvWriter<T> csv(List<ExcelKitSchema.SchemaColumn<T>> columns) {
        CsvWriter<T> writer = CsvWriter.create();
        for (ExcelKitSchema.SchemaColumn<T> column : columns) {
            writer.column(column.name(), column.writeFunction());
        }
        return writer;
    }
}
