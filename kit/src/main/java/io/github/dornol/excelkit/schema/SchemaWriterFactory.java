package io.github.dornol.excelkit.schema;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ColumnConfig;
import io.github.dornol.excelkit.excel.ExcelWriter;

import java.util.List;
import java.util.function.Consumer;

/** Builds writers from the write side of a unified schema. */
final class SchemaWriterFactory {
    private SchemaWriterFactory() {
    }

    static <T> ExcelWriter<T> excel(List<ExcelKitSchema.SchemaColumn<T>> columns,
                                    WriteErrorPolicy writeErrorPolicy,
                                    Consumer<ExcelWriter.InitOptions> configurer) {
        ExcelWriter<T> writer = ExcelWriter.create(configurer);
        if (writeErrorPolicy != null) writer.writeErrorPolicy(writeErrorPolicy);
        for (ExcelKitSchema.SchemaColumn<T> column : columns) {
            Consumer<ColumnConfig<T>> columnConfigurer = column.writeConfigurer();
            if (column.required()) {
                // A required column shows the mark on its header; the column's own configuration may still turn it off.
                Consumer<ColumnConfig<T>> own = columnConfigurer;
                columnConfigurer = builder -> {
                    builder.requiredMark();
                    if (own != null) own.accept(builder);
                };
            }
            if (columnConfigurer != null) {
                writer.column(column.name(), column.writeFunction(), columnConfigurer);
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
