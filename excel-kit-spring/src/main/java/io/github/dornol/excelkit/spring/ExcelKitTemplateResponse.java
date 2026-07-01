package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitSchema;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.Collection;
import java.util.stream.Stream;

/**
 * Spring MVC response helpers for schema-based empty upload templates.
 */
public final class ExcelKitTemplateResponse {

    private ExcelKitTemplateResponse() {
    }

    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename) {
        return excel(schema, filename, Stream.empty());
    }

    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return excel(schema, filename, sampleRows.stream());
    }

    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename, Stream<T> sampleRows) {
        var handler = schema.excelWriter()
                .sheetName("Template")
                .autoFilter(true)
                .freezeRows(1)
                .write(sampleRows);

        return ExcelKitResponse.excel(handler, filename);
    }

    public static <T> ResponseEntity<StreamingResponseBody> excelWithGuidance(
            ExcelKitSchema<T> schema, String filename) {
        return excelWithGuidance(schema, filename, Stream.empty());
    }

    public static <T> ResponseEntity<StreamingResponseBody> excelWithGuidance(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return excelWithGuidance(schema, filename, sampleRows.stream());
    }

    public static <T> ResponseEntity<StreamingResponseBody> excelWithGuidance(
            ExcelKitSchema<T> schema, String filename, Stream<T> sampleRows) {
        var handler = schema.excelWriter()
                .sheetName("Template")
                .beforeHeader(context -> {
                    var row = context.getSheet().createRow(context.getCurrentRow());
                    row.createCell(0).setCellValue("Required columns: " + requiredColumnNames(schema));
                    if (context.getColumnCount() > 1) {
                        context.mergeCells(context.getCurrentRow(), context.getCurrentRow(), 0, context.getColumnCount() - 1);
                    }
                    return context.getCurrentRow() + 1;
                })
                .autoFilter(true)
                .freezeRows(1)
                .write(sampleRows);

        return ExcelKitResponse.excel(handler, filename);
    }

    public static <T> ResponseEntity<StreamingResponseBody> csv(
            ExcelKitSchema<T> schema, String filename) {
        return csv(schema, filename, Stream.empty());
    }

    public static <T> ResponseEntity<StreamingResponseBody> csv(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return csv(schema, filename, sampleRows.stream());
    }

    public static <T> ResponseEntity<StreamingResponseBody> csv(
            ExcelKitSchema<T> schema, String filename, Stream<T> sampleRows) {
        var handler = schema.csvWriter().write(sampleRows);

        return ExcelKitResponse.csv(handler, filename);
    }

    private static <T> String requiredColumnNames(ExcelKitSchema<T> schema) {
        String joined = schema.getColumns().stream()
                .filter(ExcelKitSchema.SchemaColumn::required)
                .map(column -> {
                    if (column.readHeaderNames().size() <= 1) {
                        return column.name();
                    }
                    return column.name() + " (aliases: "
                            + String.join(", ", column.readHeaderNames().subList(1, column.readHeaderNames().size()))
                            + ")";
                })
                .collect(java.util.stream.Collectors.joining(", "));
        return joined.isBlank() ? "(none)" : joined;
    }
}
