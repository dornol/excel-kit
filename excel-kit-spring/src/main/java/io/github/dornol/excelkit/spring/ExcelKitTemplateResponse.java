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
}
