package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.Collection;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Spring MVC response helpers for schema-based empty upload templates.
 *
 * @since 0.18.4
 */
public final class ExcelKitTemplateResponse {

    private ExcelKitTemplateResponse() {
    }

    /**
     * Streams an Excel template: the schema's header row, with no data rows.
     *
     * @param schema   the schema that defines the columns
     * @param filename the download filename
     * @param <T>      the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename) {
        return excel(schema, filename, Stream.empty());
    }

    /**
     * Streams an Excel template: the schema's header row followed by the sample rows.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in
     * @param <T>        the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return excel(schema, filename, sampleRows.stream());
    }

    /**
     * Streams an Excel template: the schema's header row followed by the sample rows.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in; not closed here
     * @param <T>        the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> excel(
            ExcelKitSchema<T> schema, String filename, Stream<T> sampleRows) {
        var handler = schema.excelWriter()
                .sheetName("Template")
                .autoFilter(true)
                .freezeRows(1)
                .write(sampleRows);

        return ExcelKitResponse.excel(handler, filename);
    }

    /**
     * Like {@link #excel(ExcelKitSchema, String)}, with a first row that lists the required columns.
     *
     * @param schema   the schema that defines the columns
     * @param filename the download filename
     * @param <T>      the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> excelWithGuidance(
            ExcelKitSchema<T> schema, String filename) {
        return excelWithGuidance(schema, filename, Stream.empty());
    }

    /**
     * Like {@link #excel(ExcelKitSchema, String, Collection)}, with a first row that lists the
     * required columns.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in
     * @param <T>        the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> excelWithGuidance(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return excelWithGuidance(schema, filename, sampleRows.stream());
    }

    /**
     * Like {@link #excel(ExcelKitSchema, String, Stream)}, with a first row that lists the required
     * columns.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in; not closed here
     * @param <T>        the row type
     * @return the streaming response
     */
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

    /**
     * Streams a CSV template: the schema's header row, with no data rows.
     *
     * @param schema   the schema that defines the columns
     * @param filename the download filename
     * @param <T>      the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> csv(
            ExcelKitSchema<T> schema, String filename) {
        return csv(schema, filename, Stream.empty());
    }

    /**
     * Streams a CSV template: the schema's header row followed by the sample rows.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in
     * @param <T>        the row type
     * @return the streaming response
     */
    public static <T> ResponseEntity<StreamingResponseBody> csv(
            ExcelKitSchema<T> schema, String filename, Collection<T> sampleRows) {
        return csv(schema, filename, sampleRows.stream());
    }

    /**
     * Streams a CSV template: the schema's header row followed by the sample rows.
     *
     * @param schema     the schema that defines the columns
     * @param filename   the download filename
     * @param sampleRows rows that show how to fill the sheet in; not closed here
     * @param <T>        the row type
     * @return the streaming response
     */
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
                .collect(Collectors.joining(", "));
        return joined.isBlank() ? "(none)" : joined;
    }
}
