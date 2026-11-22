package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.Collection;
import java.time.Duration;
import java.util.List;
import java.util.stream.Stream;

/**
 * Spring MVC response helpers for downloadable read-error reports.
 *
 * @since 0.18.4
 */
public final class ExcelKitErrorResponse {

    private ExcelKitErrorResponse() {
    }

    /**
     * Streams the errors of an upload result as a CSV download.
     *
     * @param result   the upload result
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> csv(
            UploadResult<?> result, String filename) {
        return csv(result.errors(), filename);
    }

    /**
     * Streams the errors of an upload result as a CSV download with the given cache max-age.
     *
     * @param result      the upload result
     * @param filename    the download filename
     * @param cacheMaxAge how long a private cache may keep the download
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> csv(
            UploadResult<?> result, String filename, Duration cacheMaxAge) {
        return csv(result.errors(), filename, cacheMaxAge);
    }

    /**
     * Streams errors as a CSV download.
     *
     * @param errors   the errors to report
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> csv(
            Collection<UploadError> errors, String filename) {
        return csv(errors, filename, ExcelKitResponse.csv(filename));
    }

    /**
     * Streams errors as a CSV download with the given cache max-age.
     *
     * @param errors      the errors to report
     * @param filename    the download filename
     * @param cacheMaxAge how long a private cache may keep the download
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> csv(
            Collection<UploadError> errors, String filename, Duration cacheMaxAge) {
        return csv(errors, filename, ExcelKitResponse.csv(filename, cacheMaxAge));
    }

    private static ResponseEntity<StreamingResponseBody> csv(
            Collection<UploadError> errors, String filename,
            ResponseEntity.BodyBuilder response) {
        var handler = CsvWriter.<ErrorReportRow>create()
                .column("rowNum", ErrorReportRow::rowNum)
                .column("fileRowNum", ErrorReportRow::fileRowNum)
                .column("columnIndex", ErrorReportRow::columnIndex)
                .column("headerName", ErrorReportRow::headerName)
                .column("cellValue", ErrorReportRow::cellValue)
                .column("message", ErrorReportRow::message)
                .column("rawValues", ErrorReportRow::rawValues)
                .write(reportRows(errors));

        return response.body(handler::writeTo);
    }

    /**
     * Streams the errors of an upload result as an Excel download.
     *
     * @param result   the upload result
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> excel(
            UploadResult<?> result, String filename) {
        return excel(result.errors(), filename);
    }

    /**
     * Streams the errors of an upload result as an Excel download with the given cache max-age.
     *
     * @param result      the upload result
     * @param filename    the download filename
     * @param cacheMaxAge how long a private cache may keep the download
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> excel(
            UploadResult<?> result, String filename, Duration cacheMaxAge) {
        return excel(result.errors(), filename, cacheMaxAge);
    }

    /**
     * Streams errors as an Excel download.
     *
     * @param errors   the errors to report
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> excel(
            Collection<UploadError> errors, String filename) {
        return excel(errors, filename, ExcelKitResponse.excel(filename));
    }

    /**
     * Streams errors as an Excel download with the given cache max-age.
     *
     * @param errors      the errors to report
     * @param filename    the download filename
     * @param cacheMaxAge how long a private cache may keep the download
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> excel(
            Collection<UploadError> errors, String filename, Duration cacheMaxAge) {
        return excel(errors, filename, ExcelKitResponse.excel(filename, cacheMaxAge));
    }

    private static ResponseEntity<StreamingResponseBody> excel(
            Collection<UploadError> errors, String filename,
            ResponseEntity.BodyBuilder response) {
        var handler = ExcelWriter.<ErrorReportRow>create()
                .sheetName("Read Errors")
                .autoFilter(true)
                .freezeRows(1)
                .column("rowNum", ErrorReportRow::rowNum, c -> c.type(ExcelDataType.LONG))
                .column("fileRowNum", ErrorReportRow::fileRowNum, c -> c.type(ExcelDataType.LONG))
                .column("columnIndex", ErrorReportRow::columnIndex, c -> c.type(ExcelDataType.INTEGER))
                .column("headerName", ErrorReportRow::headerName)
                .column("cellValue", ErrorReportRow::cellValue)
                .column("message", ErrorReportRow::message)
                .column("rawValues", ErrorReportRow::rawValues)
                .write(reportRows(errors));

        return response.body(handler::writeTo);
    }

    /**
     * Flattens errors into report rows: one row per failed cell, or one row per message for an error
     * that has no cell errors.
     *
     * @param errors the errors to report
     * @return the report rows
     */
    public static Stream<ErrorReportRow> reportRows(Collection<UploadError> errors) {
        return errors.stream().flatMap(error -> {
            if (error.cellErrors().isEmpty()) {
                List<String> messages = error.messages().isEmpty() ? List.of("Unknown read error") : error.messages();
                return messages.stream().map(message -> new ErrorReportRow(
                        error.rowNum(), error.fileRowNum(), null, null, null, message, joinRawValues(error.rawValues())));
            }
            return error.cellErrors().stream().map(cell -> new ErrorReportRow(
                    error.rowNum(),
                    error.fileRowNum(),
                    cell.columnIndex(),
                    cell.headerName(),
                    cell.cellValue(),
                    cell.message(),
                    joinRawValues(error.rawValues())
            ));
        });
    }

    private static String joinRawValues(List<String> rawValues) {
        return rawValues == null || rawValues.isEmpty() ? "" : String.join(" | ", rawValues);
    }
}
