package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.Collection;
import java.util.List;
import java.util.stream.Stream;

/**
 * Spring MVC response helpers for downloadable read-error reports.
 */
public final class ExcelKitErrorResponse {

    private ExcelKitErrorResponse() {
    }

    public static ResponseEntity<StreamingResponseBody> csv(
            UploadResult<?> result, String filename) {
        return csv(result.errors(), filename);
    }

    public static ResponseEntity<StreamingResponseBody> csv(
            Collection<UploadError> errors, String filename) {
        var handler = CsvWriter.<ErrorReportRow>create()
                .column("rowNum", ErrorReportRow::rowNum)
                .column("fileRowNum", ErrorReportRow::fileRowNum)
                .column("columnIndex", ErrorReportRow::columnIndex)
                .column("headerName", ErrorReportRow::headerName)
                .column("cellValue", ErrorReportRow::cellValue)
                .column("message", ErrorReportRow::message)
                .write(reportRows(errors));

        return ExcelKitResponse.csv(handler, filename);
    }

    public static ResponseEntity<StreamingResponseBody> excel(
            UploadResult<?> result, String filename) {
        return excel(result.errors(), filename);
    }

    public static ResponseEntity<StreamingResponseBody> excel(
            Collection<UploadError> errors, String filename) {
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
                .write(reportRows(errors));

        return ExcelKitResponse.excel(handler, filename);
    }

    public static Stream<ErrorReportRow> reportRows(Collection<UploadError> errors) {
        return errors.stream().flatMap(error -> {
            if (error.cellErrors().isEmpty()) {
                List<String> messages = error.messages().isEmpty() ? List.of("Unknown read error") : error.messages();
                return messages.stream().map(message -> new ErrorReportRow(
                        error.rowNum(), error.fileRowNum(), null, null, null, message));
            }
            return error.cellErrors().stream().map(cell -> new ErrorReportRow(
                    error.rowNum(),
                    error.fileRowNum(),
                    cell.columnIndex(),
                    cell.headerName(),
                    cell.cellValue(),
                    cell.message()
            ));
        });
    }
}
