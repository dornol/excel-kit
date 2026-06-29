package io.github.dornol.excelkit.example.app.common;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.excel.ExcelHandler;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

public final class DownloadResponse {

    private DownloadResponse() {
    }

    public static ResponseEntity<StreamingResponseBody> excel(ExcelHandler handler, String filename) {
        return excel(filename).body(handler::writeTo);
    }

    public static ResponseEntity<StreamingResponseBody> excel(
            ExcelHandler handler, String filename, String password) {
        return excel(filename).body(out -> handler.writeTo(out, password));
    }

    public static ResponseEntity<StreamingResponseBody> csv(CsvHandler handler, String filename) {
        return csv(filename).body(handler::writeTo);
    }

    public static ResponseEntity.BodyBuilder excel(String filename) {
        return builder(filename, DownloadFileType.EXCEL);
    }

    public static ResponseEntity.BodyBuilder csv(String filename) {
        return builder(filename, DownloadFileType.CSV);
    }

    private static ResponseEntity.BodyBuilder builder(String filename, DownloadFileType type) {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, type.getContentDisposition(filename))
                .header(HttpHeaders.CONTENT_TYPE, type.getContentType())
                .header(HttpHeaders.CACHE_CONTROL, "max-age=10");
    }
}
