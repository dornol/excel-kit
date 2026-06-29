package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.excel.ExcelHandler;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.time.Duration;

/**
 * Spring MVC response helpers for streaming Excel and CSV downloads.
 */
public final class ExcelKitResponse {
    private static final Duration DEFAULT_CACHE_MAX_AGE = Duration.ofSeconds(10);

    private ExcelKitResponse() {
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

    public static ResponseEntity.BodyBuilder builder(String filename, DownloadFileType type) {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, type.contentDisposition(filename))
                .header(HttpHeaders.CONTENT_TYPE, type.getContentType())
                .header(HttpHeaders.CACHE_CONTROL, "max-age=" + DEFAULT_CACHE_MAX_AGE.toSeconds());
    }
}
