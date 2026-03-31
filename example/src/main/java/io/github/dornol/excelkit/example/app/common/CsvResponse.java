package io.github.dornol.excelkit.example.app.common;

import io.github.dornol.excelkit.csv.CsvHandler;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

/**
 * Convenience helper for returning a {@link CsvHandler} as a Spring MVC
 * {@code ResponseEntity<StreamingResponseBody>} file download.
 *
 * <pre>{@code
 * @GetMapping("/download-csv")
 * public ResponseEntity<StreamingResponseBody> download() {
 *     CsvHandler handler = csvWriter.write(dataStream);
 *     return CsvResponse.of(handler, "report.csv");
 * }
 * }</pre>
 */
public final class CsvResponse {

    private CsvResponse() {}

    /**
     * Returns a streaming CSV file download response.
     *
     * @param handler  the CsvHandler produced by {@code CsvWriter.write()}
     * @param filename the download filename (e.g. "report.csv"); the {@code .csv}
     *                 extension is appended automatically if missing
     * @return a ResponseEntity ready to be returned from a controller method
     */
    public static ResponseEntity<StreamingResponseBody> of(CsvHandler handler, String filename) {
        String name = filename.endsWith(".csv") ? filename : filename + ".csv";
        return DownloadUtil.builder(name, DownloadFileType.CSV)
                .body(handler::consumeOutputStream);
    }
}
