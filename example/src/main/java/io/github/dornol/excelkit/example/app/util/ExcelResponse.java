package io.github.dornol.excelkit.example.app.util;

import io.github.dornol.excelkit.excel.ExcelHandler;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

/**
 * Convenience helper for returning an {@link ExcelHandler} as a Spring MVC
 * {@code ResponseEntity<StreamingResponseBody>} file download.
 *
 * <pre>{@code
 * @GetMapping("/download")
 * public ResponseEntity<StreamingResponseBody> download() {
 *     ExcelHandler handler = writer.write(dataStream);
 *     return ExcelResponse.of(handler, "report.xlsx");
 * }
 * }</pre>
 */
public final class ExcelResponse {

    private ExcelResponse() {}

    /**
     * Returns a streaming Excel file download response.
     *
     * @param handler  the ExcelHandler produced by {@code ExcelWriter.write()}
     * @param filename the download filename (e.g. "report.xlsx"); the {@code .xlsx}
     *                 extension is appended automatically if missing
     * @return a ResponseEntity ready to be returned from a controller method
     */
    public static ResponseEntity<StreamingResponseBody> of(ExcelHandler handler, String filename) {
        String name = ensureExtension(filename, ".xlsx");
        return DownloadUtil.builder(name, DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

    /**
     * Returns a streaming, password-encrypted Excel file download response.
     *
     * @param handler  the ExcelHandler produced by {@code ExcelWriter.write()}
     * @param filename the download filename
     * @param password the password to encrypt the Excel file with
     * @return a ResponseEntity ready to be returned from a controller method
     */
    public static ResponseEntity<StreamingResponseBody> of(ExcelHandler handler, String filename, String password) {
        String name = ensureExtension(filename, ".xlsx");
        return DownloadUtil.builder(name, DownloadFileType.EXCEL)
                .body(out -> handler.consumeOutputStreamWithPassword(out, password));
    }

    private static String ensureExtension(String filename, String ext) {
        return filename.endsWith(ext) ? filename : filename + ext;
    }
}
