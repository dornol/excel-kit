package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.excel.ExcelHandler;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.time.Duration;

/**
 * Spring MVC response helpers for streaming Excel and CSV downloads.
 *
 * @since 0.18.3
 */
public final class ExcelKitResponse {
    private static final Duration DEFAULT_CACHE_MAX_AGE = Duration.ofSeconds(10);

    private ExcelKitResponse() {
    }

    /**
     * Streams the handler as an Excel download.
     *
     * @param handler  the handler to stream
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> excel(ExcelHandler handler, String filename) {
        return excel(filename).body(handler::writeTo);
    }

    /**
     * Streams an Excel download encrypted with {@code password}.
     * <p>
     * A null or blank password is rejected here, before any response is built, instead of failing
     * later while the body streams. The handler is {@linkplain ExcelHandler#discard() discarded}, so
     * its workbook is closed.
     *
     * @param handler  the handler to stream
     * @param filename the download filename
     * @param password the password to encrypt the file with
     * @return the streaming response
     * @throws IllegalArgumentException if the password is null or blank
     */
    public static ResponseEntity<StreamingResponseBody> excel(
            ExcelHandler handler, String filename, String password) {
        if (password == null || password.isBlank()) {
            handler.discard();
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        return excel(filename).body(out -> handler.writeTo(out, password));
    }

    /**
     * Streams the handler as a CSV download.
     *
     * @param handler  the handler to stream
     * @param filename the download filename
     * @return the streaming response
     */
    public static ResponseEntity<StreamingResponseBody> csv(CsvHandler handler, String filename) {
        return csv(filename).body(handler::writeTo);
    }

    /**
     * Creates an Excel response builder with download headers and the default cache max-age of 10
     * seconds.
     *
     * @param filename the download filename
     * @return a response builder
     */
    public static ResponseEntity.BodyBuilder excel(String filename) {
        return builder(filename, DownloadFileType.EXCEL);
    }

    /**
     * Creates an Excel response builder with a custom cache max-age.
     *
     * @param filename the download filename
     * @param cacheMaxAge the cache duration
     * @return a response builder with download and cache headers
     */
    public static ResponseEntity.BodyBuilder excel(String filename, Duration cacheMaxAge) {
        return builder(filename, DownloadFileType.EXCEL, cacheMaxAge);
    }

    /**
     * Creates a CSV response builder with download headers and the default cache max-age of 10
     * seconds.
     *
     * @param filename the download filename
     * @return a response builder
     */
    public static ResponseEntity.BodyBuilder csv(String filename) {
        return builder(filename, DownloadFileType.CSV);
    }

    /**
     * Creates a CSV response builder with a custom cache max-age.
     *
     * @param filename the download filename
     * @param cacheMaxAge the cache duration
     * @return a response builder with download and cache headers
     */
    public static ResponseEntity.BodyBuilder csv(String filename, Duration cacheMaxAge) {
        return builder(filename, DownloadFileType.CSV, cacheMaxAge);
    }

    /**
     * Creates a response builder with download headers and the default cache max-age of 10 seconds.
     *
     * @param filename the download filename
     * @param type     the file type
     * @return a response builder
     */
    public static ResponseEntity.BodyBuilder builder(String filename, DownloadFileType type) {
        return builder(filename, type, DEFAULT_CACHE_MAX_AGE);
    }

    /**
     * Creates a response builder with download headers and a private cache of the given max-age.
     *
     * @param filename     the download filename
     * @param type         the file type
     * @param cacheMaxAge  how long a private cache may keep the download
     * @return a response builder
     * @throws IllegalArgumentException if {@code cacheMaxAge} is null or negative
     */
    public static ResponseEntity.BodyBuilder builder(
            String filename, DownloadFileType type, Duration cacheMaxAge) {
        if (cacheMaxAge == null || cacheMaxAge.isNegative()) {
            throw new IllegalArgumentException("cacheMaxAge must not be null or negative");
        }
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, type.contentDisposition(filename))
                .header(HttpHeaders.CONTENT_TYPE, type.getContentType())
                // A download is for the person who asked for it: "private" keeps shared caches (a proxy, a CDN)
                // from storing and handing it to somebody else within the max-age.
                .header(HttpHeaders.CACHE_CONTROL, "private, max-age=" + cacheMaxAge.toSeconds());
    }
}
