package io.github.dornol.excelkit.spring;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.Objects;

/**
 * Supported download file types for Spring MVC responses.
 */
public enum DownloadFileType {
    EXCEL("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"),
    CSV("text/csv; charset=UTF-8", "csv");

    private final String contentType;
    private final String extension;

    DownloadFileType(String contentType, String extension) {
        this.contentType = contentType;
        this.extension = extension;
    }

    public String getContentType() {
        return contentType;
    }

    public String getExtension() {
        return extension;
    }

    public String contentDisposition(String filename) {
        String downloadFilename = normalizeFilename(filename);
        String fallback = asciiFallback(downloadFilename);
        String encoded = URLEncoder.encode(downloadFilename, StandardCharsets.UTF_8)
                .replace("+", "%20")
                .replace("%2B", "+");

        return String.format(Locale.ROOT,
                "attachment; filename=\"%s\"; filename*=UTF-8''%s", fallback, encoded);
    }

    private String normalizeFilename(String filename) {
        Objects.requireNonNull(filename, "filename must not be null");
        String sanitized = filename.trim()
                .replaceAll("[\\p{Cntrl}\\\\/]+", "_")
                .replaceAll("_+", "_");
        if (sanitized.isBlank()) {
            sanitized = "download";
        }
        return sanitized.toLowerCase(Locale.ROOT).endsWith("." + extension)
                ? sanitized
                : sanitized + "." + extension;
    }

    private static String asciiFallback(String filename) {
        String fallback = filename.replaceAll("[^A-Za-z0-9._ -]", "_")
                .replaceAll("_+", "_");
        return fallback.isBlank() ? "download" : fallback;
    }
}
