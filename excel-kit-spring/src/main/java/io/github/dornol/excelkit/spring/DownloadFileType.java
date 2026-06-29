package io.github.dornol.excelkit.spring;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

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
        String downloadFilename = filename.endsWith("." + extension)
                ? filename
                : filename + "." + extension;
        String fallback = downloadFilename.replaceAll("[^\\x20-\\x7E]", "_");
        String encoded = URLEncoder.encode(downloadFilename, StandardCharsets.UTF_8)
                .replace("+", "%20")
                .replace("%2B", "+");

        return String.format(Locale.ROOT,
                "attachment; filename=\"%s\"; filename*=UTF-8''%s", fallback, encoded);
    }
}
