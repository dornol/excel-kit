package io.github.dornol.excelkit.example.app.util;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

public enum DownloadFileType {
    EXCEL("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"),
    CSV("text/csv; charset=UTF-8", "csv"),
    ;
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

    public String getContentDisposition(String filename) {
        String csvFilename = filename + "." + extension;

        // ASCII fallback (따옴표 안에 ASCII 문자만)
        String fallback = csvFilename.replaceAll("[^\\x20-\\x7E]", "_");

        // RFC 5987 인코딩 (UTF-8 + percent encoding)
        String encoded;
        try {
            encoded = URLEncoder.encode(csvFilename, StandardCharsets.UTF_8)
                    .replace("+", "%20")
                    .replace("%2B", "+"); // 선택적으로 '+' 복원
        } catch (Exception _) {
            encoded = fallback;
        }

        return String.format(Locale.ROOT,
                "attachment; filename=\"%s\"; filename*=UTF-8''%s", fallback, encoded);
    }
}
