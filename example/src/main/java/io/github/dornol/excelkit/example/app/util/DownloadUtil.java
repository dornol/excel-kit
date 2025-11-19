package io.github.dornol.excelkit.example.app.util;

import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;

public final class DownloadUtil {

    private DownloadUtil() {
        /* empty */
    }

    public static ResponseEntity.BodyBuilder builder(String filename, DownloadFileType type) {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, type.getContentDisposition(filename))
                .header(HttpHeaders.CONTENT_TYPE, type.getContentType())
                .header(HttpHeaders.CACHE_CONTROL, "max-age=10");
    }

}
