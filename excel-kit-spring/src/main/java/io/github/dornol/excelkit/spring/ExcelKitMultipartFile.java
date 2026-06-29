package io.github.dornol.excelkit.spring;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;

/**
 * Small adapter for opening upload streams from Spring {@link MultipartFile}s.
 */
public final class ExcelKitMultipartFile {

    private ExcelKitMultipartFile() {
    }

    public static InputStream open(MultipartFile file) {
        Objects.requireNonNull(file, "file must not be null");
        if (file.isEmpty()) {
            throw new ExcelKitUploadException("Upload file is empty");
        }
        try {
            return file.getInputStream();
        } catch (IOException e) {
            throw new ExcelKitUploadException("Failed to open upload file: " + file.getOriginalFilename(), e);
        }
    }
}
