package io.github.dornol.excelkit.spring;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Small adapter for opening upload streams from Spring {@link MultipartFile}s.
 */
public final class ExcelKitMultipartFile {
    private static final Set<String> EXCEL_AND_CSV_EXTENSIONS = Set.of("xlsx", "xlsm", "csv");

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
            throw new ExcelKitUploadException("Failed to open upload file: " + safeOriginalFilename(file), e);
        }
    }

    /**
     * Requires the upload to be at or below the given size.
     *
     * @param file     upload file
     * @param maxBytes maximum allowed size in bytes
     * @return the same file for fluent validation before passing it to {@link ExcelKitUpload}
     */
    public static MultipartFile requireSizeAtMost(MultipartFile file, long maxBytes) {
        Objects.requireNonNull(file, "file must not be null");
        if (maxBytes < 0) {
            throw new IllegalArgumentException("maxBytes must not be negative");
        }
        if (file.getSize() > maxBytes) {
            throw new ExcelKitUploadException("Upload file exceeds maximum size of " + maxBytes + " bytes");
        }
        return file;
    }

    /**
     * Requires the upload filename to end with one of the given extensions.
     * Extensions may be supplied with or without a leading dot.
     *
     * @param file       upload file
     * @param extensions allowed extensions such as {@code "xlsx"} or {@code ".csv"}
     * @return the same file for fluent validation before passing it to {@link ExcelKitUpload}
     */
    public static MultipartFile requireExtension(MultipartFile file, String... extensions) {
        Objects.requireNonNull(file, "file must not be null");
        Set<String> allowed = normalizeExtensions(extensions);
        String extension = extensionOf(file.getOriginalFilename());
        if (!allowed.contains(extension)) {
            throw new ExcelKitUploadException(
                    "Unsupported upload file extension. Allowed extensions: " + String.join(", ", allowed));
        }
        return file;
    }

    /**
     * Requires an Excel or CSV upload filename extension.
     *
     * @param file upload file
     * @return the same file for fluent validation before passing it to {@link ExcelKitUpload}
     */
    public static MultipartFile requireExcelOrCsv(MultipartFile file) {
        return requireExtension(file, EXCEL_AND_CSV_EXTENSIONS.toArray(String[]::new));
    }

    private static Set<String> normalizeExtensions(String... extensions) {
        if (extensions == null || extensions.length == 0) {
            throw new IllegalArgumentException("extensions must not be empty");
        }
        Set<String> normalized = Arrays.stream(extensions)
                .map(ExcelKitMultipartFile::normalizeExtension)
                .collect(Collectors.toUnmodifiableSet());
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("extensions must not be empty");
        }
        return normalized;
    }

    private static String normalizeExtension(String extension) {
        if (extension == null || extension.isBlank()) {
            throw new IllegalArgumentException("extension must not be blank");
        }
        String normalized = extension.trim().toLowerCase(Locale.ROOT);
        while (normalized.startsWith(".")) {
            normalized = normalized.substring(1);
        }
        if (normalized.isBlank() || normalized.contains("/") || normalized.contains("\\")) {
            throw new IllegalArgumentException("extension must be a simple file extension");
        }
        return normalized;
    }

    private static String extensionOf(String filename) {
        if (filename == null || filename.isBlank()) {
            return "";
        }
        int slash = Math.max(filename.lastIndexOf('/'), filename.lastIndexOf('\\'));
        int dot = filename.lastIndexOf('.');
        if (dot <= slash || dot == filename.length() - 1) {
            return "";
        }
        return filename.substring(dot + 1).toLowerCase(Locale.ROOT);
    }

    static String safeOriginalFilename(MultipartFile file) {
        String filename = file.getOriginalFilename();
        if (filename == null || filename.isBlank()) {
            return "<unknown>";
        }
        return filename.replaceAll("[\\p{Cntrl}\\\\/]+", "_");
    }
}
