package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.detect.TabularFileDetector;
import io.github.dornol.excelkit.core.detect.TabularFileType;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.charset.Charset;
import java.nio.charset.CoderResult;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;

/**
 * Small adapter for opening upload streams from Spring {@link MultipartFile}s.
 *
 * @since 0.18.4
 */
public final class ExcelKitMultipartFile {
    private static final int CONTENT_PREFIX_BYTES = 4096;
    private static final Set<String> EXCEL_AND_CSV_EXTENSIONS = Set.of("xlsx", "xlsm", "csv");

    private ExcelKitMultipartFile() {
    }

    /**
     * Opens the stream of an upload.
     *
     * @param file the upload
     * @return the upload's stream; the caller closes it
     * @throws ExcelKitUploadException if the file is empty or cannot be opened
     */
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

    /**
     * Requires an Excel or CSV extension and verifies that the uploaded content
     * looks like the declared format.
     *
     * @param file upload file
     * @return the same file for fluent validation before passing it to {@link ExcelKitUpload}
     * @since 0.19.0
     */
    public static MultipartFile requireExcelOrCsvContent(MultipartFile file) {
        return requireExcelOrCsvContent(file, StandardCharsets.UTF_8);
    }

    /**
     * Like {@link #requireExcelOrCsvContent(MultipartFile)}, for a CSV that is not UTF-8 (for example the
     * CP949 a Korean Excel saves as "CSV"). The CSV content must decode in {@code csvCharset} without errors;
     * pass the same charset to the reader ({@code CsvReader.charset(...)}). The encoding is not guessed:
     * a short file can be valid in several encodings. It has no effect on an {@code .xlsx}/{@code .xlsm} upload.
     *
     * @param file       upload file
     * @param csvCharset the encoding the CSV content must have
     * @return the same file for fluent validation before passing it to {@link ExcelKitUpload}
     * @since 0.29.0
     */
    public static MultipartFile requireExcelOrCsvContent(MultipartFile file, Charset csvCharset) {
        Objects.requireNonNull(csvCharset, "csvCharset must not be null");
        requireExcelOrCsv(file);
        String extension = extensionOf(file.getOriginalFilename());
        byte[] prefix = readPrefix(file, CONTENT_PREFIX_BYTES);
        if (extension.equals("xlsx") || extension.equals("xlsm")) {
            // The signature of a ZIP package, not just "PK": a text file that starts with those letters is no workbook.
            if (TabularFileDetector.detect(new ByteArrayInputStream(prefix)) != TabularFileType.XLSX) {
                throw new ExcelKitUploadException("Upload file content is not a valid XLSX/XLSM package");
            }
            return file;
        }
        if (extension.equals("csv")) {
            // A prefix shorter than the limit is the whole file; otherwise it may end inside a character.
            requireTextCsv(prefix, prefix.length < CONTENT_PREFIX_BYTES, csvCharset);
            return file;
        }
        throw new ExcelKitUploadException("Unsupported upload file extension: " + extension);
    }

    private static byte[] readPrefix(MultipartFile file, int maxBytes) {
        try (InputStream inputStream = open(file)) {
            return inputStream.readNBytes(maxBytes);
        } catch (IOException e) {
            throw new ExcelKitUploadException("Failed to inspect upload file: " + safeOriginalFilename(file), e);
        }
    }

    private static void requireTextCsv(byte[] prefix, boolean wholeFile, Charset charset) {
        if (prefix.length == 0) {
            throw new ExcelKitUploadException("Upload file is empty");
        }
        // A workbook (or any binary) with a .csv name is not text CSV; the detector knows the signatures and
        // tells text from binary by its share of control characters.
        if (TabularFileDetector.detectDetailed(new ByteArrayInputStream(prefix),
                TabularFileDetector.DEFAULT_SAMPLE_BYTES, charset, '"', '\0').type() != TabularFileType.CSV) {
            throw new ExcelKitUploadException("Upload file content is not text CSV");
        }
        var decoder = charset.newDecoder()
                .onMalformedInput(CodingErrorAction.REPORT)
                .onUnmappableCharacter(CodingErrorAction.REPORT);
        // With endOfInput=false an incomplete trailing sequence is left undecoded instead of reported.
        CoderResult result = decoder.decode(ByteBuffer.wrap(prefix), CharBuffer.allocate(prefix.length), wholeFile);
        if (result.isError()) {
            throw new ExcelKitUploadException("Upload file content is not " + charset.name() + " CSV");
        }
    }

    private static Set<String> normalizeExtensions(String... extensions) {
        if (extensions == null || extensions.length == 0) {
            throw new IllegalArgumentException("extensions must not be empty");
        }
        // Sorted, so the message that lists the allowed extensions reads the same in every JVM.
        Set<String> normalized = Arrays.stream(extensions)
                .map(ExcelKitMultipartFile::normalizeExtension)
                .collect(Collectors.toCollection(TreeSet::new));
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
