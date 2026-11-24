package io.github.dornol.excelkit.core.detect;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import org.jspecify.annotations.Nullable;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/**
 * Signature-based file type detection that does not depend on filename extensions.
 * <p>
 * An Excel 97-2003 workbook ({@code .xls}) and a password-protected {@code .xlsx} are both OLE2 containers, which this
 * detector does not look into: both are reported as {@link TabularFileType#XLS}. A password-protected workbook is read
 * without detection, giving the reader its {@code password(...)}; a reader that is handed an {@code .xls} says so in its
 * {@code ExcelReadException}.
 *
 * @since 0.21.0
 */
public final class TabularFileDetector {
    public static final int DEFAULT_SAMPLE_BYTES = 8192;
    private static final String DELIMITER_CANDIDATES = ",\t;|";
    private TabularFileDetector() {}

    /** Detects the type of a file from its content, ignoring the filename extension. */
    public static TabularFileType detect(Path path) {
        return detectDetailed(path).type();
    }

    /** Detects the type of the stream a source opens; the stream is opened and closed here. */
    public static TabularFileType detect(InputStreamSource source) {
        return detectDetailed(source).type();
    }

    /**
     * Like {@link #detect(InputStreamSource)}, returning the full {@link TabularDetectionResult};
     * samples up to {@link #DEFAULT_SAMPLE_BYTES} bytes.
     */
    public static TabularDetectionResult detectDetailed(InputStreamSource source) {
        return detectDetailed(source, DEFAULT_SAMPLE_BYTES);
    }

    /** Like {@link #detectDetailed(InputStreamSource)}, sampling at most {@code sampleBytes} bytes. */
    public static TabularDetectionResult detectDetailed(InputStreamSource source, int sampleBytes) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return detectDetailed(new BufferedInputStream(input), sampleBytes);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    /** Like {@link #detect(Path)}, returning the full {@link TabularDetectionResult}. */
    public static TabularDetectionResult detectDetailed(Path path) {
        try (InputStream input = Files.newInputStream(path)) {
            return detectDetailed(new BufferedInputStream(input));
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    /** Detects from a mark-capable stream and restores its position. */
    public static TabularFileType detect(InputStream input) {
        return detectDetailed(input).type();
    }

    /**
     * Like {@link #detect(InputStream)}, returning the full {@link TabularDetectionResult}; samples up
     * to {@link #DEFAULT_SAMPLE_BYTES} bytes.
     */
    public static TabularDetectionResult detectDetailed(InputStream input) {
        return detectDetailed(input, DEFAULT_SAMPLE_BYTES);
    }

    /** Detects from a mark-capable stream using at most {@code sampleBytes}. */
    public static TabularDetectionResult detectDetailed(InputStream input, int sampleBytes) {
        return detectDetailed(input, sampleBytes, null, '"', '\0');
    }

    /**
     * Detects file type and CSV delimiter from a bounded sample using the supplied quoting rules.
     * A supplied charset takes precedence over the text encoding inferred from the sample.
     */
    public static TabularDetectionResult detectDetailed(InputStream input, int sampleBytes,
                                                         @Nullable Charset configuredCharset,
                                                         char quoteChar, char escapeChar) {
        return detectDetailed(input, sampleBytes, configuredCharset, quoteChar, escapeChar, 0);
    }

    /** Detects the delimiter from the specified zero-based logical CSV header record. */
    public static TabularDetectionResult detectDetailed(InputStream input, int sampleBytes,
                                                         @Nullable Charset configuredCharset,
                                                         char quoteChar, char escapeChar, int headerRowIndex) {
        Objects.requireNonNull(input, "input cannot be null");
        if (sampleBytes < 64) throw new IllegalArgumentException("sampleBytes must be at least 64");
        if (headerRowIndex < 0) throw new IllegalArgumentException("headerRowIndex cannot be negative");
        if (!input.markSupported()) throw new IllegalArgumentException("input must support mark/reset");
        try {
            input.mark(sampleBytes);
            byte[] bytes = input.readNBytes(sampleBytes);
            input.reset();
            if (starts(bytes, 0x50, 0x4b, 0x03, 0x04)) return new TabularDetectionResult(
                    TabularFileType.XLSX, DetectionConfidence.HIGH, null, null);
            if (starts(bytes, 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)) return new TabularDetectionResult(
                    TabularFileType.XLS, DetectionConfidence.HIGH, null, null);
            Charset charset = configuredCharset == null ? charset(bytes) : configuredCharset;
            String text = new String(bytes, charset);
            if (looksTextual(text)) {
                Character delimiter = delimiter(text, quoteChar, escapeChar, headerRowIndex);
                return new TabularDetectionResult(TabularFileType.CSV,
                        delimiter == null ? DetectionConfidence.LOW : DetectionConfidence.MEDIUM,
                        charset, delimiter);
            }
            return new TabularDetectionResult(TabularFileType.UNKNOWN, DetectionConfidence.LOW, null, null);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    private static Charset charset(byte[] bytes) {
        if (starts(bytes, 0xff, 0xfe)) return StandardCharsets.UTF_16LE;
        if (starts(bytes, 0xfe, 0xff)) return StandardCharsets.UTF_16BE;
        return StandardCharsets.UTF_8;
    }

    private static @Nullable Character delimiter(String text, char quoteChar, char escapeChar, int headerRowIndex) {
        int[] counts = new int[DELIMITER_CANDIDATES.length()];
        boolean quoted = false;
        boolean atFieldStart = true;
        int start = text.startsWith("\uFEFF") ? 1 : 0;
        int recordIndex = 0;
        int recordStart = start;
        for (int i = start; i < text.length(); i++) {
            char c = text.charAt(i);
            char next = i + 1 < text.length() ? text.charAt(i + 1) : '\0';
            if (escapeChar != '\0' && c == escapeChar && (quoted || !atFieldStart)
                    && (next == quoteChar || next == escapeChar || DELIMITER_CANDIDATES.indexOf(next) >= 0)) {
                i++;
            } else if (quoteChar != '\0' && c == quoteChar && (quoted || atFieldStart)) {
                if (quoted && next == quoteChar) {
                    i++;
                } else {
                    quoted = !quoted;
                    atFieldStart = false;
                }
            } else if (!quoted) {
                if (c == '\r' || c == '\n') {
                    if (recordIndex == headerRowIndex) break;
                    if (c == '\r' && next == '\n') i++;
                    recordIndex++;
                    recordStart = i + 1;
                    atFieldStart = true;
                    continue;
                }
                int candidate = DELIMITER_CANDIDATES.indexOf(c);
                if (candidate >= 0) {
                    if (recordIndex == headerRowIndex) counts[candidate]++;
                    atFieldStart = true;
                } else if (c != ' ') {
                    atFieldStart = false;
                }
            }
        }
        if (headerRowIndex > 0 && (recordIndex != headerRowIndex || recordStart == text.length())) {
            throw new ExcelKitException("CSV header is outside the detection sample; increase detectionSampleBytes "
                    + "or configure the delimiter");
        }
        int best = -1;
        for (int i = 0; i < counts.length; i++) {
            if (counts[i] > 0 && (best < 0 || counts[i] > counts[best])) best = i;
        }
        return best < 0 ? null : DELIMITER_CANDIDATES.charAt(best);
    }

    private static boolean starts(byte[] bytes, int... signature) {
        if (bytes.length < signature.length) return false;
        for (int i = 0; i < signature.length; i++) if ((bytes[i] & 0xff) != signature[i]) return false;
        return true;
    }

    private static boolean looksTextual(String text) {
        int controls = 0;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (c == 0) return false;
            if (c < 0x09 || (c > 0x0d && c < 0x20)) controls++;
        }
        return controls * 20L < text.length();
    }
}
