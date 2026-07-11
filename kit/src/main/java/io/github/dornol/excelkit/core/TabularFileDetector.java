package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/** Signature-based file type detection that does not depend on filename extensions. */
public final class TabularFileDetector {
    private TabularFileDetector() {}

    public static TabularFileType detect(Path path) {
        return detectDetailed(path).type();
    }

    public static TabularDetectionResult detectDetailed(Path path) {
        try (InputStream input = Files.newInputStream(path)) {
            return detectDetailed(new java.io.BufferedInputStream(input));
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    /** Detects from a mark-capable stream and restores its position. */
    public static TabularFileType detect(InputStream input) {
        return detectDetailed(input).type();
    }

    public static TabularDetectionResult detectDetailed(InputStream input) {
        java.util.Objects.requireNonNull(input, "input cannot be null");
        if (!input.markSupported()) throw new IllegalArgumentException("input must support mark/reset");
        try {
            input.mark(8192);
            byte[] bytes = input.readNBytes(8192);
            input.reset();
            if (starts(bytes, 0x50, 0x4b, 0x03, 0x04)) return new TabularDetectionResult(
                    TabularFileType.XLSX, DetectionConfidence.HIGH, null, null);
            if (starts(bytes, 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)) return new TabularDetectionResult(
                    TabularFileType.XLS, DetectionConfidence.HIGH, null, null);
            java.nio.charset.Charset charset = charset(bytes);
            if (looksTextual(bytes, charset)) {
                Character delimiter = delimiter(bytes, charset);
                return new TabularDetectionResult(TabularFileType.CSV,
                        delimiter == null ? DetectionConfidence.LOW : DetectionConfidence.MEDIUM,
                        charset, delimiter);
            }
            return new TabularDetectionResult(TabularFileType.UNKNOWN, DetectionConfidence.LOW, null, null);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    private static java.nio.charset.Charset charset(byte[] bytes) {
        if (starts(bytes, 0xff, 0xfe)) return java.nio.charset.StandardCharsets.UTF_16LE;
        if (starts(bytes, 0xfe, 0xff)) return java.nio.charset.StandardCharsets.UTF_16BE;
        return java.nio.charset.StandardCharsets.UTF_8;
    }

    private static Character delimiter(byte[] bytes, java.nio.charset.Charset charset) {
        String text = new String(bytes, charset);
        String first = text.lines().findFirst().orElse("");
        char best = 0; int count = 0;
        for (char candidate : new char[]{',', '\t', ';', '|'}) {
            int found = 0;
            for (int i = 0; i < first.length(); i++) if (first.charAt(i) == candidate) found++;
            if (found > count) { count = found; best = candidate; }
        }
        return count == 0 ? null : best;
    }

    private static boolean starts(byte[] bytes, int... signature) {
        if (bytes.length < signature.length) return false;
        for (int i = 0; i < signature.length; i++) if ((bytes[i] & 0xff) != signature[i]) return false;
        return true;
    }

    private static boolean looksTextual(byte[] bytes, java.nio.charset.Charset charset) {
        if (bytes.length == 0) return false;
        if (charset.equals(java.nio.charset.StandardCharsets.UTF_16LE)
                || charset.equals(java.nio.charset.StandardCharsets.UTF_16BE)) {
            String text = new String(bytes, charset);
            return text.chars().noneMatch(ch -> ch == 0 || (ch < 0x09) || (ch > 0x0d && ch < 0x20));
        }
        int controls = 0;
        for (byte value : bytes) {
            int b = value & 0xff;
            if (b == 0) return false;
            if (b < 0x09 || (b > 0x0d && b < 0x20)) controls++;
        }
        return controls * 20 < bytes.length;
    }
}
