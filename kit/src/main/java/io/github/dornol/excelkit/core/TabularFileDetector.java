package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/** Signature-based file type detection that does not depend on filename extensions. */
public final class TabularFileDetector {
    private TabularFileDetector() {}

    public static TabularFileType detect(Path path) {
        try (InputStream input = Files.newInputStream(path)) { return detect(input); }
        catch (IOException e) { throw new ExcelKitException("Failed to inspect input", e); }
    }

    /** Detects from a mark-capable stream and restores its position. */
    public static TabularFileType detect(InputStream input) {
        java.util.Objects.requireNonNull(input, "input cannot be null");
        if (!input.markSupported()) throw new IllegalArgumentException("input must support mark/reset");
        try {
            input.mark(8192);
            byte[] bytes = input.readNBytes(8192);
            input.reset();
            if (starts(bytes, 0x50, 0x4b, 0x03, 0x04)) return TabularFileType.XLSX;
            if (starts(bytes, 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)) return TabularFileType.XLS;
            if (looksTextual(bytes)) return TabularFileType.CSV;
            return TabularFileType.UNKNOWN;
        } catch (IOException e) {
            throw new ExcelKitException("Failed to inspect input", e);
        }
    }

    private static boolean starts(byte[] bytes, int... signature) {
        if (bytes.length < signature.length) return false;
        for (int i = 0; i < signature.length; i++) if ((bytes[i] & 0xff) != signature[i]) return false;
        return true;
    }

    private static boolean looksTextual(byte[] bytes) {
        if (bytes.length == 0) return false;
        int controls = 0;
        for (byte value : bytes) {
            int b = value & 0xff;
            if (b == 0) return false;
            if (b < 0x09 || (b > 0x0d && b < 0x20)) controls++;
        }
        return controls * 20 < bytes.length;
    }
}
