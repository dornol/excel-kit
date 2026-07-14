package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadSecurityException;
import io.github.dornol.excelkit.core.ReadSecurityPolicy;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Bounded preflight scanner for formula and external-link policy enforcement. */
final class ExcelSecurityScanner {
    private ExcelSecurityScanner() {}

    static void scan(Path file, ReadSecurityPolicy policy) throws IOException {
        if (policy.allowFormulas() && policy.allowExternalLinks()) return;
        long total = 0;
        try (ZipFile zip = new ZipFile(file.toFile())) {
            var entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if (!policy.allowExternalLinks() && name.startsWith("xl/externalLinks/"))
                    throw security(ReadSecurityException.Reason.EXTERNAL_LINK, "External workbook links are not allowed");
                if (!policy.allowFormulas() && name.startsWith("xl/worksheets/") && name.endsWith(".xml")) {
                    validateEntry(entry, policy);
                    try (InputStream input = zip.getInputStream(entry)) {
                        ScanResult result = scanFormula(input, policy.maxScannedEntryBytes());
                        total += result.bytes();
                        if (total > policy.maxTotalScannedBytes())
                            throw security(ReadSecurityException.Reason.TOTAL_SCAN_SIZE,
                                    "Workbook security scan exceeds total byte limit");
                        if (result.formula()) throw security(ReadSecurityException.Reason.FORMULA,
                                "Excel formulas are not allowed");
                    }
                }
            }
        }
    }

    private static void validateEntry(ZipEntry entry, ReadSecurityPolicy policy) {
        long size = entry.getSize(), compressed = entry.getCompressedSize();
        if (size > policy.maxScannedEntryBytes()) throw security(ReadSecurityException.Reason.ENTRY_SIZE,
                "Worksheet XML exceeds security scan entry limit: " + size);
        if (size > 0 && compressed > 0 && (double) size / compressed > policy.maxCompressionRatio())
            throw security(ReadSecurityException.Reason.COMPRESSION_RATIO,
                    "Worksheet XML exceeds compression ratio limit");
    }

    private static ScanResult scanFormula(InputStream input, long maximum) throws IOException {
        int state = 0; long bytes = 0;
        for (int value; (value = input.read()) >= 0;) {
            if (++bytes > maximum) throw security(ReadSecurityException.Reason.ENTRY_SIZE,
                    "Worksheet XML exceeds security scan entry limit");
            if (state == 0) state = value == '<' ? 1 : 0;
            else if (state == 1) state = value == 'f' ? 2 : (value == '<' ? 1 : 0);
            else {
                if (value == '>' || Character.isWhitespace(value)) return new ScanResult(true, bytes);
                state = value == '<' ? 1 : 0;
            }
        }
        return new ScanResult(false, bytes);
    }

    private static ReadSecurityException security(ReadSecurityException.Reason reason, String message) {
        return new ReadSecurityException(reason, message);
    }
    private record ScanResult(boolean formula, long bytes) {}
}
