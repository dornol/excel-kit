package io.github.dornol.excelkit.core;

/**
 * Optional content restrictions for untrusted Excel workbooks.
 *
 * <p>The five-argument constructor keeps the pre-entry-count API and uses the
 * default ZIP entry limit of 50,000. Use the six-argument constructor when a
 * stricter entry-count limit is required.</p>
 *
 * @param allowFormulas whether formulas are permitted
 * @param allowExternalLinks whether external workbook links are permitted
 * @param maxScannedEntryBytes maximum decompressed size of a scanned entry
 * @param maxTotalScannedBytes maximum total bytes inspected by the scanner
 * @param maxCompressionRatio maximum allowed uncompressed/compressed ratio
 * @param maxZipEntries maximum number of ZIP entries in the workbook
 */
public record ReadSecurityPolicy(boolean allowFormulas, boolean allowExternalLinks,
        long maxScannedEntryBytes, long maxTotalScannedBytes, double maxCompressionRatio,
        int maxZipEntries) {
    public static final ReadSecurityPolicy DEFAULT = new ReadSecurityPolicy(true, true,
            32L * 1024 * 1024, 128L * 1024 * 1024, 100.0, 50_000);
    public static final ReadSecurityPolicy STRICT = new ReadSecurityPolicy(false, false,
            16L * 1024 * 1024, 64L * 1024 * 1024, 50.0, 10_000);

    /** Creates a policy using the default maximum of 50,000 ZIP entries. */
    public ReadSecurityPolicy(boolean allowFormulas, boolean allowExternalLinks,
                              long maxScannedEntryBytes, long maxTotalScannedBytes,
                              double maxCompressionRatio) {
        this(allowFormulas, allowExternalLinks, maxScannedEntryBytes, maxTotalScannedBytes,
                maxCompressionRatio, 50_000);
    }

    public ReadSecurityPolicy {
        if (maxScannedEntryBytes <= 0 || maxTotalScannedBytes <= 0 || maxCompressionRatio <= 0
                || maxZipEntries <= 0)
            throw new IllegalArgumentException("security scan limits must be positive");
    }
}
