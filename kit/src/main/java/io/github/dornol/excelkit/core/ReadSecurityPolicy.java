package io.github.dornol.excelkit.core;

/** Optional content restrictions for untrusted Excel workbooks. */
public record ReadSecurityPolicy(boolean allowFormulas, boolean allowExternalLinks,
        long maxScannedEntryBytes, long maxTotalScannedBytes, double maxCompressionRatio) {
    public static final ReadSecurityPolicy DEFAULT = new ReadSecurityPolicy(true, true,
            32L * 1024 * 1024, 128L * 1024 * 1024, 100.0);
    public static final ReadSecurityPolicy STRICT = new ReadSecurityPolicy(false, false,
            16L * 1024 * 1024, 64L * 1024 * 1024, 50.0);
    public ReadSecurityPolicy {
        if (maxScannedEntryBytes <= 0 || maxTotalScannedBytes <= 0 || maxCompressionRatio <= 0)
            throw new IllegalArgumentException("security scan limits must be positive");
    }
}
