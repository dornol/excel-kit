package io.github.dornol.excelkit.core;

/** Optional content restrictions for untrusted Excel workbooks. */
public record ReadSecurityPolicy(boolean allowFormulas, boolean allowExternalLinks) {
    public static final ReadSecurityPolicy DEFAULT = new ReadSecurityPolicy(true, true);
    public static final ReadSecurityPolicy STRICT = new ReadSecurityPolicy(false, false);
}
