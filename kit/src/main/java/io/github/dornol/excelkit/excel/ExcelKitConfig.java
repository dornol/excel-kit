package io.github.dornol.excelkit.excel;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.util.IOUtils;

/**
 * Application-level configuration for excel-kit.
 * <p>
 * These settings affect JVM-global Apache POI limits and should be called once
 * at application startup — typically in a Spring {@code @PostConstruct} or
 * {@code main()} method.
 *
 * @author dhkim
 * @since 0.14.0
 */
public final class ExcelKitConfig {

    private static final int DEFAULT_MAX_FILE_COUNT = 1_000_000;
    private static final int DEFAULT_MAX_BYTE_ARRAY_SIZE = 500_000_000;

    private ExcelKitConfig() {}

    /**
     * Configures Apache POI's internal limits for reading large Excel files.
     * <p>
     * Adjusts:
     * <ul>
     *     <li>{@code ZipSecureFile.setMaxFileCount(1,000,000)} — max internal zip entries</li>
     *     <li>{@code IOUtils.setByteArrayMaxOverride(500,000,000)} — max in-memory byte array size</li>
     * </ul>
     * <p>
     * <b>Note:</b> These are JVM-global settings and affect all POI operations in the same process.
     */
    public static void configureLargeFileSupport() {
        configureLargeFileSupport(DEFAULT_MAX_FILE_COUNT, DEFAULT_MAX_BYTE_ARRAY_SIZE);
    }

    /**
     * Configures Apache POI's internal limits with custom values.
     *
     * @param maxFileCount     Maximum number of zip entries (default: 1,000,000)
     * @param maxByteArraySize Maximum byte array size in bytes (default: 500,000,000)
     */
    public static void configureLargeFileSupport(int maxFileCount, int maxByteArraySize) {
        ZipSecureFile.setMaxFileCount(maxFileCount);
        IOUtils.setByteArrayMaxOverride(maxByteArraySize);
    }
}
