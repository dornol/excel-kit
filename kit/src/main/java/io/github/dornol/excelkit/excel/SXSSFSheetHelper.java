package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Package-private utility that consolidates reflective access to the
 * underlying {@link XSSFSheet} from an {@link SXSSFSheet}.
 * <p>
 * SXSSFSheet does not expose its backing XSSFSheet via public API, so
 * reflection on the {@code _sh} field is required for features such as
 * chart creation and tab color.
 *
 * @author dhkim
 * @since 0.7.0
 */
class SXSSFSheetHelper {
    private static final Logger log = LoggerFactory.getLogger(SXSSFSheetHelper.class);

    private SXSSFSheetHelper() {
    }

    /**
     * Returns the underlying {@link XSSFSheet} from the given {@link SXSSFSheet},
     * or {@code null} if reflective access fails.
     *
     * @param sheet the SXSSFSheet to unwrap
     * @return the underlying XSSFSheet, or {@code null}
     */
    static @Nullable XSSFSheet getXSSFSheet(SXSSFSheet sheet) {
        try {
            var field = SXSSFSheet.class.getDeclaredField("_sh");
            field.setAccessible(true);
            return (XSSFSheet) field.get(sheet);
        } catch (Exception e) {
            log.warn("Failed to access underlying XSSFSheet via reflection. "
                    + "Features requiring XSSFSheet (charts, tab colors) will be unavailable.", e);
            return null;
        }
    }

    /**
     * Returns the underlying {@link XSSFSheet} from the given {@link SXSSFSheet},
     * or throws an {@link ExcelWriteException} if reflective access fails.
     *
     * @param sheet the SXSSFSheet to unwrap
     * @return the underlying XSSFSheet
     * @throws ExcelWriteException if the underlying sheet cannot be accessed
     */
    static XSSFSheet getXSSFSheetOrThrow(SXSSFSheet sheet) {
        XSSFSheet result = getXSSFSheet(sheet);
        if (result == null) {
            throw new ExcelWriteException("Failed to access underlying XSSFSheet");
        }
        return result;
    }
}
