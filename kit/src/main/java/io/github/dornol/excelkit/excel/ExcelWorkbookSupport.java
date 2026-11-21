package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;

import java.util.Locale;

/** Workbook-level operations independent of sheet row rendering. */
final class ExcelWorkbookSupport {

    /** The streaming workbooks that were created with a shared strings table. */
    private static final java.util.Map<SXSSFWorkbook, Boolean> SHARED_STRINGS =
            java.util.Collections.synchronizedMap(new java.util.WeakHashMap<>());

    /** Creates the streaming workbook a writer works on, and remembers whether it keeps shared strings. */
    static SXSSFWorkbook streamingWorkbook(StreamingOptions streaming) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(null, streaming.rowAccessWindowSize(),
                streaming.compressTempFiles(), streaming.useSharedStrings());
        if (streaming.useSharedStrings()) SHARED_STRINGS.put(workbook, Boolean.TRUE);
        return workbook;
    }

    /** Creates the streaming workbook of a template writer: it wraps the template and keeps shared strings. */
    static SXSSFWorkbook templateStreamingWorkbook(org.apache.poi.xssf.usermodel.XSSFWorkbook template, int rowAccessWindowSize) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(template, rowAccessWindowSize, false, true);
        SHARED_STRINGS.put(workbook, Boolean.TRUE);
        return workbook;
    }

    /**
     * Whether the workbook keeps shared strings. Without them POI writes every string cell as its plain text,
     * so formatting inside a cell (rich text) cannot be kept.
     */
    static boolean usesSharedStrings(SXSSFWorkbook workbook) {
        return SHARED_STRINGS.containsKey(workbook);
    }

    /**
     * Closes the workbook and deletes the temp files that back its sheets.
     * <p>
     * Older Apache POI releases (for example 5.2.5) close the sheet writers in
     * {@code SXSSFWorkbook.close()} but do not delete their temp files; newer ones call
     * {@code dispose()} from {@code close()}. Calling {@code dispose()} here keeps every supported
     * POI version from leaving one temp file per sheet behind. A second {@code dispose()} is a no-op.
     */
    static void close(SXSSFWorkbook workbook) throws java.io.IOException {
        try {
            workbook.close();
        } finally {
            dispose(workbook);
        }
    }

    /**
     * {@code dispose()} is deprecated in newer POI releases, where {@code close()} already does the
     * work. Should a release drop it, that release's {@code close()} is all that is needed.
     */
    @SuppressWarnings("deprecation")
    private static void dispose(SXSSFWorkbook workbook) {
        try {
            workbook.dispose();
        } catch (NoSuchMethodError removed) {
            // nothing left to dispose beyond what close() did
        }
    }
    private ExcelWorkbookSupport() { }

    static void applyProtection(SXSSFWorkbook workbook, @Nullable String password) {
        if (password == null) return;
        workbook.getXSSFWorkbook().lockStructure();
        workbook.getXSSFWorkbook().setWorkbookPassword(password, null);
    }

    static void applyDocumentProperty(SXSSFWorkbook workbook, String key, String value) {
        var properties = workbook.getXSSFWorkbook().getProperties();
        var core = properties.getCoreProperties();
        switch (key.toLowerCase(Locale.ROOT)) {
            case "title" -> core.setTitle(value);
            case "subject" -> core.setSubjectProperty(value);
            case "author", "creator" -> core.setCreator(value);
            case "keywords" -> core.setKeywords(value);
            case "description" -> core.setDescription(value);
            case "category" -> core.setCategory(value);
            default -> {
                var custom = properties.getCustomProperties();
                if (custom.contains(key)) custom.getProperty(key).setLpwstr(value);
                else custom.addProperty(key, value);
            }
        }
    }
}
