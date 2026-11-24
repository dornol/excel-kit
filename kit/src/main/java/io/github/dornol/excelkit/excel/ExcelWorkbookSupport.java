package io.github.dornol.excelkit.excel;

import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.Collections;
import java.util.Locale;
import java.util.Map;
import java.util.WeakHashMap;

/** Workbook-level operations independent of sheet row rendering. */
final class ExcelWorkbookSupport {

    /** The streaming workbooks that were created with a shared strings table. */
    private static final Map<SXSSFWorkbook, Boolean> SHARED_STRINGS =
            Collections.synchronizedMap(new WeakHashMap<>());

    /** Creates the streaming workbook a writer works on, and remembers whether it keeps shared strings. */
    static SXSSFWorkbook streamingWorkbook(StreamingOptions streaming) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(null, streaming.rowAccessWindowSize(),
                streaming.compressTempFiles(), streaming.useSharedStrings());
        if (streaming.useSharedStrings()) SHARED_STRINGS.put(workbook, Boolean.TRUE);
        readableByStreamingZipReaders(workbook);
        return workbook;
    }

    /**
     * POI writes every entry of a streamed workbook with an 8-byte (ZIP64) data descriptor by default
     * ({@code Zip64Mode.Always}), which LibreOffice and {@code java.util.zip.ZipInputStream} cannot read.
     * {@code AlwaysWithCompatibility}, which keeps entries beyond 4 GiB possible, writes a form that Excel rejects
     * ("we found a problem with some content", opened on a real Excel). {@code AsNeeded} writes an ordinary ZIP for
     * every entry that fits 32-bit sizes, which LibreOffice and streaming readers read, and ZIP64 only for an entry
     * that does not (a worksheet of 7 GB was written and read back whole).
     */
    private static void readableByStreamingZipReaders(SXSSFWorkbook workbook) {
        workbook.setZip64Mode(Zip64Mode.AsNeeded);
    }

    /** Creates the streaming workbook of a template writer: it wraps the template and keeps shared strings. */
    static SXSSFWorkbook templateStreamingWorkbook(XSSFWorkbook template, int rowAccessWindowSize) {
        SXSSFWorkbook workbook = new SXSSFWorkbook(template, rowAccessWindowSize, false, true);
        SHARED_STRINGS.put(workbook, Boolean.TRUE);
        readableByStreamingZipReaders(workbook);
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
    static void close(SXSSFWorkbook workbook) throws IOException {
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
            LoggerFactory.getLogger(ExcelWorkbookSupport.class).debug("SXSSFWorkbook.dispose() is gone; close() did the work");
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

    /**
     * Checks that a workbook keeps a visible sheet (Excel cannot open one with none) and that it opens on one.
     *
     * @throws ExcelWriteException if every sheet is hidden
     */
    static void requireVisibleSheet(org.apache.poi.xssf.streaming.SXSSFWorkbook wb) {
        int first = -1;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (wb.getSheetVisibility(i) == org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE) {
                first = i;
                break;
            }
        }
        if (first < 0) {
            if (wb.getNumberOfSheets() == 0) return;
            throw new ExcelWriteException("Every sheet of the workbook is hidden; Excel cannot open a workbook "
                    + "without a visible sheet, so leave at least one sheet visible");
        }
        if (wb.getSheetVisibility(wb.getActiveSheetIndex()) != org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE) {
            wb.setActiveSheet(first);
            SXSSFSheetHelper.getXSSFSheet(wb.getSheetAt(first)).setSelected(true);
        }
    }

    /**
     * Rejects a sheet name Excel cannot hold, so a mistake surfaces where the name is given instead of
     * when the first sheet is created. The length is not checked here: a long base name is shortened for
     * the sheets that roll over.
     */
    static String requireValidSheetName(String name) {
        if (name.isEmpty()) {
            throw new IllegalArgumentException("Sheet name must not be empty");
        }
        for (int i = 0; i < name.length(); i++) {
            if ("\\/?*[]:".indexOf(name.charAt(i)) >= 0) {
                throw new IllegalArgumentException("Invalid char (" + name.charAt(i) + ") found at index (" + i
                        + ") in sheet name '" + name + "'");
            }
        }
        if (name.charAt(0) == '\'' || name.charAt(name.length() - 1) == '\'') {
            throw new IllegalArgumentException("Invalid sheet name '" + name
                    + "'. Sheet names must not begin or end with (').");
        }
        return name;
    }

    private static final int MAX_SHEET_NAME_LENGTH = 31;

    /**
     * Name of the {@code index}-th sheet (0-based) of a rolled-over write: the first keeps the base
     * name, later ones get a " (n)" suffix. Excel limits sheet names to 31 characters, so the base is
     * shortened to keep the suffix and the names unique.
     */
    static String rolloverSheetName(String baseName, int index) {
        if (index == 0) return baseName;
        String suffix = " (" + (index + 1) + ")";
        int baseLength = Math.min(baseName.length(), Math.max(0, MAX_SHEET_NAME_LENGTH - suffix.length()));
        if (baseLength > 0 && baseLength < baseName.length()
                && Character.isHighSurrogate(baseName.charAt(baseLength - 1))) {
            baseLength--; // do not split a surrogate pair
        }
        return baseName.substring(0, baseLength) + suffix;
    }
}
