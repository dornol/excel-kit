package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;

import java.util.Locale;

/** Workbook-level operations independent of sheet row rendering. */
final class ExcelWorkbookSupport {
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
