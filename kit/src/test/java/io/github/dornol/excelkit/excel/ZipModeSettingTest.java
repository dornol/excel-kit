package io.github.dornol.excelkit.excel;

import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * The streaming workbooks ask POI for {@code Zip64Mode.AsNeeded}. POI 6's default, {@code Always}, writes a data
 * descriptor that LibreOffice and streaming ZIP readers reject, and {@code AlwaysWithCompatibility} writes a form that
 * Excel rejects ("we found a problem with some content", found on a real Excel). {@code AsNeeded} writes an ordinary
 * ZIP, and ZIP64 only for an entry above 4 GiB (see {@code Zip64HugeSheetTest}). This holds on every JDK and POI
 * version, which is why the setting itself is checked here.
 */
class ZipModeSettingTest {

    private static Zip64Mode modeOf(SXSSFWorkbook workbook) throws Exception {
        Field field = SXSSFWorkbook.class.getDeclaredField("zip64Mode");
        field.setAccessible(true);
        return (Zip64Mode) field.get(workbook);
    }

    @Test
    void aWriterWorkbookUsesAnOrdinaryZip() throws Exception {
        SXSSFWorkbook workbook = ExcelWorkbookSupport.streamingWorkbook(StreamingOptions.DEFAULT);
        try {
            assertEquals(Zip64Mode.AsNeeded, modeOf(workbook));
        } finally {
            ExcelWorkbookSupport.close(workbook);
        }
    }

    @Test
    void aWorkbookWithSharedStringsAndCompressedTempFilesUsesItToo() throws Exception {
        SXSSFWorkbook workbook = ExcelWorkbookSupport.streamingWorkbook(new StreamingOptions(10, true, true));
        try {
            assertEquals(Zip64Mode.AsNeeded, modeOf(workbook));
        } finally {
            ExcelWorkbookSupport.close(workbook);
        }
    }

    @Test
    void aTemplateWorkbookUsesItToo() throws Exception {
        try (XSSFWorkbook template = new XSSFWorkbook()) {
            template.createSheet("S");
            SXSSFWorkbook workbook = ExcelWorkbookSupport.templateStreamingWorkbook(template, 10);
            try {
                assertEquals(Zip64Mode.AsNeeded, modeOf(workbook));
            } finally {
                ExcelWorkbookSupport.close(workbook);
            }
        }
    }
}
