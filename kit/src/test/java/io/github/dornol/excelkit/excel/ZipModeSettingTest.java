package io.github.dornol.excelkit.excel;

import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * The streaming workbooks ask POI for {@code Zip64Mode.AlwaysWithCompatibility}. POI 6's default, {@code Always},
 * writes a data descriptor that LibreOffice and streaming ZIP readers reject; {@code AsNeeded} would make a worksheet
 * above 4 GiB fail while it is written. This holds on every JDK and POI version, unlike reading the result back with
 * {@code ZipInputStream}, which depends on the JDK (see {@code StreamingZipCompatibilityTest}).
 */
class ZipModeSettingTest {

    private static Zip64Mode modeOf(SXSSFWorkbook workbook) throws Exception {
        Field field = SXSSFWorkbook.class.getDeclaredField("zip64Mode");
        field.setAccessible(true);
        return (Zip64Mode) field.get(workbook);
    }

    @Test
    void aWriterWorkbookUsesTheCompatibleMode() throws Exception {
        SXSSFWorkbook workbook = ExcelWorkbookSupport.streamingWorkbook(StreamingOptions.DEFAULT);
        try {
            assertEquals(Zip64Mode.AlwaysWithCompatibility, modeOf(workbook));
        } finally {
            ExcelWorkbookSupport.close(workbook);
        }
    }

    @Test
    void aWorkbookWithSharedStringsAndCompressedTempFilesUsesItToo() throws Exception {
        SXSSFWorkbook workbook = ExcelWorkbookSupport.streamingWorkbook(new StreamingOptions(10, true, true));
        try {
            assertEquals(Zip64Mode.AlwaysWithCompatibility, modeOf(workbook));
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
                assertEquals(Zip64Mode.AlwaysWithCompatibility, modeOf(workbook));
            } finally {
                ExcelWorkbookSupport.close(workbook);
            }
        }
    }
}
