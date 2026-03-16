package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link SXSSFSheetHelper} — reflection-based XSSFSheet access.
 */
class SXSSFSheetHelperTest {

    private SXSSFWorkbook workbook;

    @BeforeEach
    void setUp() {
        workbook = new SXSSFWorkbook();
    }

    @AfterEach
    void tearDown() throws IOException {
        workbook.close();
    }

    @Test
    void getXSSFSheet_returnsUnderlyingSheet() {
        SXSSFSheet sxssfSheet = workbook.createSheet("Test");
        XSSFSheet result = SXSSFSheetHelper.getXSSFSheet(sxssfSheet);
        assertNotNull(result, "Should return the underlying XSSFSheet");
        assertEquals("Test", result.getSheetName(), "Sheet name should match");
    }

    @Test
    void getXSSFSheetOrThrow_returnsUnderlyingSheet() {
        SXSSFSheet sxssfSheet = workbook.createSheet("Test");
        XSSFSheet result = SXSSFSheetHelper.getXSSFSheetOrThrow(sxssfSheet);
        assertNotNull(result);
        assertEquals("Test", result.getSheetName());
    }

    @Test
    void getXSSFSheet_multipleSheets_returnsCorrectOne() {
        SXSSFSheet sheet1 = workbook.createSheet("Sheet1");
        SXSSFSheet sheet2 = workbook.createSheet("Sheet2");

        XSSFSheet xssf1 = SXSSFSheetHelper.getXSSFSheet(sheet1);
        XSSFSheet xssf2 = SXSSFSheetHelper.getXSSFSheet(sheet2);

        assertNotNull(xssf1);
        assertNotNull(xssf2);
        assertEquals("Sheet1", xssf1.getSheetName());
        assertEquals("Sheet2", xssf2.getSheetName());
        assertNotSame(xssf1, xssf2);
    }
}
