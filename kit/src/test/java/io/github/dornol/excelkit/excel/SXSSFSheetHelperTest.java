package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.io.IOException;
import static org.junit.jupiter.api.Assertions.*;

class SXSSFSheetHelperTest {
    @Test
    void followsSheetIdentityAcrossRenameAndReorder() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var first = workbook.createSheet("First");
            var second = workbook.createSheet("Second");
            var backingFirst = workbook.getXSSFWorkbook().getSheet("First");
            var backingSecond = workbook.getXSSFWorkbook().getSheet("Second");
            workbook.setSheetName(0, "Renamed");
            workbook.setSheetOrder("Second", 0);
            assertSame(backingFirst, SXSSFSheetHelper.getXSSFSheet(first));
            assertSame(backingSecond, SXSSFSheetHelper.getXSSFSheet(second));
            assertEquals("Renamed", SXSSFSheetHelper.getXSSFSheet(first).getSheetName());
        }
    }

    @Test
    void resolvesExistingTemplateSheetAfterStreamingRowsHaveFlushed() throws IOException {
        try (var template = new XSSFWorkbook()) {
            var original = template.createSheet("Template");
            original.createRow(0).createCell(0).setCellValue("Original header");
            try (var workbook = new SXSSFWorkbook(template, 1)) {
                var sheet = workbook.getSheetAt(0);
                sheet.createRow(1).createCell(0).setCellValue("one");
                sheet.createRow(2).createCell(0).setCellValue("two");
                sheet.flushRows();
                assertSame(original, SXSSFSheetHelper.getXSSFSheet(sheet));
                assertEquals("Original header", SXSSFSheetHelper.getXSSFSheet(sheet).getRow(0).getCell(0).getStringCellValue());
            }
        }
    }

    @Test
    void removedSheetNeverResolvesToAReplacementWithTheSameName() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var removed = workbook.createSheet("Data");
            workbook.removeSheetAt(0);
            var replacement = workbook.createSheet("Data");
            var error = assertThrows(ExcelWriteException.class, () -> SXSSFSheetHelper.getXSSFSheet(removed));
            assertTrue(error.getMessage().contains("backing XSSFSheet"));
            assertSame(workbook.getXSSFWorkbook().getSheetAt(0), SXSSFSheetHelper.getXSSFSheet(replacement));
        }
    }

    @Test
    void publicApiFailureKeepsItsOriginalCause() throws IOException {
        var cause = new IllegalStateException("index lookup failed");
        try (var workbook = new SXSSFWorkbook() {
            @Override public int getSheetIndex(Sheet sheet) { throw cause; }
        }) {
            var sheet = workbook.createSheet("Data");
            var error = assertThrows(ExcelWriteException.class, () -> SXSSFSheetHelper.getXSSFSheet(sheet));
            assertSame(cause, error.getCause());
        }
    }

    @Test
    void configuredFeaturesRejectAnUnregisteredSheetInsteadOfBeingSilentlyOmitted() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = workbook.createSheet("Removed");
            workbook.removeSheetAt(0);
            assertAll(
                () -> assertThrows(ExcelWriteException.class, () -> ExcelWriteSupport.applyTabColor(sheet, new int[]{1, 2, 3})),
                () -> assertThrows(ExcelWriteException.class, () -> StructuredTableWriter.apply(sheet, "DataTable", 0, 2, java.util.List.of("Name", "Value"), "TableStyleMedium2", true)),
                () -> assertThrows(ExcelWriteException.class, () -> StructuredTableWriter.validateExistingHeaders(sheet, 0, 2)),
                () -> assertThrows(ExcelWriteException.class, () -> new ExcelChartConfig().categoryColumn(0).valueColumn(1, "Value").apply(sheet, 0, 2)),
                () -> assertThrows(ExcelWriteException.class, () -> new ExcelConditionalRule().columns(1).dataBar(ExcelColor.BLUE).apply(sheet, 0, 2, 2)),
                () -> assertThrows(ExcelWriteException.class, () -> new ExcelConditionalRule().columns(1).iconSet(ExcelConditionalRule.IconSetType.ARROWS_3).apply(sheet, 0, 2, 2))
            );
        }
    }
}
