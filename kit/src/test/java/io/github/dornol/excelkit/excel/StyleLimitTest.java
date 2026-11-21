package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.testsupport.DisposingSxssfWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** A workbook that cannot take another style says why, instead of failing inside POI. */
class StyleLimitTest {

    @Test
    void theLimitIsTheXlsxOne() {
        assertEquals(64_000, ExcelWriteSupport.MAX_CELL_STYLES);
    }

    @Test
    void aStyleIsCreatedWhileThereIsRoom() {
        SXSSFWorkbook workbook = new DisposingSxssfWorkbook();
        try {
            int before = workbook.getNumCellStyles();
            assertDoesNotThrow(() -> ExcelWriteSupport.createDerivedStyle(workbook, before + 1));
            assertEquals(before + 1, workbook.getNumCellStyles());
        } finally {
            workbook.dispose();
        }
    }

    @Test
    void aFullWorkbookGetsAnExplanation() {
        SXSSFWorkbook workbook = new DisposingSxssfWorkbook();
        try {
            int existing = workbook.getNumCellStyles();
            ExcelWriteException error = assertThrows(ExcelWriteException.class,
                    () -> ExcelWriteSupport.createDerivedStyle(workbook, existing));
            assertTrue(error.getMessage().contains("Too many distinct cell styles"), error.getMessage());
            assertTrue(error.getMessage().contains("cellColor"), error.getMessage());
            assertEquals(existing, workbook.getNumCellStyles());
        } finally {
            workbook.dispose();
        }
    }
}
