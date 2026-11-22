package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
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

    @Test
    void theXlsxLimitIsNamedAsSuchAndACustomOneByItsSetting() {
        SXSSFWorkbook workbook = new DisposingSxssfWorkbook();
        try {
            ExcelWriteException custom = assertThrows(ExcelWriteException.class,
                    () -> ExcelWriteSupport.createDerivedStyle(workbook, 1));
            assertTrue(custom.getMessage().contains("maxCellStyles(1)"), custom.getMessage());
        } finally {
            workbook.dispose();
        }
    }

    @Test
    void maxCellStylesMustBeBetweenOneAndTheXlsxLimit() {
        ExcelWriter<String> writer = ExcelWriter.create();
        assertThrows(IllegalArgumentException.class, () -> writer.maxCellStyles(0));
        assertThrows(IllegalArgumentException.class, () -> writer.maxCellStyles(64_001));
        assertDoesNotThrow(() -> writer.maxCellStyles(1).maxCellStyles(64_000));
    }

    @Test
    void aColorFunctionWithTooManyDistinctColorsFailsAtTheConfiguredLimit() {
        ExcelWriteException error = assertThrows(ExcelWriteException.class, () ->
                ExcelWriter.<Integer>create().maxCellStyles(30)
                        .column("v", i -> i, column -> column.cellColor((value, row) ->
                                ExcelColor.of(row % 256, (row / 256) % 256, 7)))
                        .write(java.util.stream.IntStream.range(0, 500).boxed())
                        .writeTo(java.io.OutputStream.nullOutputStream()));

        assertTrue(error.getMessage().contains("maxCellStyles(30)"), error.getMessage());
        assertTrue(error.getMessage().contains("palette"), error.getMessage());
    }

    @Test
    void aSmallPaletteStaysWithinTheLimit() {
        assertDoesNotThrow(() ->
                ExcelWriter.<Integer>create().maxCellStyles(30)
                        .column("v", i -> i, column -> column.cellColor((value, row) ->
                                ExcelColor.of(row % 4 * 60, 0, 7)))
                        .write(java.util.stream.IntStream.range(0, 500).boxed())
                        .writeTo(java.io.OutputStream.nullOutputStream()));
    }
}
