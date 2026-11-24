package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelWorkbookInspector;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.read.ExcelSheetInfo;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** An Excel sheet holds 1,048,576 rows; the writers roll over before that and say so when they cannot. */
class SheetRowLimitTest {

    @TempDir
    Path tempDir;

    @Test
    void aWorkbookSheetRollsOverByDefaultInsteadOfHittingTheExcelLimit() throws IOException {
        Path file = tempDir.resolve("big.xlsx");
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<Integer>sheet("Big").column("v", i -> i)
                    .write(IntStream.range(0, 1_000_005).boxed());
            try (OutputStream out = Files.newOutputStream(file)) {
                workbook.finish().writeTo(out);
            }
        }
        List<ExcelSheetInfo> sheets = ExcelWorkbookInspector.create().sheetNames(file);
        assertEquals(2, sheets.size());
        assertEquals("Big", sheets.get(0).name());
        assertEquals("Big (2)", sheets.get(1).name());
    }

    @Test
    void maxRowsMustFitBelowTheHeaderRow() {
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().maxRows(1_048_576));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().maxRows(0));
        assertDoesNotThrow(() -> ExcelWriter.<String>create().maxRows(1_048_575));
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            assertThrows(IllegalArgumentException.class, () -> workbook.<String>sheet("S").maxRows(2_000_000));
            assertDoesNotThrow(() -> workbook.<String>sheet("T").maxRows(1_048_575));
        }
    }

    @Test
    void aSheetThatRunsOutOfRowsSaysWhyAndHowToFixIt() {
        // maxRows leaves no room for the five rows beforeHeader writes plus the header.
        ExcelWriteException error = assertThrows(ExcelWriteException.class, () ->
                ExcelWriter.<Integer>create().maxRows(1_048_575).beforeHeader(context -> 5)
                        .column("v", i -> i)
                        .write(IntStream.range(0, 1_048_580).boxed())
                        .writeTo(OutputStream.nullOutputStream()));
        assertTrue(error.getMessage().contains("is full"), error.getMessage());
        assertTrue(error.getMessage().contains("maxRows"), error.getMessage());
    }

    @Test
    void aListInATemplateCannotRollOverSoItDoesNotSuggestMaxRows() throws Exception {
        var template = new java.io.ByteArrayOutputStream();
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            wb.createSheet("S");
            wb.write(template);
        }
        try (var writer = new ExcelTemplateWriter(new java.io.ByteArrayInputStream(template.toByteArray()))) {
            var list = writer.<Integer>list(0, 0).column("v", i -> i);
            ExcelWriteException error = assertThrows(ExcelWriteException.class,
                    () -> list.write(IntStream.range(0, 1_048_580).boxed()));
            assertTrue(error.getMessage().contains("is full"), error.getMessage());
            assertTrue(error.getMessage().contains("template"), error.getMessage());
            assertFalse(error.getMessage().contains("maxRows"), "TemplateListWriter has no maxRows: " + error.getMessage());
        }
    }
}
