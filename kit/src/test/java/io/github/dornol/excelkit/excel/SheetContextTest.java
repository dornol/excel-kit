package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link SheetContext}.
 */
class SheetContextTest {

    private SXSSFWorkbook wb;
    private SXSSFSheet sheet;

    @BeforeEach
    void setUp() {
        wb = new SXSSFWorkbook(10);
        sheet = wb.createSheet("TestSheet");
    }

    @AfterEach
    void tearDown() throws Exception {
        wb.close();
    }

    private SheetContext ctx(List<ExcelColumn<String>> columns, int currentRow, int headerRowIndex) {
        return new SheetContext(sheet, wb, currentRow, columns, headerRowIndex);
    }

    private SheetContext ctx(List<ExcelColumn<String>> columns, int currentRow) {
        return new SheetContext(sheet, wb, currentRow, columns);
    }

    private List<ExcelColumn<String>> sampleColumns() {
        return List.of(
                new ExcelColumn<>("Name", (r, c) -> r, null, ExcelDataType.STRING.getSetter(),
                        0, 0, false, null, null, null, 0, null, null, null, false, null),
                new ExcelColumn<>("Age", (r, c) -> r, null, ExcelDataType.INTEGER.getSetter(),
                        0, 0, false, null, null, null, 0, null, null, null, false, null),
                new ExcelColumn<>("Score", (r, c) -> r, null, ExcelDataType.DOUBLE.getSetter(),
                        0, 0, false, null, null, null, 0, null, null, null, false, null)
        );
    }

    // ============================================================
    // Basic accessors
    // ============================================================
    @Nested
    class AccessorTests {

        @Test
        void getSheet_returnsCorrectSheet() {
            SheetContext c = ctx(sampleColumns(), 5);
            assertSame(sheet, c.getSheet());
        }

        @Test
        void getWorkbook_returnsCorrectWorkbook() {
            SheetContext c = ctx(sampleColumns(), 5);
            assertSame(wb, c.getWorkbook());
        }

        @Test
        void getCurrentRow_returnsCorrectValue() {
            SheetContext c = ctx(sampleColumns(), 42);
            assertEquals(42, c.getCurrentRow());
        }

        @Test
        void getColumnCount_returnsCorrectCount() {
            SheetContext c = ctx(sampleColumns(), 0);
            assertEquals(3, c.getColumnCount());
        }

        @Test
        void getHeaderRowIndex_defaultIsZero() {
            SheetContext c = ctx(sampleColumns(), 5);
            assertEquals(0, c.getHeaderRowIndex());
        }

        @Test
        void getHeaderRowIndex_customValue() {
            SheetContext c = ctx(sampleColumns(), 5, 3);
            assertEquals(3, c.getHeaderRowIndex());
        }

        @Test
        void getColumnNames_returnsCorrectNames() {
            SheetContext c = ctx(sampleColumns(), 0);
            assertEquals(List.of("Name", "Age", "Score"), c.getColumnNames());
        }

        @Test
        void getColumnNames_isUnmodifiable() {
            SheetContext c = ctx(sampleColumns(), 0);
            assertThrows(UnsupportedOperationException.class, () -> c.getColumnNames().add("Extra"));
        }
    }

    // ============================================================
    // columnLetter
    // ============================================================
    @Nested
    class ColumnLetterTests {

        @Test
        void singleLetters() {
            assertEquals("A", SheetContext.columnLetter(0));
            assertEquals("B", SheetContext.columnLetter(1));
            assertEquals("Z", SheetContext.columnLetter(25));
        }

        @Test
        void doubleLetters() {
            assertEquals("AA", SheetContext.columnLetter(26));
            assertEquals("AB", SheetContext.columnLetter(27));
            assertEquals("AZ", SheetContext.columnLetter(51));
            assertEquals("BA", SheetContext.columnLetter(52));
        }

        @Test
        void tripleLetters() {
            // 26 + 26*26 = 702 → "AAA"
            assertEquals("AAA", SheetContext.columnLetter(702));
        }
    }

    // ============================================================
    // namedRange
    // ============================================================
    @Nested
    class NamedRangeTests {

        @Test
        void namedRange_withFormula_createsRange() {
            SheetContext c = ctx(sampleColumns(), 10);
            SheetContext result = c.namedRange("MyRange", "TestSheet!$A$1:$A$10");
            assertSame(c, result);

            var name = wb.getName("MyRange");
            assertNotNull(name);
            assertEquals("TestSheet!$A$1:$A$10", name.getRefersToFormula());
        }

        @Test
        void namedRange_withColAndRows_createsRange() {
            SheetContext c = ctx(sampleColumns(), 10);
            SheetContext result = c.namedRange("Categories", 0, 0, 9);
            assertSame(c, result);

            var name = wb.getName("Categories");
            assertNotNull(name);
            assertEquals("'TestSheet'!$A$1:$A$10", name.getRefersToFormula());
        }
    }

    // ============================================================
    // mergeCells
    // ============================================================
    @Nested
    class MergeCellTests {

        @Test
        void mergeCells_byIndices_returnsSelf() {
            // Create rows first
            sheet.createRow(0);
            sheet.createRow(1);
            SheetContext c = ctx(sampleColumns(), 2);
            SheetContext result = c.mergeCells(0, 1, 0, 2);
            assertSame(c, result);
            assertEquals(1, sheet.getNumMergedRegions());
        }

        @Test
        void mergeCells_byRange_returnsSelf() {
            sheet.createRow(0);
            SheetContext c = ctx(sampleColumns(), 1);
            SheetContext result = c.mergeCells("A1:C1");
            assertSame(c, result);
            assertEquals(1, sheet.getNumMergedRegions());
        }
    }

    // ============================================================
    // groupRows
    // ============================================================
    @Nested
    class GroupRowTests {

        @Test
        void groupRows_returnsSelf() {
            sheet.createRow(0);
            sheet.createRow(1);
            sheet.createRow(2);
            SheetContext c = ctx(sampleColumns(), 3);
            SheetContext result = c.groupRows(0, 2);
            assertSame(c, result);
        }

        @Test
        void groupRows_collapsed_returnsSelf() {
            sheet.createRow(0);
            sheet.createRow(1);
            SheetContext c = ctx(sampleColumns(), 2);
            SheetContext result = c.groupRows(0, 1, true);
            assertSame(c, result);
        }
    }
}
