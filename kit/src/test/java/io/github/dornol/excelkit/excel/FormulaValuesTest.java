package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.testsupport.DisposingSxssfWorkbook;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.Date;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/** What the formula cache works out from the written cells, op by op and cell type by cell type. */
class FormulaValuesTest {
    private SXSSFWorkbook wb;
    private SXSSFSheet sheet;
    private int next;

    @BeforeEach
    void setUp() {
        wb = new DisposingSxssfWorkbook();
        sheet = wb.createSheet("s");
        next = 0;
    }

    @AfterEach
    void tearDown() throws IOException {
        wb.close();
    }

    private SXSSFCell cell() {
        return sheet.createRow(next++).createCell(0);
    }

    private SXSSFCell number(double v) {
        SXSSFCell c = cell();
        c.setCellValue(v);
        return c;
    }

    private SXSSFCell text(String v) {
        SXSSFCell c = cell();
        c.setCellValue(v);
        return c;
    }

    private static FormulaValues.Accumulator acc(SXSSFCell... cells) {
        var a = new FormulaValues.Accumulator();
        for (SXSSFCell c : cells) a.add(c);
        return a;
    }

    @Nested
    class Accumulating {
        @Test
        void numbersGiveEveryFunction() {
            var a = acc(number(2), number(-4), number(8));
            assertEquals(6.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(3.0, a.value(ExcelSummary.Op.COUNT));
            assertEquals(3.0, a.value(ExcelSummary.Op.COUNTA));
            assertEquals(2.0, a.value(ExcelSummary.Op.AVERAGE));
            assertEquals(-4.0, a.value(ExcelSummary.Op.MIN));
            assertEquals(8.0, a.value(ExcelSummary.Op.MAX));
        }

        @Test
        void aSingleNumberIsItsOwnMinAndMax() {
            var a = acc(number(-7.5));
            assertEquals(-7.5, a.value(ExcelSummary.Op.MIN));
            assertEquals(-7.5, a.value(ExcelSummary.Op.MAX));
            assertEquals(-7.5, a.value(ExcelSummary.Op.AVERAGE));
        }

        @Test
        void nothingAtAllIsZeroExceptTheAverage() {
            var a = new FormulaValues.Accumulator();
            assertEquals(0.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(0.0, a.value(ExcelSummary.Op.COUNT));
            assertEquals(0.0, a.value(ExcelSummary.Op.COUNTA));
            assertEquals(0.0, a.value(ExcelSummary.Op.MIN));
            assertEquals(0.0, a.value(ExcelSummary.Op.MAX));
            assertNull(a.value(ExcelSummary.Op.AVERAGE), "Excel shows #DIV/0!, so no result is cached");
        }

        @Test
        void textCountsForCountAOnly() {
            var a = acc(text("x"), number(3), text("y"));
            assertEquals(3.0, a.value(ExcelSummary.Op.COUNTA));
            assertEquals(1.0, a.value(ExcelSummary.Op.COUNT));
            assertEquals(3.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(3.0, a.value(ExcelSummary.Op.AVERAGE));
        }

        @Test
        void onlyTextMeansNoAverageButZeroForTheRest() {
            var a = acc(text("x"), text("y"));
            assertNull(a.value(ExcelSummary.Op.AVERAGE));
            assertEquals(0.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(0.0, a.value(ExcelSummary.Op.MIN));
            assertEquals(0.0, a.value(ExcelSummary.Op.MAX));
            assertEquals(2.0, a.value(ExcelSummary.Op.COUNTA));
        }

        @Test
        void anEmptyStringLeavesCountAUndecidedAndTheNumbersAlone() {
            var a = acc(number(5), text(""), number(7));
            assertNull(a.value(ExcelSummary.Op.COUNTA), "whether COUNTA counts an empty text cell is not settled");
            assertEquals(12.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(2.0, a.value(ExcelSummary.Op.COUNT));
            assertEquals(6.0, a.value(ExcelSummary.Op.AVERAGE));
        }

        @Test
        void booleansCountForCountAOnly() {
            SXSSFCell b = cell();
            b.setCellValue(true);
            var a = acc(b, number(4));
            assertEquals(2.0, a.value(ExcelSummary.Op.COUNTA));
            assertEquals(1.0, a.value(ExcelSummary.Op.COUNT));
            assertEquals(4.0, a.value(ExcelSummary.Op.SUM));
        }

        @Test
        void blankCellsAreNotCounted() {
            SXSSFCell blank = cell();
            blank.setBlank();
            var a = acc(blank, number(1));
            assertEquals(1.0, a.value(ExcelSummary.Op.COUNTA));
            assertEquals(1.0, a.value(ExcelSummary.Op.COUNT));
        }

        @Test
        void aDateIsANumberInTheFile() {
            SXSSFCell d = cell();
            d.setCellValue(new Date(0));
            var a = acc(d);
            assertEquals(1.0, a.value(ExcelSummary.Op.COUNT));
            assertNotNull(a.value(ExcelSummary.Op.SUM));
        }
    }

    @Nested
    class Uncertain {
        private SXSSFCell formula() {
            SXSSFCell c = cell();
            c.setCellFormula("1+1");
            return c;
        }

        private SXSSFCell error() {
            SXSSFCell c = cell();
            c.setCellErrorValue(FormulaError.NA.getCode());
            return c;
        }

        @Test
        void aFormulaCellMakesEveryNumberFunctionUnknownButCountAIsKnown() {
            var a = acc(number(1), formula());
            for (var op : List.of(ExcelSummary.Op.SUM, ExcelSummary.Op.COUNT, ExcelSummary.Op.AVERAGE,
                    ExcelSummary.Op.MIN, ExcelSummary.Op.MAX)) {
                assertNull(a.value(op), op.name());
            }
            assertEquals(2.0, a.value(ExcelSummary.Op.COUNTA));
        }

        @Test
        void anErrorCellDoesTheSame() {
            var a = acc(number(1), error());
            for (var op : List.of(ExcelSummary.Op.SUM, ExcelSummary.Op.COUNT, ExcelSummary.Op.AVERAGE,
                    ExcelSummary.Op.MIN, ExcelSummary.Op.MAX)) {
                assertNull(a.value(op), op.name());
            }
            assertEquals(2.0, a.value(ExcelSummary.Op.COUNTA));
        }

        @Test
        void resetForgetsEverythingIncludingTheUncertainty() {
            var a = acc(number(1), formula(), text(""));
            a.reset();
            assertEquals(0.0, a.value(ExcelSummary.Op.SUM));
            assertEquals(0.0, a.value(ExcelSummary.Op.COUNTA), "the empty text is forgotten too");
            assertEquals(0.0, a.value(ExcelSummary.Op.MIN));
            a.add(number(9));
            assertEquals(9.0, a.value(ExcelSummary.Op.MAX));
            assertEquals(9.0, a.value(ExcelSummary.Op.MIN), "min and max start over, not from the old values");
        }

        @Test
        void resetRestoresTheExtremes() {
            var a = acc(number(100), number(-100));
            a.reset();
            a.add(number(5));
            assertEquals(5.0, a.value(ExcelSummary.Op.MIN));
            assertEquals(5.0, a.value(ExcelSummary.Op.MAX));
        }
    }

    @Nested
    class NotFinite {
        @Test
        void aSumThatOverflowsIsNotCached() {
            var a = acc(number(Double.MAX_VALUE), number(Double.MAX_VALUE));
            assertNull(a.value(ExcelSummary.Op.SUM));
            assertNull(a.value(ExcelSummary.Op.AVERAGE));
            assertEquals(Double.MAX_VALUE, a.value(ExcelSummary.Op.MAX));
        }
    }
}
