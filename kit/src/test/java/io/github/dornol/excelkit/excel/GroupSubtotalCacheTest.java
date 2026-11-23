package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;
import java.util.Random;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * A formula carries its result next to it, so that a viewer that does not calculate (Protected View, a preview, another
 * program) shows it. The result is worked out from the cells as they are written; where it cannot be known for sure
 * the formula is written without one. Excel still recalculates when it opens the file.
 */
class GroupSubtotalCacheTest {
    record Row(String dept, String name, int amount, double price) {}

    private static final List<Row> ROWS = List.of(
            new Row("A", "x", 10, 1.5), new Row("A", "y", 20, 2.5), new Row("A", "z", 60, 3.0), new Row("B", "w", 4, 9.25));

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static ExcelWriter<Row> base() {
        return ExcelWriter.<Row>create()
                .column("Dept", Row::dept).column("Name", Row::name)
                .column("Amount", Row::amount).column("Price", Row::price);
    }

    private static double cached(XSSFWorkbook wb, int sheet, int row, int col) {
        var cell = wb.getSheetAt(sheet).getRow(row).getCell(col);
        assertEquals(CellType.FORMULA, cell.getCellType());
        assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType(), "cached type at " + row + "," + col);
        return cell.getNumericCellValue();
    }

    @Nested
    class Subtotals {
        @Test
        void everyFunctionCarriesItsResult() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t")
                    .countA("Name").sum("Amount").average("Price"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(3, cached(wb, 0, 4, 1));
                assertEquals(90, cached(wb, 0, 4, 2));
                assertEquals(2.3333333333333335, cached(wb, 0, 4, 3), 1e-12);
                assertEquals(1, cached(wb, 0, 6, 1));
                assertEquals(9.25, cached(wb, 0, 6, 3));
            }
        }

        @Test
        void countMinAndMax() throws IOException {
            var w2 = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").count("Amount").min("Price"));
            try (var wb = open(w2.write(ROWS))) {
                assertEquals(3, cached(wb, 0, 4, 2));
                assertEquals(1.5, cached(wb, 0, 4, 3));
            }
            var w3 = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").max("Price"));
            try (var wb = open(w3.write(ROWS))) {
                assertEquals(3.0, cached(wb, 0, 4, 3));
                assertEquals(9.25, cached(wb, 0, 6, 3));
            }
        }

        @Test
        void aSubtotalOverTextCarriesZeroForNumberFunctionsAndTheCountForCountA() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Name"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(0, cached(wb, 0, 4, 1), "SUBTOTAL(9) over text is 0");
            }
            var count = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").count("Name"));
            try (var wb = open(count.write(ROWS))) {
                assertEquals(0, cached(wb, 0, 4, 1));
            }
        }

        @Test
        void aNumberStoredAsTextIsNotCountedAsAnExcelNumber() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .column("Code", r -> String.valueOf(r.amount()))
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Code"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(0, cached(wb, 0, 4, 1));
            }
        }

        @Test
        void aDateColumnIsNumbersToExcelSoItCountsAndHasAMaximum() throws IOException {
            record Dated(String key, LocalDate date, LocalDate other) {}
            var w = ExcelWriter.<Dated>create().column("Key", Dated::key)
                    .column("Date", Dated::date, c -> c.type(ExcelDataType.DATE))
                    .column("Other", Dated::other, c -> c.type(ExcelDataType.DATE))
                    .groupSubtotal(Dated::key, s -> s.label("Key", k -> "t").count("Date").max("Other"));
            var rows = List.of(new Dated("A", LocalDate.of(2026, 1, 1), LocalDate.of(2026, 1, 1)),
                    new Dated("A", LocalDate.of(2026, 1, 3), LocalDate.of(2026, 1, 3)));
            try (var wb = open(w.write(rows))) {
                assertEquals(2, cached(wb, 0, 3, 1));
                assertEquals(46025, cached(wb, 0, 3, 2), "the serial number of 2026-01-03");
            }
        }

        @Test
        void bigDecimalAndPercentColumnsCarryTheStoredNumber() throws IOException {
            record Money(String key, BigDecimal amount, double rate) {}
            var w = ExcelWriter.<Money>create().column("Key", Money::key)
                    .column("Amount", Money::amount, c -> c.type(ExcelDataType.BIG_DECIMAL_TO_DOUBLE))
                    .column("Rate", Money::rate, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                    .groupSubtotal(Money::key, s -> s.label("Key", k -> "t").sum("Amount").sum("Rate"));
            try (var wb = open(w.write(List.of(new Money("A", new BigDecimal("1.10"), 0.25), new Money("A", new BigDecimal("2.20"), 0.5))))) {
                assertEquals(3.3, cached(wb, 0, 3, 1), 1e-9);
                assertEquals(0.75, cached(wb, 0, 3, 2), 1e-9);
            }
        }

        @Test
        void aColumnOfFormulasIsWrittenWithoutAResultExceptForCountA() throws IOException {
            var sum = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept)
                    .column("Calc", (row, cursor) -> "1+1", c -> c.type(ExcelDataType.FORMULA))
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Calc"));
            try (var wb = open(sum.write(ROWS))) {
                var cell = wb.getSheetAt(0).getRow(4).getCell(1);
                assertEquals(CellType.FORMULA, cell.getCellType());
                assertEquals(CellType.STRING, cell.getCachedFormulaResultType(), "no result: the formula depends on formulas");
            }
            var countA = ExcelWriter.<Row>create()
                    .column("Dept", Row::dept)
                    .column("Calc", (row, cursor) -> "1+1", c -> c.type(ExcelDataType.FORMULA))
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").countA("Calc"));
            try (var wb = open(countA.write(ROWS))) {
                assertEquals(3, cached(wb, 0, 4, 1), "COUNTA counts formula cells whatever they give");
            }
        }

        @Test
        void anAverageOverNoNumbersHasNoResultLikeExcelsDivisionByZero() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").average("Name"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(CellType.STRING, wb.getSheetAt(0).getRow(4).getCell(1).getCachedFormulaResultType());
            }
        }

        @Test
        void aMinAndMaxOverNoNumbersAreZeroLikeExcel() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").min("Name"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(0, cached(wb, 0, 4, 1));
            }
        }

        @Test
        void emptyTextCellsLeaveCountAWithoutAResult() throws IOException {
            var w = ExcelWriter.<Row>create().column("Dept", Row::dept)
                    .column("Note", r -> r.name().equals("y") ? "" : "n")
                    .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").countA("Note"));
            try (var wb = open(w.write(ROWS))) {
                // whether Excel's COUNTA counts an empty text cell is left to Excel: no result for the group that has one
                assertEquals(CellType.STRING, wb.getSheetAt(0).getRow(4).getCell(1).getCachedFormulaResultType());
                assertEquals(1, cached(wb, 0, 6, 1), "the group without one has its result");
            }
        }

        @Test
        void eachGroupStartsFromNothingAndASheetBreakClosesTheGroup() throws IOException {
            var w = base().maxRows(3).groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(90, cached(wb, 0, 4, 2));
                assertEquals(4, cached(wb, 1, 2, 2), "B alone on the second sheet, not 94");
            }
        }

    }

    @Nested
    class Summaries {
        @Test
        void aPlainSummaryCarriesItsResults() throws IOException {
            var w = base().summary(s -> s.label("Dept", "All").sum("Amount").average("Price").count("Amount").countA("Name").min("Price").max("Price"));
            try (var wb = open(w.write(ROWS))) {
                // one summary row per operation, in the order they were added
                assertEquals(94, cached(wb, 0, 5, 2));
                assertEquals(4.0625, cached(wb, 0, 6, 3), 1e-12);
                assertEquals(4, cached(wb, 0, 7, 2));
                assertEquals(4, cached(wb, 0, 8, 1));
                assertEquals(1.5, cached(wb, 0, 9, 3));
                assertEquals(9.25, cached(wb, 0, 10, 3));
            }
        }

        @Test
        void underSubtotalsTheSummaryCountsTheDataRowsOnly() throws IOException {
            var w = base().groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount"))
                    .summary(s -> s.label("Dept", "All").sum("Amount"));
            try (var wb = open(w.write(ROWS))) {
                assertEquals(90, cached(wb, 0, 4, 2));
                assertEquals(94, cached(wb, 0, 7, 2), "not 94 + 90 + 4");
            }
        }

        @Test
        void everyRolledOverSheetCarriesTheSummaryOfItsOwnRows() throws IOException {
            var rows = IntStream.range(0, 10).mapToObj(i -> new Row("A", "n" + i, i + 1, 1)).toList();
            var w = base().maxRows(4).summary(s -> s.label("Dept", "All").sum("Amount"));
            try (var wb = open(w.write(rows))) {
                assertEquals(3, wb.getNumberOfSheets());
                assertEquals(10, cached(wb, 0, 5, 2), "1+2+3+4");
                assertEquals(26, cached(wb, 1, 5, 2), "5+6+7+8");
                assertEquals(19, cached(wb, 2, 3, 2), "9+10");
            }
        }

        @Test
        void aSummaryOverAColumnTheWriterDoesNotFillHasNoTroubleWithAnEmptySheet() throws IOException {
            var w = base().summary(s -> s.label("Dept", "All").sum("Amount"));
            try (var wb = open(w.write(List.<Row>of()))) {
                assertEquals(CellType.BLANK, wb.getSheetAt(0).getRow(1).getCell(2) == null ? CellType.BLANK : wb.getSheetAt(0).getRow(1).getCell(2).getCellType());
            }
        }
    }

    @Nested
    class Resilience {
        @Test
        void theResultsHoldWithAWindowOfOneRow() throws IOException {
            var rows = IntStream.range(0, 200).mapToObj(i -> new Row("K" + i / 9, "n" + i, i, i)).toList();
            var w = ExcelWriter.<Row>create(o -> o.rowAccessWindowSize(1))
                    .column("Dept", Row::dept).column("Name", Row::name).column("Amount", Row::amount)
                    .maxRows(70)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> "t").sum("Amount"))
                    .summary(s -> s.label("Name", "All").sum("Amount"));
            try (var wb = open(w.write(rows))) {
                var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                for (int sheet = 0; sheet < wb.getNumberOfSheets(); sheet++) {
                    for (var row : wb.getSheetAt(sheet)) {
                        var cell = row.getCell(2);
                        if (cell != null && cell.getCellType() == CellType.FORMULA) {
                            assertEquals(evaluator.evaluate(cell).getNumberValue(), cell.getNumericCellValue(), 1e-9,
                                    "sheet " + sheet + " row " + row.getRowNum());
                        }
                    }
                }
            }
        }

        @Test
        void aWorkbookSheetCarriesTheResultsToo() throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("S").column("Dept", Row::dept).column("Amount", Row::amount)
                        .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount"))
                        .summary(s -> s.label("Dept", "All").sum("Amount"))
                        .write(ROWS);
                try (var wb = open(workbook.finish())) {
                    assertEquals(90, cached(wb, 0, 4, 1));
                    assertEquals(94, cached(wb, 0, 7, 1));
                }
            }
        }

        @Test
        void randomDataAgreesWithTheEvaluatorForEveryFunction() throws IOException {
            for (long seed = 0; seed < 60; seed++) {
                Random random = new Random(seed);
                int count = 1 + random.nextInt(40);
                var rows = IntStream.range(0, count)
                        .mapToObj(i -> new Row("K" + random.nextInt(4), random.nextBoolean() ? "n" : "m", random.nextInt(200) - 100,
                                random.nextInt(1000) / 8.0))
                        .sorted(java.util.Comparator.comparing(Row::dept)).toList();
                var w = ExcelWriter.<Row>create()
                        .column("Dept", Row::dept).column("Name", Row::name)
                        .column("Amount", Row::amount).column("Price", Row::price)
                        .column("Count", r -> r.amount(), c -> c.type(ExcelDataType.INTEGER))
                        .column("Min", Row::price).column("Max", Row::price)
                        .groupSubtotal(Row::dept, s -> s.label("Dept", k -> "t").sum("Amount").average("Price")
                                .count("Count").countA("Name").min("Min").max("Max"))
                        .summary(s -> s.label("Dept", "All").sum("Amount").average("Price").count("Count").countA("Name").min("Min").max("Max"));
                try (var wb = open(w.write(rows))) {
                    var evaluator = wb.getCreationHelper().createFormulaEvaluator();
                    for (var row : wb.getSheetAt(0)) {
                        for (int col = 1; col < 8; col++) {
                            var cell = row.getCell(col);
                            if (cell == null || cell.getCellType() != CellType.FORMULA) continue;
                            var result = evaluator.evaluate(cell);
                            String at = "seed " + seed + " row " + row.getRowNum() + " col " + col + " " + cell.getCellFormula();
                            if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                assertEquals(result.getNumberValue(), cell.getNumericCellValue(), 1e-9, at);
                            } else {
                                // no result written: only where Excel's own result is not a number (an error)
                                assertNotEquals(CellType.NUMERIC, result.getCellType(), at + " (a number the writer could have known)");
                            }
                        }
                    }
                }
            }
        }
    }
}
