package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Random;

import static org.junit.jupiter.api.Assertions.*;

/**
 * A workbook with subtotals read back by the library's own reader, as a file that was saved by this library and then
 * sent on: the subtotal rows carry their results (the reader reads what a file stores, it calculates nothing).
 */
class GroupSubtotalRoundTripTest {
    record Line(String dept, String item, int amount) {}

    private static ExcelWriter<Line> writer() {
        return ExcelWriter.<Line>create()
                .column("Dept", Line::dept)
                .column("Item", Line::item)
                .column("Amount", Line::amount, c -> c.type(ExcelDataType.INTEGER));
    }

    private static byte[] bytes(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return out.toByteArray();
    }

    /** Every sheet of the workbook, in order: the reader reads one sheet at a time. */
    private static List<ReadResult<Map<String, String>>> read(byte[] file) throws IOException {
        int sheets;
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new ByteArrayInputStream(file))) {
            sheets = wb.getNumberOfSheets();
        }
        List<ReadResult<Map<String, String>>> rows = new ArrayList<>();
        for (int i = 0; i < sheets; i++) {
            ExcelReader.forMap().sheetIndex(i).read(new ByteArrayInputStream(file), r -> rows.add(r));
        }
        return rows;
    }

    /** A number as the reader gives it: the text of the cell, with the thousands separator of the format. */
    private static double number(String text) {
        return Double.parseDouble(text.replace(",", ""));
    }

    @Nested
    class Reading {
        @Test
        void subtotalRowsComeBackWithTheirLabelAndTheirResult() throws IOException {
            var lines = List.of(new Line("A", "x", 10), new Line("A", "y", 20), new Line("B", "z", 5));
            var file = bytes(writer().groupSubtotal(Line::dept, s -> s.label("Item", k -> k + " total").sum("Amount")).write(lines));
            var rows = read(file);
            assertEquals(5, rows.size(), "three data rows and two subtotal rows");
            assertTrue(rows.stream().allMatch(r -> r.success()));
            assertEquals("A total", rows.get(2).data().get("Item"));
            assertEquals("30", rows.get(2).data().get("Amount"), "the cached result is read");
            assertEquals("B total", rows.get(4).data().get("Item"));
            assertEquals("5", rows.get(4).data().get("Amount"));
            assertEquals(List.of(2L, 3L, 4L, 5L, 6L), rows.stream().map(r -> r.fileRowNum()).toList());
        }

        @Test
        void aSummaryReadsAsTheGrandTotalWithoutCountingSubtotalsTwice() throws IOException {
            var lines = List.of(new Line("A", "x", 10), new Line("A", "y", 20), new Line("B", "z", 5));
            var file = bytes(writer().groupSubtotal(Line::dept, s -> s.label("Item", k -> "sub").sum("Amount"))
                    .summary(s -> s.sum("Amount")).write(lines));
            var rows = read(file);
            var last = rows.get(rows.size() - 1).data();
            assertEquals("35", last.get("Amount"), "10 + 20 + 5, not 70");
        }

        @Test
        void withoutTheCacheTheFileReadsEmptyWhichIsWhyTheCacheIsWritten() throws IOException {
            // Count of text cells has no number result for AVERAGE: the cell is left without one and reads as empty
            var lines = List.of(new Line("A", "x", 1));
            var file = bytes(ExcelWriter.<Line>create().column("Dept", Line::dept).column("Item", Line::item)
                    .column("Amount", Line::amount).groupSubtotal(Line::dept, s -> s.label("Dept", k -> "t").average("Item")).write(lines));
            var rows = read(file);
            assertEquals(2, rows.size());
            assertTrue(rows.get(1).data().get("Item") == null || rows.get(1).data().get("Item").isEmpty());
        }
    }

    @Nested
    class RolloverAndScale {
        @Test
        void everySheetOfARolledOverWorkbookReadsBackWithItsOwnSubtotals() throws IOException {
            List<Line> lines = new ArrayList<>();
            for (int d = 0; d < 6; d++) for (int i = 0; i < 4; i++) lines.add(new Line("D" + d, "i" + i, d * 10 + i));
            var file = bytes(writer().maxRows(8).groupSubtotal(Line::dept, s -> s.label("Item", k -> k + "!").sum("Amount")).write(lines));
            var rows = read(file);
            long subtotals = rows.stream().filter(r -> r.data().get("Item") != null && r.data().get("Item").endsWith("!")).count();
            assertEquals(6, subtotals, "one per department, none lost at the rollover");
            double total = 0;
            for (var r : rows) {
                if (r.data().get("Item") != null && r.data().get("Item").endsWith("!")) total += number(r.data().get("Amount"));
            }
            assertEquals(lines.stream().mapToInt(Line::amount).sum(), total, 1e-9);
        }

        @Test
        void randomWorkbooksReadBackWithTheSubtotalsTheModelExpects() throws IOException {
            Random random = new Random(424242);
            for (int round = 0; round < 60; round++) {
                int groups = 1 + random.nextInt(6);
                int maxRows = 3 + random.nextInt(12);
                List<Line> lines = new ArrayList<>();
                Map<String, Integer> expected = new java.util.LinkedHashMap<>();
                for (int g = 0; g < groups; g++) {
                    int size = 1 + random.nextInt(5);
                    for (int i = 0; i < size; i++) {
                        int amount = random.nextInt(2000) - 1000;
                        lines.add(new Line("G" + g, "i" + i, amount));
                        expected.merge("G" + g, amount, Integer::sum);
                    }
                }
                var file = bytes(writer().maxRows(maxRows)
                        .groupSubtotal(Line::dept, s -> s.label("Item", k -> "sub " + k).sum("Amount")).write(lines));
                var rows = read(file);
                Map<String, Integer> actual = new java.util.LinkedHashMap<>();
                int dataRows = 0;
                for (var r : rows) {
                    String item = r.data().get("Item");
                    if (item != null && item.startsWith("sub ")) {
                        actual.merge(item.substring(4), (int) number(r.data().get("Amount")), Integer::sum);
                    } else {
                        dataRows++;
                    }
                }
                String desc = "round " + round + " maxRows=" + maxRows;
                assertEquals(lines.size(), dataRows, desc);
                // a group split by a rollover has one subtotal per part: the parts add up to the group
                assertEquals(expected, actual, desc);
            }
        }
    }
}
