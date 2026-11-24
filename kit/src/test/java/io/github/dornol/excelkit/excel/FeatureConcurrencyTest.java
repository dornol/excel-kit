package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Workbooks written at the same time by different threads (a web server does this) must not see each other: every
 * file holds exactly its own rows, subtotals, cached results, chart and message.
 */
class FeatureConcurrencyTest {
    record Line(String dept, String name, int amount) {}

    private static byte[] write(int seed, boolean viaWorkbook) throws Exception {
        List<Line> lines = new ArrayList<>();
        int groups = 2 + seed % 5;
        for (int g = 0; g < groups; g++) for (int i = 0; i <= g + seed % 3; i++) lines.add(new Line("D" + seed + "-" + g, "n" + i, seed * 100 + g * 10 + i));
        var out = new ByteArrayOutputStream();
        if (viaWorkbook) {
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Line>sheet("Data " + seed).column("Dept", Line::dept).column("Name", Line::name).column("Amount", Line::amount)
                        .groupSubtotal(Line::dept, s -> s.label("Name", k -> k + " total").keyColumn("Dept").sum("Amount"))
                        .summary(s -> s.label("Total").sum("Amount")).stripedRows().write(lines);
                workbook.<Line>sheet("Empty " + seed).column("Name", Line::name).emptyMessage("empty " + seed).write(List.of());
                workbook.finish().writeTo(out);
            }
        } else {
            ExcelWriter.<Line>create().maxRows(4).sheetName("Data " + seed)
                    .column("Dept", Line::dept).column("Name", Line::name).column("Amount", Line::amount)
                    .groupSubtotal(Line::dept, s -> s.label("Name", k -> k + " total").keyColumn("Dept").sum("Amount"))
                    .summary(s -> s.label("Total").sum("Amount")).write(lines).writeTo(out);
        }
        return out.toByteArray();
    }

    private static void verify(int seed, boolean viaWorkbook, byte[] file) throws Exception {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            int subtotals = 0;
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                var sheet = wb.getSheetAt(s);
                for (var row : sheet) {
                    var first = row.getCell(0);
                    if (first != null && first.getCellType() == CellType.STRING && first.getStringCellValue().length() > 1 && Character.isDigit(first.getStringCellValue().charAt(1))) {
                        assertTrue(first.getStringCellValue().startsWith("D" + seed + "-"), "seed " + seed + " holds only its own rows: " + first);
                    }
                    var formula = row.getCell(2);
                    if (formula != null && formula.getCellType() == CellType.FORMULA) {
                        subtotals++;
                        assertEquals(evaluator.evaluate(formula).getNumberValue(), formula.getNumericCellValue(), 1e-9,
                                "seed " + seed + " cached result of " + formula.getCellFormula());
                    }
                }
            }
            assertTrue(subtotals >= 2, "seed " + seed);
            if (viaWorkbook) assertEquals("empty " + seed, wb.getSheetAt(1).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void manyThreadsWriteTheirOwnWorkbooks() throws Exception {
        int threads = 12;
        ExecutorService pool = Executors.newFixedThreadPool(threads);
        CountDownLatch start = new CountDownLatch(1);
        List<Future<?>> futures = Collections.synchronizedList(new ArrayList<>());
        try {
            for (int t = 0; t < threads; t++) {
                final int seed = t + 1;
                futures.add(pool.submit(() -> {
                    start.await();
                    for (int i = 0; i < 25; i++) {
                        boolean viaWorkbook = i % 2 == 0;
                        verify(seed, viaWorkbook, write(seed, viaWorkbook));
                    }
                    return null;
                }));
            }
            start.countDown();
            for (Future<?> f : futures) f.get();
        } finally {
            pool.shutdownNow();
        }
    }
}
