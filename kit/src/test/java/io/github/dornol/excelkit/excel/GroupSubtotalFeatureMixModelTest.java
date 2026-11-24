package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Random;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Subtotals with the rest of the sheet features thrown in at random: a preamble, a leading constant column that
 * shifts every column letter, frozen panes, sheet protection, row colors and heights, empty values, a missing group
 * key, four different functions, a summary, after-data content and the empty message, written from a list or a
 * stream, through the writer and through a workbook sheet. Every sheet is compared with a model of what it must hold.
 */
class GroupSubtotalFeatureMixModelTest {
    record R(@org.jspecify.annotations.Nullable String dept, String name, @org.jspecify.annotations.Nullable Integer amount,
             @org.jspecify.annotations.Nullable Double score, @org.jspecify.annotations.Nullable Integer qty,
             @org.jspecify.annotations.Nullable String note) {}

    private static final String[] KEYS = {"A", "B", null, "C"};
    private static final String SUBTOTAL_FILL = "D9D9D9";

    /** What was switched on for one run. */
    private record Mix(int maxRows, boolean preamble, boolean leadingColumn, boolean freeze, boolean protect,
                       boolean rowColor, float rowHeight, boolean summary, boolean afterData, boolean emptyMessage,
                       boolean maxNotMin, boolean keyColumn, boolean stripes, boolean viaStream, boolean viaWorkbook) {
        int off() { return leadingColumn ? 1 : 0; }
        int header() { return preamble ? 2 : 0; }
    }

    @Test
    void randomMixesMatchTheModel() throws IOException {
        for (long seed = 0; seed < 500; seed++) check(seed);
    }

    private void check(long seed) throws IOException {
        Random random = new Random(seed);
        int count = random.nextInt(5) == 0 ? 0 : random.nextInt(45);
        List<R> rows = new ArrayList<>();
        String key = KEYS[0];
        for (int i = 0; i < count; i++) {
            if (random.nextInt(3) == 0) key = KEYS[random.nextInt(KEYS.length)];
            rows.add(new R(key, "r" + i,
                    random.nextInt(4) == 0 ? null : random.nextInt(2000) - 1000,
                    random.nextInt(4) == 0 ? null : Math.round(random.nextDouble() * 10000) / 100.0,
                    random.nextInt(4) == 0 ? null : random.nextInt(50),
                    random.nextInt(3) == 0 ? null : "n" + i));
        }
        boolean viaWorkbook = random.nextInt(3) == 0;
        Mix mix = new Mix(1 + random.nextInt(10), random.nextBoolean(), random.nextBoolean(), random.nextBoolean(),
                random.nextBoolean(), random.nextBoolean(), 12 + random.nextInt(20), random.nextBoolean(),
                random.nextBoolean(), random.nextBoolean(), random.nextBoolean(), random.nextBoolean(),
                random.nextBoolean(), !viaWorkbook && random.nextBoolean(), viaWorkbook);
        String context = "seed " + seed + " " + mix + " rows " + count;

        byte[] file = write(rows, mix);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            if (rows.isEmpty()) {
                assertEmpty(wb, mix, context);
                return;
            }
            int sheets = (rows.size() + mix.maxRows() - 1) / mix.maxRows();
            assertEquals(sheets, wb.getNumberOfSheets(), context + ": sheets");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            for (int s = 0; s < sheets; s++) {
                List<R> chunk = rows.subList(s * mix.maxRows(), Math.min(rows.size(), (s + 1) * mix.maxRows()));
                checkSheet(wb.getSheetAt(s), chunk, mix, evaluator, context + " sheet " + s);
            }
        }
    }

    private static void assertEmpty(XSSFWorkbook wb, Mix mix, String context) {
        assertEquals(1, wb.getNumberOfSheets(), context);
        XSSFSheet sheet = wb.getSheetAt(0);
        assertEquals("Dept", sheet.getRow(mix.header()).getCell(mix.off()).getStringCellValue(), context + ": header");
        for (var row : sheet) {
            for (Cell cell : row) {
                assertFalse(cell.getCellType() == CellType.FORMULA && cell.getCellFormula().startsWith("SUBTOTAL(")
                        && !mix.summary(), context + ": no subtotal row without rows");
            }
        }
        if (mix.emptyMessage()) {
            var line = sheet.getRow(mix.header() + 1);
            assertNotNull(line, context + ": the empty message row");
            assertEquals("nothing here", line.getCell(0).getStringCellValue(), context);
        }
    }

    private void checkSheet(XSSFSheet sheet, List<R> chunk, Mix mix, org.apache.poi.ss.usermodel.FormulaEvaluator evaluator,
                            String at) {
        int off = mix.off();
        int h = mix.header();
        assertEquals("Dept", sheet.getRow(h).getCell(off).getStringCellValue(), at + ": header position");
        if (mix.preamble()) assertEquals("Title", sheet.getRow(0).getCell(0).getStringCellValue(), at + ": preamble");
        if (mix.leadingColumn()) assertEquals("Src", sheet.getRow(h).getCell(0).getStringCellValue(), at);

        int rowIndex = h + 1;
        int sheetAmount = 0;
        int sheetNotes = 0;
        int i = 0;
        int dataIndex = 0;
        while (i < chunk.size()) {
            String groupKey = chunk.get(i).dept();
            int start = rowIndex;
            long sum = 0;
            int numbersAmount = 0;
            double scoreSum = 0;
            int scores = 0;
            Integer min = null;
            Integer max = null;
            int notes = 0;
            while (i < chunk.size() && Objects.equals(chunk.get(i).dept(), groupKey)) {
                R r = chunk.get(i++);
                var row = sheet.getRow(rowIndex);
                String w = at + " row " + rowIndex;
                assertNotNull(row, w);
                assertEquals(r.name(), row.getCell(off + 1).getStringCellValue(), w + ": data order");
                assertEquals(r.dept() == null ? "" : r.dept(), text(row.getCell(off)), w + ": dept");
                if (r.amount() == null) assertTrue(isBlank(row.getCell(off + 2)), w);
                else assertEquals(r.amount(), (int) row.getCell(off + 2).getNumericCellValue(), w);
                assertEquals(mix.rowHeight(), row.getHeightInPoints(), 0.01, w + ": row height");
                assertDataLook(row, off + 2, r, mix, dataIndex, w);
                if (mix.protect()) assertFalse(row.getCell(off + 2).getCellStyle().getLocked(), w + ": data stays editable");
                dataIndex++;
                if (r.amount() != null) { sum += r.amount(); numbersAmount++; }
                if (r.score() != null) { scoreSum += r.score(); scores++; }
                if (r.qty() != null) {
                    min = min == null ? r.qty() : Math.min(min, r.qty());
                    max = max == null ? r.qty() : Math.max(max, r.qty());
                }
                if (r.note() != null) notes++;
                rowIndex++;
            }
            // the subtotal row
            var sub = sheet.getRow(rowIndex);
            String w = at + " subtotal row " + rowIndex;
            assertNotNull(sub, w);
            assertEquals((groupKey == null ? "none" : groupKey) + " total", sub.getCell(off + 1).getStringCellValue(), w + ": label");
            assertEquals(mix.rowHeight(), sub.getHeightInPoints(), 0.01, w + ": row height");
            if (mix.keyColumn()) assertEquals(groupKey == null ? "" : groupKey, text(sub.getCell(off)), w + ": key");
            else assertTrue(isBlank(sub.getCell(off)), w + ": no key");
            String amountLetter = letter(off + 2);
            assertEquals("SUBTOTAL(9," + amountLetter + (start + 1) + ":" + amountLetter + rowIndex + ")",
                    sub.getCell(off + 2).getCellFormula(), w);
            assertValue(sub.getCell(off + 2), sum, evaluator, w + " sum");
            String scoreLetter = letter(off + 3);
            assertEquals("SUBTOTAL(1," + scoreLetter + (start + 1) + ":" + scoreLetter + rowIndex + ")",
                    sub.getCell(off + 3).getCellFormula(), w);
            if (scores > 0) assertValue(sub.getCell(off + 3), scoreSum / scores, evaluator, w + " average");
            String qtyLetter = letter(off + 4);
            assertEquals("SUBTOTAL(" + (mix.maxNotMin() ? 4 : 5) + "," + qtyLetter + (start + 1) + ":" + qtyLetter + rowIndex + ")",
                    sub.getCell(off + 4).getCellFormula(), w);
            assertValue(sub.getCell(off + 4), (mix.maxNotMin() ? max : min) == null ? 0 : (mix.maxNotMin() ? max : min), evaluator, w + " min/max");
            String noteLetter = letter(off + 5);
            assertEquals("SUBTOTAL(3," + noteLetter + (start + 1) + ":" + noteLetter + rowIndex + ")",
                    sub.getCell(off + 5).getCellFormula(), w);
            assertValue(sub.getCell(off + 5), notes, evaluator, w + " countA");
            assertSubtotalLook(sub, off, mix, w);
            sheetAmount += sum;
            sheetNotes += notes;
            rowIndex++;
        }
        int lastBody = rowIndex - 1;
        if (mix.afterData()) {
            assertEquals("AFTER", sheet.getRow(rowIndex).getCell(0).getStringCellValue(), at + ": after data, above the summary");
            rowIndex++;
        }
        if (mix.summary()) {
            var row = sheet.getRow(rowIndex);
            String w = at + " summary row " + rowIndex;
            assertNotNull(row, w);
            // SUBTOTAL skips the subtotal rows inside its range, so the sheet total counts every row once
            assertEquals("Total", row.getCell(0).getStringCellValue(), w);
            assertValue(row.getCell(off + 2), sheetAmount, evaluator, w + " sum");
            assertEquals("SUBTOTAL(9," + letter(off + 2) + (h + 2) + ":" + letter(off + 2) + (lastBody + 1) + ")",
                    row.getCell(off + 2).getCellFormula(), w);
            rowIndex++;
        }
        assertEquals(rowIndex - 1, sheet.getLastRowNum(), at + ": nothing but what the model holds, and no gap");
        if (mix.freeze()) {
            var pane = sheet.getPaneInformation();
            assertNotNull(pane, at + ": frozen pane");
            assertTrue(pane.isFreezePane(), at);
            assertEquals(mix.header() + 1, pane.getHorizontalSplitPosition(), at + ": rows kept in view: the preamble and the header");
            assertEquals(1, pane.getVerticalSplitPosition(), at + ": the first column kept in view");
        }
        assertEquals(mix.protect(), sheet.getProtect(), at + ": protection");
    }

    private static void assertDataLook(org.apache.poi.ss.usermodel.Row row, int amountCol, R r, Mix mix, int dataIndex, String w) {
        var style = (XSSFCellStyle) row.getCell(amountCol).getCellStyle();
        String fill = fill(style);
        boolean colored = mix.rowColor() && "A".equals(r.dept());
        if (colored) {
            assertNotEquals("-", fill, w + ": row color");
            assertNotEquals(SUBTOTAL_FILL, fill, w);
        } else if (mix.stripes() && dataIndex % 2 == 1) {
            assertNotEquals("-", fill, w + ": stripe");
        } else {
            assertEquals("-", fill, w + ": no fill");
        }
    }

    private static void assertSubtotalLook(org.apache.poi.ss.usermodel.Row row, int off, Mix mix, String w) {
        for (int c = 0; c < off + 6; c++) {
            var style = (XSSFCellStyle) row.getCell(c).getCellStyle();
            assertEquals(SUBTOTAL_FILL, fill(style), w + ": fill of column " + c + " (no row color, no stripe)");
            assertTrue(row.getSheet().getWorkbook().getFontAt(style.getFontIndex()).getBold(), w + ": bold " + c);
            if (mix.protect()) assertTrue(style.getLocked(), w + ": locked " + c);
        }
    }

    private static String fill(XSSFCellStyle style) {
        if (style.getFillPattern() == FillPatternType.NO_FILL) return "-";
        byte[] rgb = style.getFillForegroundXSSFColor().getRGB();
        return String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    private static void assertValue(Cell cell, double expected, org.apache.poi.ss.usermodel.FormulaEvaluator evaluator, String w) {
        assertEquals(CellType.FORMULA, cell.getCellType(), w);
        assertEquals(expected, evaluator.evaluate(cell).getNumberValue(), 1e-6, w + ": what Excel calculates");
        assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType(), w + ": has a cached result");
        assertEquals(expected, cell.getNumericCellValue(), 1e-6, w + ": the cached result");
    }

    private static boolean isBlank(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private static String text(Cell cell) {
        return isBlank(cell) ? "" : cell.getStringCellValue();
    }

    private static String letter(int index) {
        return String.valueOf((char) ('A' + index));
    }

    private static byte[] write(List<R> rows, Mix mix) throws IOException {
        var out = new ByteArrayOutputStream();
        if (mix.viaWorkbook()) {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<R>sheet("Data");
                configure(sheet, mix);
                sheet.write(rows);
                workbook.finish().writeTo(out);
            }
        } else {
            var writer = ExcelWriter.<R>create();
            configure(writer, mix);
            (mix.viaStream() ? writer.write(rows.stream()) : writer.write(rows)).writeTo(out);
        }
        return out.toByteArray();
    }

    private static <W extends AbstractSheetWriter<R, W>> void configure(W w, Mix mix) {
        if (mix.protect()) w.defaultStyle(d -> d.locked(false)); // applies to the columns added after it
        if (mix.leadingColumn()) w.constColumn("Src", "x");
        w.column("Dept", R::dept)
                .column("Name", R::name)
                .column("Amount", R::amount)
                .column("Score", R::score)
                .column("Qty", R::qty)
                .column("Note", R::note)
                .maxRows(mix.maxRows())
                .rowHeight(mix.rowHeight())
                .groupSubtotal(R::dept, s -> {
                    s.label("Name", k -> (k == null ? "none" : k) + " total").sum("Amount").average("Score").countA("Note");
                    if (mix.maxNotMin()) s.max("Qty"); else s.min("Qty");
                    if (mix.keyColumn()) s.keyColumn("Dept");
                });
        if (mix.preamble()) w.beforeHeader(ctx -> {
            ctx.getSheet().createRow(0).createCell(0).setCellValue("Title");
            return 2;
        });
        if (mix.freeze()) w.freezePane(1, 1);
        if (mix.protect()) w.protectSheet("pw");
        if (mix.rowColor()) w.rowColor(r -> "A".equals(r.dept()) ? ExcelColor.LIGHT_YELLOW : null);
        if (mix.stripes()) w.stripedRows();
        if (mix.summary()) w.summary(s -> s.label("Total").sum("Amount"));
        if (mix.emptyMessage()) w.emptyMessage("nothing here");
        if (mix.afterData()) w.afterData(ctx -> {
            ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("AFTER");
            return ctx.getCurrentRow() + 1;
        });
    }
}
