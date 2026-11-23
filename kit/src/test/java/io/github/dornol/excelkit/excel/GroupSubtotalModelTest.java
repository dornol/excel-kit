package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Random;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Writes random data with random group runs, sheet sizes and feature mixes through both writers, and compares
 * every sheet with a model of what it must hold: data rows in order, one subtotal row after every run of a
 * key (a sheet's last run included), the subtotal values as Excel calculates them, and the summary rows.
 */
class GroupSubtotalModelTest {
    record Row(String dept, String name, int amount) {}

    private static final String[] KEYS = {"A", "B", "C"};

    @Test
    void excelWriterMatchesTheModel() throws IOException {
        for (long seed = 0; seed < 300; seed++) check(seed, false);
    }

    @Test
    void workbookSheetWriterMatchesTheModel() throws IOException {
        for (long seed = 0; seed < 300; seed++) check(seed, true);
    }

    private void check(long seed, boolean viaWorkbook) throws IOException {
        Random random = new Random(seed);
        int count = random.nextInt(60);
        List<Row> rows = new ArrayList<>();
        String key = KEYS[0];
        for (int i = 0; i < count; i++) {
            if (random.nextInt(4) == 0) key = KEYS[random.nextInt(KEYS.length)];
            rows.add(new Row(key, "r" + i, random.nextInt(100)));
        }
        int maxRows = 1 + random.nextInt(12);
        boolean summary = random.nextBoolean();
        boolean autoFilter = random.nextBoolean();
        boolean stripes = random.nextBoolean();
        boolean mergeKey = random.nextBoolean();
        boolean pageBreaks = random.nextBoolean();
        boolean cf = random.nextBoolean();
        boolean keyColumn = random.nextBoolean();
        String context = "seed " + seed + (viaWorkbook ? " (workbook)" : " (writer)") + " maxRows " + maxRows + " rows " + count;

        byte[] file = viaWorkbook
                ? writeViaWorkbook(rows, maxRows, summary, autoFilter, stripes, mergeKey, pageBreaks, cf, keyColumn)
                : writeViaWriter(rows, maxRows, summary, autoFilter, stripes, mergeKey, pageBreaks, cf, keyColumn);

        List<List<String>> expected = expectedSheets(rows, maxRows, summary);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            int expectedSheets = Math.max(1, expected.size());
            assertEquals(expectedSheets, wb.getNumberOfSheets(), context + ": sheet count");
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                var sheet = wb.getSheetAt(s);
                List<String> actual = new ArrayList<>();
                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    var row = sheet.getRow(r);
                    assertEquals(true, row != null, context + ": sheet " + s + " has a gap at row " + r);
                    String name = string(row.getCell(1));
                    String first = string(row.getCell(0));
                    if (name.endsWith(" total")) {
                        double value = evaluator.evaluate(row.getCell(2)).getNumberValue();
                        actual.add("S:" + name + ":" + (long) value);
                    } else if (first.equals("Total")) {
                        var cell = row.getCell(2);
                        // Without data rows a summary has a label and no formula.
                        actual.add(cell == null ? "T:none" : "T:" + (long) evaluator.evaluate(cell).getNumberValue());
                    } else {
                        actual.add("D:" + first + ":" + name + ":" + (long) row.getCell(2).getNumericCellValue());
                    }
                }
                List<String> want = s < expected.size() ? expected.get(s) : List.of();
                assertEquals(want, actual, context + ": sheet " + s);
                structure(wb, s, want, context, autoFilter, stripes, mergeKey, pageBreaks, cf, keyColumn);
            }
        }
    }

    /** What a sheet must look like around its rows: filter, rules, merges, page breaks and the look of every row. */
    private static void structure(XSSFWorkbook wb, int sheetIndex, List<String> lines, String context, boolean autoFilter,
                                  boolean stripes, boolean mergeKey, boolean pageBreaks, boolean cf, boolean keyColumn) {
        var sheet = wb.getSheetAt(sheetIndex);
        String at = context + ", sheet " + sheetIndex;
        List<String> body = lines.stream().filter(l -> !l.startsWith("T:")).toList();
        int lastRow = body.size(); // 0-based index of the last data or subtotal row (the header is row 0)

        if (autoFilter) {
            assertEquals("A1:C" + (lastRow + 1), sheet.getCTWorksheet().getAutoFilter().getRef(), at + ": filter range");
        }
        if (cf && !body.isEmpty()) {
            var formatting = sheet.getSheetConditionalFormatting();
            assertEquals(1, formatting.getNumConditionalFormattings(), at + ": rules");
            assertEquals("C2:C" + (lastRow + 1), formatting.getConditionalFormattingAt(0).getFormattingRanges()[0].formatAsString(),
                    at + ": rule range stops below the last subtotal, above the summary");
        }

        // runs of data rows, from the model: [first line index, last line index] (0-based within the body)
        List<int[]> runs = new ArrayList<>();
        List<Integer> breaks = new ArrayList<>();
        String previousKey = null;
        int runStart = -1;
        for (int i = 0; i < body.size(); i++) {
            String line = body.get(i);
            if (line.startsWith("D:")) {
                String key = line.split(":")[1];
                if (runStart < 0) {
                    runStart = i;
                    // a new run after a subtotal (not the first one of the sheet): the page starts here
                    if (previousKey != null) breaks.add(i); // POI: below the row before it, the subtotal at body index i - 1 = sheet row i
                }
                previousKey = key;
            } else if (runStart >= 0) {
                runs.add(new int[]{runStart, i - 1});
                runStart = -1;
            }
        }
        if (mergeKey) {
            List<String> want = new ArrayList<>();
            for (int[] run : runs) if (run[1] > run[0]) want.add("A" + (run[0] + 2) + ":A" + (run[1] + 2));
            List<String> got = new ArrayList<>();
            for (var region : sheet.getMergedRegions()) got.add(region.formatAsString());
            assertEquals(want, got, at + ": merged regions never take a subtotal row");
        } else {
            assertEquals(0, sheet.getMergedRegions().size(), at + ": no merges");
        }
        if (pageBreaks) {
            int[] got = sheet.getRowBreaks();
            assertEquals(breaks, java.util.Arrays.stream(got).boxed().toList(), at + ": a break below every subtotal row that is followed by another key");
        } else {
            assertEquals(0, sheet.getRowBreaks().length, at + ": no breaks");
        }

        int dataIndex = 0;
        for (int i = 0; i < body.size(); i++) {
            var row = sheet.getRow(i + 1);
            var style = (org.apache.poi.xssf.usermodel.XSSFCellStyle) row.getCell(2).getCellStyle();
            boolean bold = wb.getFontAt(style.getFontIndex()).getBold();
            String fill = style.getFillPattern() == org.apache.poi.ss.usermodel.FillPatternType.NO_FILL ? "-"
                    : hex(style.getFillForegroundXSSFColor().getRGB());
            if (body.get(i).startsWith("S:")) {
                String key = body.get(i).split(":")[1].replace(" total", "");
                var keyCell = row.getCell(0);
                if (keyColumn) assertEquals(key, keyCell.getStringCellValue(), at + ": row " + (i + 2) + " shows its key");
                else assertEquals(org.apache.poi.ss.usermodel.CellType.BLANK, keyCell.getCellType(), at + ": row " + (i + 2) + " shows no key");
                assertEquals(true, bold, at + ": row " + (i + 2) + " subtotal is bold");
                assertEquals("D9D9D9", fill, at + ": row " + (i + 2) + " subtotal fill");
            } else {
                assertEquals(false, bold, at + ": row " + (i + 2) + " data is not bold");
                assertEquals(stripes && dataIndex % 2 == 1 ? "F2F2F2" : "-", fill, at + ": row " + (i + 2) + " stripe of data row " + dataIndex);
                dataIndex++;
            }
        }
    }

    private static String hex(byte[] rgb) {
        return String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }

    /** The rows each sheet must hold below its header: data, a subtotal after every run, then the summary. */
    private static List<List<String>> expectedSheets(List<Row> rows, int maxRows, boolean summary) {
        List<List<String>> sheets = new ArrayList<>();
        if (rows.isEmpty() && summary) sheets.add(List.of("T:none"));
        for (int from = 0; from < rows.size(); from += maxRows) {
            List<Row> chunk = rows.subList(from, Math.min(rows.size(), from + maxRows));
            List<String> lines = new ArrayList<>();
            int sheetTotal = 0;
            int i = 0;
            while (i < chunk.size()) {
                String key = chunk.get(i).dept();
                int sum = 0;
                while (i < chunk.size() && Objects.equals(chunk.get(i).dept(), key)) {
                    Row row = chunk.get(i++);
                    lines.add("D:" + row.dept() + ":" + row.name() + ":" + row.amount());
                    sum += row.amount();
                }
                lines.add("S:" + key + " total:" + sum);
                sheetTotal += sum;
            }
            if (summary) lines.add("T:" + sheetTotal);
            sheets.add(lines);
        }
        return sheets;
    }

    private static byte[] writeViaWriter(List<Row> rows, int maxRows, boolean summary, boolean autoFilter,
                                         boolean stripes, boolean mergeKey, boolean pageBreaks, boolean cf, boolean keyColumn) throws IOException {
        var writer = ExcelWriter.<Row>create()
                .column("Dept", Row::dept, c -> c.mergeSameValues(mergeKey))
                .column("Name", Row::name)
                .column("Amount", Row::amount)
                .maxRows(maxRows)
                .groupSubtotal(Row::dept, s -> { s.label("Name", k -> k + " total").sum("Amount"); if (keyColumn) s.keyColumn("Dept"); });
        if (summary) writer.summary(s -> s.label("Total").sum("Amount"));
        if (autoFilter) writer.autoFilter(true);
        if (stripes) writer.stripedRows();
        if (pageBreaks) writer.pageBreakOnChange(Row::dept);
        if (cf) writer.conditionalFormatting(c -> c.columns(2).greaterThan("50", io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW));
        var out = new ByteArrayOutputStream();
        writer.write(rows).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] writeViaWorkbook(List<Row> rows, int maxRows, boolean summary, boolean autoFilter,
                                           boolean stripes, boolean mergeKey, boolean pageBreaks, boolean cf, boolean keyColumn) throws IOException {
        try (var workbook = ExcelWorkbook.create()) {
            var sheet = workbook.<Row>sheet("Data")
                    .column("Dept", Row::dept, c -> c.mergeSameValues(mergeKey))
                    .column("Name", Row::name)
                    .column("Amount", Row::amount)
                    .maxRows(maxRows)
                    .groupSubtotal(Row::dept, s -> { s.label("Name", k -> k + " total").sum("Amount"); if (keyColumn) s.keyColumn("Dept"); });
            if (summary) sheet.summary(s -> s.label("Total").sum("Amount"));
            if (autoFilter) sheet.autoFilter(true);
            if (stripes) sheet.stripedRows();
            if (pageBreaks) sheet.pageBreakOnChange(Row::dept);
            if (cf) sheet.conditionalFormatting(c -> c.columns(2).greaterThan("50", io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW));
            sheet.write(rows);
            var out = new ByteArrayOutputStream();
            workbook.finish().writeTo(out);
            return out.toByteArray();
        }
    }

    private static String string(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return "";
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : String.valueOf(cell.getNumericCellValue());
    }
}
