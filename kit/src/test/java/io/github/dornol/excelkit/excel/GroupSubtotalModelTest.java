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
        String context = "seed " + seed + (viaWorkbook ? " (workbook)" : " (writer)") + " maxRows " + maxRows + " rows " + count;

        byte[] file = viaWorkbook
                ? writeViaWorkbook(rows, maxRows, summary, autoFilter, stripes, mergeKey, pageBreaks)
                : writeViaWriter(rows, maxRows, summary, autoFilter, stripes, mergeKey, pageBreaks);

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
            }
        }
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
                                         boolean stripes, boolean mergeKey, boolean pageBreaks) throws IOException {
        var writer = ExcelWriter.<Row>create()
                .column("Dept", Row::dept, c -> c.mergeSameValues(mergeKey))
                .column("Name", Row::name)
                .column("Amount", Row::amount)
                .maxRows(maxRows)
                .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
        if (summary) writer.summary(s -> s.label("Total").sum("Amount"));
        if (autoFilter) writer.autoFilter(true);
        if (stripes) writer.stripedRows();
        if (pageBreaks) writer.pageBreakOnChange(Row::dept);
        var out = new ByteArrayOutputStream();
        writer.write(rows).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] writeViaWorkbook(List<Row> rows, int maxRows, boolean summary, boolean autoFilter,
                                           boolean stripes, boolean mergeKey, boolean pageBreaks) throws IOException {
        try (var workbook = ExcelWorkbook.create()) {
            var sheet = workbook.<Row>sheet("Data")
                    .column("Dept", Row::dept, c -> c.mergeSameValues(mergeKey))
                    .column("Name", Row::name)
                    .column("Amount", Row::amount)
                    .maxRows(maxRows)
                    .groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").sum("Amount"));
            if (summary) sheet.summary(s -> s.label("Total").sum("Amount"));
            if (autoFilter) sheet.autoFilter(true);
            if (stripes) sheet.stripedRows();
            if (pageBreaks) sheet.pageBreakOnChange(Row::dept);
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
