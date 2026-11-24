package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.testsupport.WorkbookSnapshot;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.fail;

/**
 * The same rows with the same settings, written by {@link ExcelWriter} and by a sheet of an {@link ExcelWorkbook}
 * ({@link ExcelSheetWriter}), must give the same workbook, sheet names aside: two writers that share the settings
 * ({@link SheetWriter}) must not each have an engine of their own that drifts.
 */
class WriterEngineParityTest {
    record Row(String dept, String name, int amount, double ratio, LocalDate day, String note) {}

    private static List<Row> rows(Random random, int count) {
        List<Row> rows = new ArrayList<>();
        String dept = "A";
        for (int i = 0; i < count; i++) {
            if (random.nextInt(3) == 0) dept = "DEPT" + (char) ('A' + random.nextInt(4));
            rows.add(new Row(dept, "n" + i, random.nextInt(2000) - 500, random.nextInt(1000) / 10.0,
                    LocalDate.of(2026, 1 + random.nextInt(12), 1 + random.nextInt(28)), random.nextInt(4) == 0 ? null : "note" + i));
        }
        return rows;
    }

    /** What is switched on for one run. */
    private record Mix(int count, int maxRows, boolean subtotal, boolean summary, boolean emptyMessage, boolean stripes,
                       boolean rowColor, boolean rowStyle, boolean autoFilter, boolean freeze, boolean pageBreak,
                       boolean condFmt, boolean chart, boolean table, boolean namedRange, boolean print, boolean protect,
                       boolean tab, boolean heights, boolean preamble, boolean afterData, boolean defaultStyle,
                       boolean mergeDept, boolean groups, boolean comments, boolean validation, boolean dropdown,
                       boolean hiddenColumn, boolean outline, boolean rowNumber, boolean constant, boolean widths,
                       boolean cellColor, boolean groupComment, boolean autoWidthRows, boolean zoom, boolean progress) {}

    private static Mix mix(Random r) {
        int count = r.nextInt(6) == 0 ? 0 : r.nextInt(40);
        boolean subtotal = r.nextInt(4) == 0;
        boolean table = !subtotal && r.nextInt(6) == 0;
        boolean mergeDept = !table && r.nextBoolean();
        int maxRows = table ? 1_000 : 1 + r.nextInt(12);
        boolean emptyMessage = !table && r.nextBoolean();
        boolean chart = !subtotal && r.nextInt(5) == 0;
        boolean namedRange = !subtotal && r.nextInt(5) == 0;
        return new Mix(count, maxRows, subtotal, r.nextBoolean(), emptyMessage, r.nextBoolean(), r.nextBoolean(), r.nextBoolean(),
                r.nextBoolean(), r.nextBoolean(), !subtotal && !table && r.nextBoolean(), r.nextBoolean(), chart, table, namedRange,
                r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(),
                mergeDept, r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(),
                r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(), r.nextBoolean(),
                r.nextBoolean());
    }

    /** One setting written once, for both writers. */
    private static <W extends SheetWriter<Row, W>> void configure(W w, Mix m) {
        if (m.defaultStyle()) w.defaultStyle(d -> d.border(ExcelBorderStyle.THIN));
        if (m.rowNumber()) w.rowNumberColumn("No");
        w.column("Dept", Row::dept, c -> {
            if (m.mergeDept()) c.mergeSameValues();
            if (m.dropdown()) c.dropdown("DEPTA", "DEPTB");
            if (m.groups()) c.group("People");
        });
        w.column("Name", Row::name, c -> {
            if (m.comments()) c.comment(row -> "about " + row.name()).headerComment("the name").requiredMark();
            if (m.groups()) c.group("People");
            if (m.widths()) c.width(18);
        });
        w.column("Amount", Row::amount, c -> {
            c.type(ExcelDataType.INTEGER).format("#,##0");
            if (m.validation()) c.validation(ExcelValidation.integerBetween(0, 1000));
            if (m.cellColor()) c.cellColor((v, row) -> row.amount() > 500 ? ExcelColor.LIGHT_GREEN : null);
            if (m.outline()) c.outline(1);
        });
        w.column("Ratio", Row::ratio, c -> c.type(ExcelDataType.DOUBLE).bold(m.stripes()));
        w.column("Day", Row::day, c -> {
            if (m.hiddenColumn()) c.hidden();
        });
        w.column("Note", Row::note, c -> c.nullValue("-"));
        if (m.constant()) w.constColumn("Source", "x");
        w.maxRows(m.maxRows());
        if (m.subtotal()) w.groupSubtotal(Row::dept, s -> s.label("Name", k -> k + " total").keyColumn("Dept").sum("Amount").average("Ratio"));
        if (m.summary()) w.summary(s -> s.label("Total").sum("Amount"));
        if (m.emptyMessage()) w.emptyMessage("nothing");
        if (m.stripes()) w.stripedRows();
        if (m.rowColor()) w.rowColor(row -> row.amount() < 0 ? ExcelColor.LIGHT_YELLOW : null);
        if (m.rowStyle()) w.rowStyle(row -> row.amount() > 1000, st -> st.bold(true).italic(true));
        if (m.autoFilter()) w.autoFilter(true);
        if (m.freeze()) w.freezePane(1, 1);
        if (m.pageBreak()) w.pageBreakOnChange(Row::dept);
        if (m.condFmt()) w.conditionalFormatting(c -> c.columns(2).greaterThan("100", ExcelColor.LIGHT_YELLOW));
        if (m.chart()) w.chart(c -> c.categoryColumn(m.rowNumber() ? 2 : 1).valueColumn(m.rowNumber() ? 3 : 2, "Amount"));
        if (m.table()) w.table("Rows");
        if (m.namedRange()) w.namedRange("Amounts", m.rowNumber() ? 3 : 2);
        if (m.print()) w.printSetup(p -> p.margins(1, 1, 1, 1));
        if (m.protect()) w.protectSheet("pw");
        if (m.tab()) w.tabColor(ExcelColor.RED);
        if (m.zoom()) w.zoom(80);
        if (m.heights()) w.rowHeight(22).headerRowHeight(30);
        if (m.preamble()) w.beforeHeader(ctx -> {
            ctx.getSheet().createRow(0).createCell(0).setCellValue("Title");
            return 2;
        });
        if (m.afterData()) w.afterData(ctx -> {
            ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("END");
            return ctx.getCurrentRow() + 1;
        });
        if (m.groupComment() && m.groups()) w.groupComment("People", "People");
        if (m.autoWidthRows()) w.autoWidthSampleRows(5);
        if (m.progress()) w.onProgress(3, (n, cursor) -> { });
    }

    private static byte[] viaWriter(List<Row> rows, Mix m) {
        var out = new ByteArrayOutputStream();
        ExcelWriter<Row> writer = ExcelWriter.create();
        configure(writer, m);
        writer.write(rows).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] viaWorkbook(List<Row> rows, Mix m) throws IOException {
        var out = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            var sheet = workbook.<Row>sheet("Data");
            configure(sheet, m);
            sheet.write(rows);
            workbook.finish().writeTo(out);
        }
        return out.toByteArray();
    }

    private static String firstDifference(List<String> a, List<String> b) {
        for (int i = 0; i < Math.max(a.size(), b.size()); i++) {
            String x = i < a.size() ? a.get(i) : "<missing>";
            String y = i < b.size() ? b.get(i) : "<missing>";
            if (!x.equals(y)) return "line " + i + "\n  ExcelWriter:     " + x + "\n  ExcelSheetWriter: " + y;
        }
        return "none";
    }

    @Test
    void bothWritersGiveTheSameWorkbook() throws IOException {
        List<String> failures = new ArrayList<>();
        int compared = 0;
        for (long seed = 0; seed < 400; seed++) {
            Random random = new Random(seed);
            Mix m = mix(random);
            List<Row> rows = rows(random, m.count());
            List<String> a;
            List<String> b;
            try {
                a = WorkbookSnapshot.of(viaWriter(rows, m));
            } catch (RuntimeException e) {
                // a combination the library refuses must be refused by both, in the same words
                try {
                    viaWorkbook(rows, m);
                    failures.add("seed " + seed + ": ExcelWriter refused (" + e.getMessage() + "), the sheet writer did not");
                } catch (RuntimeException other) {
                    if (!String.valueOf(other.getMessage()).equals(String.valueOf(e.getMessage()))) {
                        failures.add("seed " + seed + ": refused differently\n  ExcelWriter:     " + e.getMessage() + "\n  ExcelSheetWriter: " + other.getMessage());
                    }
                }
                continue;
            }
            try {
                b = WorkbookSnapshot.of(viaWorkbook(rows, m));
            } catch (RuntimeException e) {
                failures.add("seed " + seed + ": the sheet writer refused (" + e.getMessage() + "), ExcelWriter did not");
                continue;
            }
            compared++;
            if (!a.equals(b)) failures.add("seed " + seed + " " + m + "\n" + firstDifference(a, b));
        }
        System.out.println("[parity] compared " + compared + " workbooks, " + failures.size() + " differences");
        if (!failures.isEmpty()) {
            fail(failures.size() + " of 400 runs differ; first ones:\n" + String.join("\n\n", failures.subList(0, Math.min(failures.size(), 6))));
        }
    }
}
