package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Every combination of the sheet-level features that can be switched on together. Each feature has its own
 * tests; what this catches is two features that work alone but not side by side (a table over a rolled-over
 * sheet, a chart on an encrypted workbook, a summary row under group headers, ...). Every case is checked
 * for its data, for what each feature must have produced per sheet, and for a structurally valid package.
 */
class FeatureMatrixTest {

    private static final int ROWS = 11;

    private enum Feature {ROLLOVER, TABLE, SUMMARY, CHART, ENCRYPT, GROUP, CONDITIONAL, PROTECT, ROW_NUMBER, SHEET_OPTIONS}

    private record Item(String name, int qty) {
    }

    private static IntStream masks() {
        return IntStream.range(0, 1 << Feature.values().length);
    }

    private static boolean on(int mask, Feature f) {
        return (mask & (1 << f.ordinal())) != 0;
    }

    private static String describe(int mask) {
        List<Feature> enabled = new ArrayList<>();
        for (Feature f : Feature.values()) if (on(mask, f)) enabled.add(f);
        return "mask=" + mask + " " + enabled;
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("masks")
    void featuresCombineWithoutBreakingEachOther(int mask) throws Exception {
        String at = describe(mask);
        int maxRows = on(mask, Feature.ROLLOVER) ? 4 : 0;
        int sheets = maxRows > 0 ? (int) Math.ceil(ROWS / (double) maxRows) : 1;

        ExcelWriter<Item> writer = ExcelWriter.create();
        if (maxRows > 0) writer.maxRows(maxRows);
        if (on(mask, Feature.ENCRYPT)) writer.password("pw");
        if (on(mask, Feature.TABLE)) writer.table("MatrixTable");
        if (on(mask, Feature.PROTECT)) writer.protectSheet("lock");
        if (on(mask, Feature.SHEET_OPTIONS)) writer.autoFilter(true).freezeRows(1).tabColor(ExcelColor.GREEN);
        boolean group = on(mask, Feature.GROUP);
        writer.column("Name", Item::name, c -> {
            if (group) c.group("G");
        });
        writer.column("Qty", Item::qty, c -> {
            c.type(ExcelDataType.INTEGER);
            if (group) c.group("G");
        });
        // appended last so the Name/Qty positions that the chart and the formatting refer to do not move
        if (on(mask, Feature.ROW_NUMBER)) writer.rowNumberColumn("No.");
        if (on(mask, Feature.CONDITIONAL)) {
            writer.conditionalFormatting(cf -> cf.columns(1).greaterThan("5", ExcelColor.LIGHT_RED));
        }
        if (on(mask, Feature.SUMMARY)) writer.summary(s -> s.label("Total").sum("Qty"));
        if (on(mask, Feature.CHART)) {
            writer.chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).title("Qty")
                    .categoryColumn(0).valueColumn(1, "Qty").position(4, 0, 12, 20));
        }

        List<Item> items = IntStream.range(0, ROWS).mapToObj(i -> new Item("item " + i, i)).toList();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(items).writeTo(out);
        byte[] xlsx = decrypted(out.toByteArray(), on(mask, Feature.ENCRYPT), at);

        // ---- the data: every item once, in order, on the expected sheet ----
        List<String> names = new ArrayList<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            assertEquals(sheets, wb.getNumberOfSheets(), at);
            int globalIndex = 0;
            for (int s = 0; s < sheets; s++) {
                Sheet sheet = wb.getSheetAt(s);
                int dataRowsHere = 0;
                boolean summarySeen = false;
                for (Row row : sheet) {
                    Cell first = row.getCell(0);
                    if (first == null || first.getCellType() != CellType.STRING) continue;
                    String text = first.getStringCellValue();
                    if (text.startsWith("item ")) {
                        names.add(text);
                        assertEquals((double) Integer.parseInt(text.substring(5)),
                                row.getCell(1).getNumericCellValue(), at + " qty of " + text);
                        if (on(mask, Feature.ROW_NUMBER)) {
                            assertEquals(globalIndex + 1, (int) row.getCell(2).getNumericCellValue(), at + " number of " + text);
                        }
                        dataRowsHere++;
                        globalIndex++;
                    } else if (text.equals("Total")) {
                        summarySeen = true;
                    }
                }
                int expectedHere = Math.min(maxRows > 0 ? maxRows : ROWS, ROWS - s * (maxRows > 0 ? maxRows : ROWS));
                assertEquals(expectedHere, dataRowsHere, at + " data rows on sheet " + s);
                assertEquals(on(mask, Feature.SUMMARY), summarySeen, at + " summary on sheet " + s);
            }
        }
        List<String> expected = items.stream().map(Item::name).toList();
        assertEquals(expected, names, at);

        // ---- what each feature must have produced, once per sheet ----
        OoxmlStructureTest.Pkg pkg = new OoxmlStructureTest.Pkg(xlsx);
        List<String> sheetParts = pkg.sheetParts();
        assertEquals(sheets, sheetParts.size(), at);
        for (String part : sheetParts) {
            assertEquals(on(mask, Feature.TABLE) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("tablePart")).size(), at + " table in " + part);
            assertEquals(on(mask, Feature.PROTECT) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("sheetProtection")).size(), at + " protection in " + part);
            assertEquals(on(mask, Feature.CONDITIONAL) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("conditionalFormatting")).size(), at + " formatting in " + part);
            assertEquals(on(mask, Feature.SHEET_OPTIONS) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("autoFilter")).size(), at + " filter in " + part);
            assertEquals(on(mask, Feature.SHEET_OPTIONS) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("pane")).size(), at + " pane in " + part);
            assertEquals(on(mask, Feature.CHART) ? 1 : 0, pkg.select(part, OoxmlStructureTest.local("drawing")).size(), at + " drawing in " + part);
        }
        long charts = pkg.parts.keySet().stream().filter(n -> n.matches("xl/charts/chart\\d+\\.xml")).count();
        assertEquals(on(mask, Feature.CHART) ? sheets : 0, charts, at + " chart parts");
        long tables = pkg.parts.keySet().stream().filter(n -> n.matches("xl/tables/table\\d+\\.xml")).count();
        assertEquals(on(mask, Feature.TABLE) ? sheets : 0, tables, at + " table parts");

        OoxmlStructureTest.assertValidPackage(pkg);
    }

    private static byte[] decrypted(byte[] bytes, boolean encrypted, String at) throws Exception {
        if (!encrypted) return bytes;
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            Decryptor dec = Decryptor.getInstance(new EncryptionInfo(fs));
            assertTrue(dec.verifyPassword("pw"), at);
            return dec.getDataStream(fs).readAllBytes();
        }
    }
}
