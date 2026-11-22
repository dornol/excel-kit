package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Every combination of what a template list can be given, over a template that has content of its own. */
class TemplateFeatureMatrixTest {

    private enum Feature {OWN_HEADERS, TABLE, SUMMARY, AFTER_DATA, ROW_COLOR, ROW_STYLE, ROW_HEIGHT, DEFAULT_STYLE}

    private record Item(String name, int qty) {
    }

    private static final int ROWS = 6;

    private static IntStream masks() {
        return IntStream.range(0, 1 << Feature.values().length);
    }

    private static boolean on(int mask, Feature f) {
        return (mask & (1 << f.ordinal())) != 0;
    }

    private static byte[] template(boolean withHeader) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("Title");
            if (withHeader) {
                var header = sheet.createRow(2);
                header.createCell(0).setCellValue("name");
                header.createCell(1).setCellValue("qty");
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("masks")
    void templateListFeaturesCombine(int mask) throws Exception {
        List<Feature> enabled = new ArrayList<>();
        for (Feature f : Feature.values()) if (on(mask, f)) enabled.add(f);
        String at = "mask=" + mask + " " + enabled;
        boolean ownHeaders = on(mask, Feature.OWN_HEADERS);
        List<Item> items = IntStream.range(0, ROWS).mapToObj(i -> new Item("item " + i, i)).toList();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(!ownHeaders)))) {
            TemplateListWriter<Item> list = template.<Item>list(ownHeaders ? 2 : 3);
            if (on(mask, Feature.DEFAULT_STYLE)) list.defaultStyle(d -> d.bold(true));
            list.column("name", Item::name).column("qty", Item::qty);
            if (on(mask, Feature.TABLE)) list.table("TemplateTable");
            if (on(mask, Feature.SUMMARY)) list.summary(s -> s.label("Total").sum("qty"));
            if (on(mask, Feature.AFTER_DATA)) {
                list.afterData(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("AFTER");
                    return ctx.getCurrentRow() + 1;
                });
            }
            if (on(mask, Feature.ROW_COLOR)) list.rowColor(i -> i.qty() == 1 ? ExcelColor.LIGHT_RED : null);
            if (on(mask, Feature.ROW_STYLE)) list.rowStyle(i -> i.qty() == 2, s -> s.italic(true));
            if (on(mask, Feature.ROW_HEIGHT)) list.rowHeight(30f);
            if (ownHeaders) list.writeWithHeaders(items);
            else list.write(items);
            template.finish().writeTo(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFSheet sheet = wb.getSheetAt(0);
            List<String> col0 = new ArrayList<>();
            int firstItemRow = -1;
            for (Row row : sheet) {
                Cell c = row.getCell(0);
                if (c == null || c.getCellType() != CellType.STRING) continue;
                col0.add(c.getStringCellValue());
                if (firstItemRow < 0 && c.getStringCellValue().equals("item 0")) firstItemRow = row.getRowNum();
            }
            assertTrue(col0.contains("Title"), at + " the template's title survives: " + col0);

            List<String> names = col0.stream().filter(s -> s.startsWith("item ")).toList();
            assertEquals(items.stream().map(Item::name).toList(), names, at);
            assertTrue(firstItemRow > 0, at);
            for (int i = 0; i < ROWS; i++) {
                Row row = sheet.getRow(firstItemRow + i);
                assertEquals("item " + i, row.getCell(0).getStringCellValue(), at);
                assertEquals((double) i, row.getCell(1).getNumericCellValue(), at + " qty " + i);
                if (on(mask, Feature.ROW_HEIGHT)) assertEquals(30f, row.getHeightInPoints(), 0.01, at + " height of row " + i);
                if (on(mask, Feature.ROW_COLOR) && i == 1) {
                    assertEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND,
                            row.getCell(0).getCellStyle().getFillPattern(), at + " the coloured row");
                }
                if (on(mask, Feature.ROW_COLOR) && i == 3) {
                    assertEquals(org.apache.poi.ss.usermodel.FillPatternType.NO_FILL,
                            row.getCell(0).getCellStyle().getFillPattern(), at + " an uncoloured row");
                }
                if (on(mask, Feature.ROW_STYLE)) {
                    assertEquals(i == 2, ((org.apache.poi.xssf.usermodel.XSSFCellStyle) row.getCell(0).getCellStyle()).getFont().getItalic(), at + " italic on row " + i);
                }
                if (on(mask, Feature.DEFAULT_STYLE)) {
                    assertTrue(((org.apache.poi.xssf.usermodel.XSSFCellStyle) row.getCell(0).getCellStyle()).getFont().getBold(), at + " default style on row " + i);
                }
            }
            assertEquals(on(mask, Feature.AFTER_DATA), col0.contains("AFTER"), at + " afterData");
            assertEquals(on(mask, Feature.SUMMARY), col0.contains("Total"), at + " summary");
            if (on(mask, Feature.SUMMARY)) {
                // the sum covers exactly the item rows, not the title, header or footer
                String first = "B" + (firstItemRow + 1), last = "B" + (firstItemRow + ROWS);
                boolean found = false;
                for (Row row : sheet) {
                    for (Cell c : row) {
                        if (c.getCellType() == CellType.FORMULA) {
                            assertEquals("SUM(" + first + ":" + last + ")", c.getCellFormula(), at);
                            found = true;
                        }
                    }
                }
                assertTrue(found, at + " no summary formula");
            }
            assertEquals(on(mask, Feature.TABLE) ? 1 : 0, sheet.getTables().size(), at + " tables");
        }

        OoxmlStructureTest.assertValidPackage(new OoxmlStructureTest.Pkg(out.toByteArray()));
    }
}
