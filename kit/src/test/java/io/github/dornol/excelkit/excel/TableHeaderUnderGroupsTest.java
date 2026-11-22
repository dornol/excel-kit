package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A table needs its header cells to hold the column names and not to be merged: Excel reports a problem with the content
 * of a file otherwise (found on a real Excel with a table, a group header and a column without a group). Group rows above
 * the table header may be merged as before; a column without a group keeps its name in the column names row.
 */
class TableHeaderUnderGroupsTest {

    private record Row(String a, String b, String c, String d) {
    }

    private static final List<Row> ROWS = List.of(new Row("a1", "b1", "c1", "d1"), new Row("a2", "b2", "c2", "d2"));

    private static XSSFWorkbook write(boolean sheetWriter, boolean table, boolean grouped) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        if (sheetWriter) {
            try (ExcelWorkbook book = ExcelWorkbook.create()) {
                ExcelSheetWriter<Row> w = book.<Row>sheet("Data");
                if (table) w.table("GroupedTable");
                // First: a column without a group; Mid: one level; Deep: two levels; No.: a row number
                w.column("First", Row::a);
                w.column("Mid", Row::b, c -> { if (grouped) c.group("Top"); });
                w.column("Deep", Row::c, c -> { if (grouped) c.group("Top", "Inner"); });
                w.rowNumberColumn("No.");
                w.write(ROWS.stream());
                book.finish().writeTo(out);
            }
        } else {
            ExcelWriter<Row> w = ExcelWriter.create();
            if (table) w.table("GroupedTable");
            w.column("First", Row::a);
            w.column("Mid", Row::b, c -> { if (grouped) c.group("Top"); });
            w.column("Deep", Row::c, c -> { if (grouped) c.group("Top", "Inner"); });
            w.rowNumberColumn("No.");
            w.write(ROWS).writeTo(out);
        }
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static Set<String> mergedRanges(XSSFSheet sheet) {
        Set<String> merges = new HashSet<>();
        sheet.getMergedRegions().forEach(m -> merges.add(m.formatAsString()));
        return merges;
    }

    @Test
    void underATableAColumnWithoutAGroupKeepsItsNameInTheColumnNamesRowUnmerged() throws Exception {
        for (boolean sheetWriter : new boolean[]{false, true}) {
            try (XSSFWorkbook wb = write(sheetWriter, true, true)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                // 3 header rows: Top / Inner / names; the table starts on the names row
                assertEquals(2, sheet.getTables().get(0).getStartCellReference().getRow(), "sheetWriter=" + sheetWriter);
                var names = sheet.getRow(2);
                assertEquals("First", names.getCell(0).getStringCellValue());
                assertEquals("Mid", names.getCell(1).getStringCellValue());
                assertEquals("Deep", names.getCell(2).getStringCellValue());
                assertEquals("No.", names.getCell(3).getStringCellValue());
                // nothing is merged into the names row; the cells above a column without a group are empty
                for (var merged : sheet.getMergedRegions()) {
                    assertTrue(merged.getLastRow() < 2, "merged range " + merged.formatAsString() + " reaches the table header");
                }
                assertEquals(CellType.BLANK, sheet.getRow(0).getCell(0).getCellType());
                assertEquals(CellType.BLANK, sheet.getRow(1).getCell(0).getCellType());
                assertEquals(CellType.BLANK, sheet.getRow(0).getCell(3).getCellType());
                // the group labels above are kept (and merged across their columns)
                assertEquals("Top", sheet.getRow(0).getCell(1).getStringCellValue());
                assertTrue(mergedRanges(sheet).contains("B1:C1"), mergedRanges(sheet).toString());
            }
        }
    }

    @Test
    void theTablesColumnNamesAreTheHeaderCells() throws Exception {
        try (XSSFWorkbook wb = write(false, true, true)) {
            var table = wb.getSheetAt(0).getTables().get(0);
            var columns = table.getCTTable().getTableColumns().getTableColumnArray();
            assertEquals(List.of("First", "Mid", "Deep", "No."), java.util.Arrays.stream(columns).map(c -> c.getName()).toList());
        }
    }

    @Test
    void withoutATableTheNameStillMovesToTheTopAndMergesDown() throws Exception {
        for (boolean sheetWriter : new boolean[]{false, true}) {
            try (XSSFWorkbook wb = write(sheetWriter, false, true)) {
                XSSFSheet sheet = wb.getSheetAt(0);
                assertEquals("First", sheet.getRow(0).getCell(0).getStringCellValue(), "sheetWriter=" + sheetWriter);
                assertTrue(mergedRanges(sheet).contains("A1:A3"), mergedRanges(sheet).toString());
                assertTrue(mergedRanges(sheet).contains("D1:D3"), mergedRanges(sheet).toString());
            }
        }
    }

    @Test
    void aTableWithoutGroupsIsNotAffected() throws Exception {
        try (XSSFWorkbook wb = write(false, true, false)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertEquals(0, sheet.getTables().get(0).getStartCellReference().getRow());
            assertEquals("First", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals(0, sheet.getNumMergedRegions());
        }
    }
}
