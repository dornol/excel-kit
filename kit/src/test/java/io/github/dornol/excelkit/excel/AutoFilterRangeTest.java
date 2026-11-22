package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The filter of {@code autoFilter(true)} covers the header and the data rows, as Excel writes it, so that every
 * program applies it to the data. It does not reach into the summary rows below the data, and a table, which keeps its
 * filter buttons in its own definition, is not given a second, overlapping filter of the sheet.
 */
class AutoFilterRangeTest {

    private record Item(String name, int qty) {
    }

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i)).toList();
    }

    private static ExcelWriter<Item> writer() {
        return ExcelWriter.<Item>create().autoFilter(true).column("Name", Item::name)
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER));
    }

    private static OoxmlStructureTest.Pkg write(ExcelWriter<Item> writer, int rows) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(items(rows)).writeTo(out);
        return new OoxmlStructureTest.Pkg(out.toByteArray());
    }

    private static String sheetFilter(OoxmlStructureTest.Pkg pkg, int sheet) {
        var filters = pkg.select("xl/worksheets/sheet" + sheet + ".xml", OoxmlStructureTest.local("autoFilter"));
        return filters.isEmpty() ? null : filters.get(0).getAttribute("ref");
    }

    @Test
    void theFilterCoversTheHeaderAndEveryDataRow() throws Exception {
        assertEquals("A1:B8", sheetFilter(write(writer(), 7), 1));
    }

    @Test
    void aSheetWithoutDataRowsFiltersTheHeaderOnly() throws Exception {
        assertEquals("A1:B1", sheetFilter(write(writer(), 0), 1));
    }

    @Test
    void summaryRowsAndAfterDataContentAreOutsideTheFilter() throws Exception {
        ExcelWriter<Item> writer = writer().summary(s -> s.sum("Qty").average("Qty"))
                .afterData(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("after");
                    return ctx.getCurrentRow() + 1;
                });

        assertEquals("A1:B6", sheetFilter(write(writer, 5), 1));
    }

    @Test
    void aGroupedHeaderIsFilteredFromTheRowOfTheColumnNames() throws Exception {
        ExcelWriter<Item> writer = ExcelWriter.<Item>create().autoFilter(true)
                .column("Name", Item::name, c -> c.group("G")).column("Qty", Item::qty, c -> c.group("G"));

        assertEquals("A2:B5", sheetFilter(write(writer, 3), 1));
    }

    @Test
    void everyRolledOverSheetHasTheFilterOfItsOwnRows() throws Exception {
        var pkg = write(writer().maxRows(4), 10);

        assertEquals("A1:B5", sheetFilter(pkg, 1));
        assertEquals("A1:B5", sheetFilter(pkg, 2));
        assertEquals("A1:B3", sheetFilter(pkg, 3));
    }

    @Test
    void theHiddenNameExcelNeedsForAFilterHasTheSameRange() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer().maxRows(4).write(items(10)).writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            for (int index = 0; index < 3; index++) {
                final int sheet = index;
                var names = wb.getAllNames().stream()
                        .filter(n -> n.getNameName().equals("_xlnm._FilterDatabase") && n.getSheetIndex() == sheet).toList();
                assertEquals(1, names.size(), "sheet " + sheet);
                String filter = sheetFilter(new OoxmlStructureTest.Pkg(out.toByteArray()), sheet + 1);
                String expected = "$A$1:$B$" + filter.substring(filter.lastIndexOf('B') + 1);
                assertTrue(names.get(0).getRefersToFormula().endsWith(expected), names.get(0).getRefersToFormula() + " vs " + filter);
            }
        }
    }

    @Test
    void aTableKeepsItsFilterButtonsInItselfAndTheSheetHasNoFilterOverIt() throws Exception {
        var pkg = write(writer().table("FilteredTable"), 6);

        assertNotNull(pkg.parts.get("xl/tables/table1.xml"));
        assertEquals(null, sheetFilter(pkg, 1), "a sheet-level filter would overlap the table");
        var tableFilter = pkg.select("xl/tables/table1.xml", OoxmlStructureTest.local("autoFilter"));
        assertEquals(1, tableFilter.size());
        assertEquals(pkg.doc("xl/tables/table1.xml").getDocumentElement().getAttribute("ref"),
                tableFilter.get(0).getAttribute("ref"), "the table's filter covers the table");
        assertEquals("A1:B7", tableFilter.get(0).getAttribute("ref"));
    }

    @Test
    void aTableWithoutAnAutoFilterHasNoFilterButtons() throws Exception {
        var pkg = write(ExcelWriter.<Item>create().table("PlainTable").column("Name", Item::name), 3);

        assertEquals(null, sheetFilter(pkg, 1));
        assertEquals(0, pkg.select("xl/tables/table1.xml", OoxmlStructureTest.local("autoFilter")).size());
    }

    @Test
    void theMultiSheetWriterFiltersTheSameWay() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Plain").autoFilter().column("Name", Item::name).column("Qty", Item::qty)
                    .summary(s -> s.sum("Qty")).write(items(5).stream());
            book.<Item>sheet("Tabled").autoFilter().table("SheetTable").column("Name", Item::name)
                    .write(items(2).stream());
            book.<Item>sheet("Rolled").autoFilter().maxRows(2).column("Name", Item::name).write(items(3).stream());
            book.finish().writeTo(out);
        }
        var pkg = new OoxmlStructureTest.Pkg(out.toByteArray());

        assertEquals("A1:B6", sheetFilter(pkg, 1));
        assertEquals(null, sheetFilter(pkg, 2));
        assertEquals("A1:A3", sheetFilter(pkg, 3));
        assertEquals("A1:A2", sheetFilter(pkg, 4));
        assertEquals("A1:A3", pkg.select("xl/tables/table1.xml", OoxmlStructureTest.local("autoFilter")).get(0).getAttribute("ref"));
    }
}
