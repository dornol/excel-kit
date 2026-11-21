package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** What a {@link TemplateListWriter} does to the sheet of a template, beyond writing the rows. */
class TemplateListWriterDetailsTest {

    record Item(String name, int qty) {
    }

    private static final List<Item> ITEMS = List.of(new Item("a", 1), new Item("b", 2), new Item("c", 4));

    /** A template with a title on row 1 and, when {@code headerRow >= 0}, the column headers on that row (0-based). */
    private static byte[] template(int headerRow) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet sheet = workbook.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("Title");
            if (headerRow >= 0) {
                var header = sheet.createRow(headerRow);
                header.createCell(0).setCellValue("name");
                header.createCell(1).setCellValue("qty");
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private interface Listing {
        void write(TemplateListWriter<Item> list, List<Item> rows);
    }

    /** Writes the items below the template's header (header row 2, data from row 3) or with their own header. */
    private static XSSFSheet write(boolean ownHeaders, Consumer<TemplateListWriter<Item>> configure, List<Item> rows) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(ownHeaders ? -1 : 2)))) {
            TemplateListWriter<Item> list = template.<Item>list(ownHeaders ? 2 : 3)
                    .column("name", Item::name).column("qty", Item::qty);
            configure.accept(list);
            if (ownHeaders) list.writeWithHeaders(rows);
            else list.write(rows);
            template.finish().writeTo(out);
        }
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray())).getSheetAt(0);
    }

    private static String formulaBelow(XSSFSheet sheet, int row) {
        for (int r = row; r <= sheet.getLastRowNum(); r++) {
            var current = sheet.getRow(r);
            if (current == null) continue;
            for (var cell : current) {
                if (cell.getCellType() == CellType.FORMULA) return cell.getCellFormula();
            }
        }
        return null;
    }

    // ---- summary and afterData know where the data is

    @Test
    void aSummaryCoversEveryDataRowWhetherOrNotTheListWritesItsOwnHeader() throws Exception {
        // data rows are 4..6 (1-based) in both cases
        assertEquals("SUM(B4:B6)", formulaBelow(write(false, list -> list.summary(s -> s.sum("qty")), ITEMS), 6));
        assertEquals("SUM(B4:B6)", formulaBelow(write(true, list -> list.summary(s -> s.sum("qty")), ITEMS), 6));
    }

    @Test
    void aSummaryOfAListFromTheFirstRowCoversItAll() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        // the template has no title, so a list can start on the first row
        try (XSSFWorkbook blank = new XSSFWorkbook(); ByteArrayOutputStream bytes = new ByteArrayOutputStream()) {
            blank.createSheet("S");
            blank.write(bytes);
            try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(bytes.toByteArray()))) {
                template.<Item>list(0).column("name", Item::name).column("qty", Item::qty)
                        .summary(s -> s.sum("qty")).write(ITEMS);
                template.finish().writeTo(out);
            }
        }
        XSSFSheet sheet = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray())).getSheetAt(0);
        assertEquals("SUM(B1:B3)", formulaBelow(sheet, 3));
    }

    @Test
    void afterDataIsToldWhereTheHeaderIs() throws Exception {
        List<Integer> headerRows = new ArrayList<>();
        List<Integer> nextRows = new ArrayList<>();
        write(false, list -> list.afterData(ctx -> {
            headerRows.add(ctx.getHeaderRowIndex());
            nextRows.add(ctx.getCurrentRow());
            return ctx.getCurrentRow();
        }), ITEMS);
        write(true, list -> list.afterData(ctx -> {
            headerRows.add(ctx.getHeaderRowIndex());
            return ctx.getCurrentRow();
        }), ITEMS);
        assertEquals(List.of(2, 2), headerRows);
        assertEquals(List.of(6), nextRows);
    }

    // ---- styles

    @Test
    void columnSettingsApplyToTheCellsOfTheList() throws Exception {
        XSSFSheet sheet = write(false, list -> list.column("note", Item::name, c -> c.bold(true))
                .column("rowFn", (item, cursor) -> item.name(), c -> c.fontSize(14)), ITEMS);
        assertFalse(sheet.getRow(3).getCell(0).getCellStyle().getFont().getBold());
        assertTrue(sheet.getRow(3).getCell(2).getCellStyle().getFont().getBold());
        assertEquals(14, sheet.getRow(3).getCell(3).getCellStyle().getFont().getFontHeightInPoints());
    }

    @Test
    void theDefaultStyleReachesTheColumnsDeclaredAfterItUnlessTheColumnSaysOtherwise() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            template.<Item>list(3).column("before", Item::name)
                    .defaultStyle(d -> d.bold(true).backgroundColor(255, 0, 0))
                    .column("plain", Item::name).column("own", Item::name, c -> c.bold(false))
                    .write(ITEMS);
            template.finish().writeTo(out);
        }
        XSSFSheet sheet = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray())).getSheetAt(0);

        assertEquals(FillPatternType.NO_FILL, sheet.getRow(3).getCell(0).getCellStyle().getFillPattern(), "declared before");
        for (int column : new int[]{1, 2}) {
            XSSFCell cell = sheet.getRow(3).getCell(column);
            assertEquals(FillPatternType.SOLID_FOREGROUND, cell.getCellStyle().getFillPattern(), "column " + column);
            assertArrayEquals(new byte[]{(byte) 255, 0, 0}, ((XSSFColor) cell.getCellStyle().getFillForegroundColorColor()).getRGB());
        }
        assertTrue(sheet.getRow(3).getCell(1).getCellStyle().getFont().getBold());
        assertFalse(sheet.getRow(3).getCell(2).getCellStyle().getFont().getBold(), "the column's own setting wins");
    }

    @Test
    void rowStylesAndRowColorsApplyToMatchingRows() throws Exception {
        XSSFSheet sheet = write(false, list -> list.rowStyle(item -> item.qty() >= 2, style -> style.bold(true))
                .rowColor(item -> item.qty() == 1 ? ExcelColor.LIGHT_RED : null), ITEMS);
        assertFalse(sheet.getRow(3).getCell(0).getCellStyle().getFont().getBold());
        assertTrue(sheet.getRow(4).getCell(0).getCellStyle().getFont().getBold());
        assertTrue(sheet.getRow(5).getCell(1).getCellStyle().getFont().getBold());
        assertEquals(FillPatternType.SOLID_FOREGROUND, sheet.getRow(3).getCell(0).getCellStyle().getFillPattern());
        assertEquals(FillPatternType.NO_FILL, sheet.getRow(4).getCell(0).getCellStyle().getFillPattern());
    }

    @Test
    void columnWidthsFollowTheContentAndFixedWidthsAreKept() throws Exception {
        XSSFSheet sheet = write(false, list -> list.column("fixed", Item::name, c -> c.width(30 * 256))
                .column("long", item -> "x".repeat(40)), List.of(new Item("a", 1)));
        assertTrue(sheet.getColumnWidth(3) > sheet.getColumnWidth(0), "40 characters are wider than 1");
        assertTrue(sheet.getColumnWidth(2) >= 30 * 256 - 256, "the fixed width stays: " + sheet.getColumnWidth(2));
    }

    // ---- settings that must make sense

    @Test
    void theAutoWidthSampleMayBeZeroButNotNegative() throws Exception {
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            TemplateListWriter<Item> list = template.<Item>list(3);
            assertDoesNotThrow(() -> list.autoWidthSampleRows(0));
            assertThrows(IllegalArgumentException.class, () -> list.autoWidthSampleRows(-1));
        }
    }

    @Test
    void columnNamesMustBeUnique() throws Exception {
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            TemplateListWriter<Item> list = template.<Item>list(3).column("same", Item::name).column("same", Item::name);
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> list.write(ITEMS));
            assertTrue(error.getMessage().contains("Duplicate column name"), error.getMessage());
        }
    }

    @Test
    void aListNeedsColumns() throws Exception {
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> template.<Item>list(3).write(ITEMS));
            assertTrue(error.getMessage().contains("columns setting required"), error.getMessage());
        }
    }

    // ---- tables

    private static XSSFSheet writeTable(boolean ownHeaders, List<Item> rows) throws Exception {
        return write(ownHeaders, list -> list.table("Items"), rows);
    }

    @Test
    void aTableCoversTheHeaderAndTheDataRows() throws Exception {
        XSSFSheet existing = writeTable(false, ITEMS);
        assertEquals(1, existing.getTables().size());
        assertEquals("Items", existing.getTables().get(0).getName());
        assertEquals(2, existing.getTables().get(0).getStartCellReference().getRow(), "from the template's header");
        assertEquals(5, existing.getTables().get(0).getEndCellReference().getRow());

        XSSFSheet own = writeTable(true, ITEMS);
        assertEquals(2, own.getTables().get(0).getStartCellReference().getRow(), "from the header the list wrote");
        assertEquals(5, own.getTables().get(0).getEndCellReference().getRow());
    }

    @Test
    void aTableIsNotCreatedWhenThereAreNoDataRows() throws Exception {
        assertEquals(0, writeTable(false, List.of()).getTables().size());
        assertEquals(0, writeTable(true, List.of()).getTables().size());
        assertEquals(1, writeTable(false, List.of(new Item("a", 1))).getTables().size(), "one data row is enough");
    }

    @Test
    void aTableFromATemplateNeedsItsHeaderOnTheRowAboveTheData() throws Exception {
        // data from row 0: there is no row above it
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(-1)))) {
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> template.<Item>list(0)
                    .column("name", Item::name).column("qty", Item::qty).table("Items").write(ITEMS));
            assertTrue(error.getMessage().contains("header row before startRow"), error.getMessage());
        }
        // the row above is not the header
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> template.<Item>list(5)
                    .column("name", Item::name).column("qty", Item::qty).table("Items").write(ITEMS));
            assertTrue(error.getMessage().contains("header"), error.getMessage());
        }
    }

    // ---- the row order of a streaming sheet

    @Test
    void afterAListNothingMayBeWrittenAboveItsLastRow() throws Exception {
        try (ExcelTemplateWriter template = new ExcelTemplateWriter(new ByteArrayInputStream(template(2)))) {
            template.<Item>list(3).column("name", Item::name).column("qty", Item::qty).write(ITEMS);
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () -> template.cell(1, 1, "late"));
            assertTrue(error.getMessage().contains("top-down"), error.getMessage());
            assertDoesNotThrow(() -> template.cell(10, 0, "below the list"));
        }
    }
}
