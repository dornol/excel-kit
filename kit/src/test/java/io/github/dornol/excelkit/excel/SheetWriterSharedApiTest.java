package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

/**
 * {@code ExcelWriter} and {@code ExcelSheetWriter} share one base type and one column configurer type, so a setting
 * is written once for both, and the two have the same column methods.
 */
class SheetWriterSharedApiTest {
    record Row(String name, int amount, double ratio) {}

    private static final List<Row> ROWS = List.of(new Row("a", 1, 0.5), new Row("b", 2, 0.25));

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    /** A setting written once for both writers. */
    private static <W extends SheetWriter<Row, W>> W standardLook(W writer) {
        return writer.stripedRows().autoFilter(true).freezePane(0, 1).emptyMessage("No data").tabColor(ExcelColor.LIGHT_BLUE);
    }

    private static final Consumer<ColumnConfig<Row>> MONEY = c -> c.type(ExcelDataType.INTEGER).format("#,##0").bold(true);

    @Nested
    class OneSettingForBothWriters {
        @Test
        void aSettingWrittenOnceAppliesToTheWriterAndToAWorkbookSheet() throws IOException {
            try (var single = open(standardLook(ExcelWriter.<Row>create()).column("Name", Row::name).column("Amount", Row::amount, MONEY).write(ROWS))) {
                assertLook(single.getSheetAt(0));
            }
            try (var workbook = ExcelWorkbook.create()) {
                standardLook(workbook.<Row>sheet("S")).column("Name", Row::name).column("Amount", Row::amount, MONEY).write(ROWS);
                try (var wb = open(workbook.finish())) {
                    assertLook(wb.getSheetAt(0));
                }
            }
        }

        private void assertLook(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
            assertNotNull(sheet.getCTWorksheet().getAutoFilter(), "auto filter");
            assertNotNull(sheet.getPaneInformation(), "frozen header");
            assertTrue(sheet.getPaneInformation().isFreezePane());
            assertNotNull(sheet.getCTWorksheet().getSheetPr().getTabColor(), "tab color");
            assertEquals("#,##0", sheet.getRow(1).getCell(1).getCellStyle().getDataFormatString(), "the column setting");
        }

        @Test
        void theWriterTypesAreTheBaseType() {
            SheetWriter<Row, ExcelWriter<Row>> a = ExcelWriter.<Row>create();
            assertNotNull(a);
            try (var workbook = ExcelWorkbook.create()) {
                SheetWriter<Row, ExcelSheetWriter<Row>> b = workbook.<Row>sheet("S");
                assertNotNull(b);
            }
        }

        @Test
        void aColumnSettingKeptInAVariableWorksOnEveryWriter() throws IOException {
            try (var single = open(ExcelWriter.<Row>create().column("Amount", Row::amount, MONEY).write(ROWS))) {
                assertEquals("#,##0", single.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getDataFormatString());
            }
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("S").column("Amount", Row::amount, MONEY).write(ROWS);
                try (var wb = open(workbook.finish())) {
                    assertEquals("#,##0", wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getDataFormatString());
                }
            }
        }
    }

    @Nested
    class SameColumnMethodsOnBoth {
        private ExcelHandler viaWriter(Consumer<SheetWriter<Row, ?>> columns) {
            ExcelWriter<Row> writer = ExcelWriter.create();
            columns.accept(writer);
            return writer.write(ROWS);
        }

        private byte[] viaSheet(Consumer<SheetWriter<Row, ?>> columns) throws IOException {
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S");
                columns.accept(sheet);
                sheet.write(ROWS);
                var out = new ByteArrayOutputStream();
                workbook.finish().writeTo(out);
                return out.toByteArray();
            }
        }

        private List<List<String>> table(XSSFWorkbook wb) {
            var fmt = new org.apache.poi.ss.usermodel.DataFormatter();
            var rows = new java.util.ArrayList<List<String>>();
            for (var row : wb.getSheetAt(0)) {
                var cells = new java.util.ArrayList<String>();
                for (var cell : row) cells.add(fmt.formatCellValue(cell));
                rows.add(cells);
            }
            return rows;
        }

        @Test
        void everyColumnMethodGivesTheSameSheetOnBothWriters() throws IOException {
            Consumer<SheetWriter<Row, ?>> columns = w -> {
                w.column("Name", Row::name);
                w.column("Amount", Row::amount, c -> c.type(ExcelDataType.INTEGER));
                w.column("Position", (row, cursor) -> cursor.getCurrentTotal());
                w.column("Position+", (row, cursor) -> cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG));
                w.columnIf("Shown", true, Row::name);
                w.columnIf("Hidden", false, Row::name);
                w.columnIf("ShownCfg", true, Row::amount, c -> c.type(ExcelDataType.INTEGER));
                w.columnIf("PositionIf", true, (row, cursor) -> cursor.getCurrentTotal());
                w.columnIf("PositionIf+", true, (row, cursor) -> cursor.getCurrentTotal(), c -> c.bold(true));
                w.columnIf("HiddenRow", false, (row, cursor) -> cursor.getCurrentTotal());
                w.constColumn("Const", "x");
                w.constColumn("Const+", "y", c -> c.bold(true));
                w.constColumnIf("ConstIf", true, "z");
                w.constColumnIf("ConstIf+", true, "z2", c -> c.bold(true));
                w.constColumnIf("ConstOff", false, "never");
            };
            try (var a = open(viaWriter(columns)); var b = new XSSFWorkbook(new ByteArrayInputStream(viaSheet(columns)))) {
                var expected = table(a);
                assertEquals(expected, table(b));
                assertEquals(List.of("Name", "Amount", "Position", "Position+", "Shown", "ShownCfg", "PositionIf", "PositionIf+",
                        "Const", "Const+", "ConstIf", "ConstIf+"), expected.get(0), "false conditions add no column");
                assertEquals("x", expected.get(1).get(8));
                assertEquals("z2", expected.get(1).get(11));
            }
        }

        @Test
        void aNullConfigurerMeansNoSettingsOnEveryMethod() throws IOException {
            Consumer<SheetWriter<Row, ?>> columns = w -> {
                w.column("A", Row::name, (Consumer<ColumnConfig<Row>>) null);
                w.column("B", (row, cursor) -> cursor.getCurrentTotal(), null);
                w.columnIf("C", true, Row::name, null);
                w.constColumn("D", "d", null);
                w.constColumnIf("E", true, "e", null);
            };
            try (var a = open(viaWriter(columns)); var b = new XSSFWorkbook(new ByteArrayInputStream(viaSheet(columns)))) {
                assertEquals(table(a), table(b));
                assertEquals(5, table(a).get(0).size());
            }
        }

        @Test
        void aColumnNeedsAFunction() {
            ExcelWriter<Row> writer = ExcelWriter.create();
            assertThrows(NullPointerException.class, () -> writer.column("A", (java.util.function.Function<Row, Object>) null));
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S");
                assertThrows(NullPointerException.class, () -> sheet.column("A", (java.util.function.Function<Row, Object>) null));
                assertThrows(NullPointerException.class, () -> sheet.column("A", (RowFunctionOfRow) null));
            }
        }

        private interface RowFunctionOfRow extends io.github.dornol.excelkit.core.read.RowFunction<Row, Object> {}
    }

    @Nested
    class OwnCellStyle {
        @Test
        void aColumnTakesACellStyleBuiltByHandOnEveryWriter() throws IOException {
            ExcelWriter<Row> writer = ExcelWriter.create();
            try (var workbook = new XSSFWorkbook()) {
                CellStyle mine = workbook.createCellStyle();
                assertSame(mine, writer.buildColumn("A", (r, c) -> r.name(), new ColumnConfig<Row>().style(mine).bold(true)).getStyle());
            }
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S");
                try (var other = new XSSFWorkbook()) {
                    CellStyle mine = other.createCellStyle();
                    assertSame(mine, sheet.buildColumn("A", (r, c) -> r.name(), new ColumnConfig<Row>().style(mine)).getStyle());
                }
            }
        }
    }
}
