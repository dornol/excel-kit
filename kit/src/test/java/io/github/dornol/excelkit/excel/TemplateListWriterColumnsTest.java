package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

/**
 * A list written into a template has the columns of every writer ({@link ColumnsWriter}), and the same result as the
 * same columns on a sheet of a workbook.
 */
class TemplateListWriterColumnsTest {
    record Row(String name, int amount, String note) {}

    private static final List<Row> ROWS = List.of(new Row("a", 1, null), new Row("b", 2, "n2"), new Row("c", 3, null));

    private static byte[] template() throws IOException {
        var out = new ByteArrayOutputStream();
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("Title");
            wb.write(out);
        }
        return out.toByteArray();
    }

    private static final Consumer<ColumnConfig<Row>> MONEY = c -> c.type(ExcelDataType.INTEGER).format("#,##0");

    private static <W extends ColumnsWriter<Row, W>> void columns(W w) {
        w.rowNumberColumn("No.");
        w.column("Name", Row::name);
        w.column("Amount", Row::amount, MONEY);
        w.column("Pos", (row, cursor) -> cursor.getCurrentTotal() * 10);
        w.column("Pos+", (row, cursor) -> cursor.getCurrentTotal() * 10, c -> c.bold(true));
        w.columnIf("On", true, Row::name);
        w.columnIf("Off", false, Row::name);
        w.columnIf("OnCfg", true, Row::amount, MONEY);
        w.columnIf("OnRow", true, (row, cursor) -> cursor.getCurrentTotal());
        w.columnIf("OffRow", false, (row, cursor) -> cursor.getCurrentTotal());
        w.constColumn("K", "k");
        w.constColumn("K+", "k2", c -> c.bold(true));
        w.constColumnIf("KIf", true, "k3");
        w.constColumnIf("KOff", false, "never");
        w.constColumnIf("KIf+", true, "k4", c -> c.bold(true));
    }

    private static List<List<String>> table(XSSFWorkbook wb, int firstRow) {
        var fmt = new DataFormatter();
        var rows = new ArrayList<List<String>>();
        var sheet = wb.getSheetAt(0);
        for (int r = firstRow; r <= sheet.getLastRowNum(); r++) {
            var cells = new ArrayList<String>();
            for (var cell : sheet.getRow(r)) cells.add(fmt.formatCellValue(cell));
            rows.add(cells);
        }
        return rows;
    }

    @Nested
    class SameAsASheetOfAWorkbook {
        @Test
        void everyColumnMethodGivesTheSameCellsInATemplateAndInAWorkbook() throws IOException {
            List<List<String>> viaTemplate;
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = w.<Row>list(0, 1);
                columns(list);
                list.writeWithHeaders(ROWS);
                var out = new ByteArrayOutputStream();
                w.finish().writeTo(out);
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    assertEquals("Title", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue(), "the template stays");
                    viaTemplate = table(wb, 1);
                }
            }
            List<List<String>> viaSheet;
            try (var workbook = ExcelWorkbook.create()) {
                var sheet = workbook.<Row>sheet("S");
                columns(sheet);
                sheet.write(ROWS);
                var out = new ByteArrayOutputStream();
                workbook.finish().writeTo(out);
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    viaSheet = table(wb, 0);
                }
            }
            assertEquals(viaSheet, viaTemplate);
            assertEquals(List.of("No.", "Name", "Amount", "Pos", "Pos+", "On", "OnCfg", "OnRow", "K", "K+", "KIf", "KIf+"), viaTemplate.get(0),
                    "false conditions add no column");
            assertEquals(List.of("1", "a", "1", "10", "10", "a", "1", "1", "k", "k2", "k3", "k4"), viaTemplate.get(1));
            assertEquals("3", viaTemplate.get(3).get(0));
        }

        @Test
        void aNullConfigurerIsNoSettings() throws IOException {
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = w.<Row>list(0, 1);
                list.column("A", Row::name, (Consumer<ColumnConfig<Row>>) null)
                        .column("B", (row, cursor) -> row.name(), null)
                        .columnIf("C", true, Row::name, null)
                        .constColumn("D", "d", null)
                        .constColumnIf("E", true, "e", null)
                        .writeWithHeaders(ROWS);
                var out = new ByteArrayOutputStream();
                w.finish().writeTo(out);
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    assertEquals(List.of("A", "B", "C", "D", "E"), table(wb, 1).get(0));
                }
            }
        }

        @Test
        void aListInATemplateIsAColumnsWriter() throws IOException {
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                ColumnsWriter<Row, TemplateListWriter<Row>> list = w.<Row>list(0, 1);
                assertNotNull(list);
            }
        }
    }

    @Nested
    class RowNumbers {
        @Test
        void startAtOneWhateverTheStartRowOfTheList() throws IOException {
            for (int start : new int[]{1, 4, 10}) {
                try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                    w.<Row>list(0, start).rowNumberColumn("No.").column("Name", Row::name).write(ROWS);
                    var out = new ByteArrayOutputStream();
                    w.finish().writeTo(out);
                    try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                        var sheet = wb.getSheetAt(0);
                        for (int i = 0; i < 3; i++) {
                            assertEquals(i + 1, (int) sheet.getRow(start + i).getCell(0).getNumericCellValue(), "start " + start);
                            assertEquals(ROWS.get(i).name(), sheet.getRow(start + i).getCell(1).getStringCellValue());
                        }
                    }
                }
            }
        }
    }

    @Nested
    class ErrorPolicy {
        private static final java.util.function.Function<Row, Object> FAILS_ON_B = r -> {
            if (r.name().equals("b")) throw new IllegalStateException("boom");
            return r.name();
        };

        @Test
        void failFastIsTheDefaultAndStopsAtTheFirstFailure() throws IOException {
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = w.<Row>list(0, 1).column("N", FAILS_ON_B);
                assertThrows(ExcelWriteException.class, () -> list.write(ROWS));
            }
        }

        @Test
        void lenientLeavesTheCellBlankAndGoesOn() throws IOException {
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                w.<Row>list(0, 1).writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT).column("N", FAILS_ON_B).write(ROWS);
                var out = new ByteArrayOutputStream();
                w.finish().writeTo(out);
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                    var sheet = wb.getSheetAt(0);
                    assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
                    var middle = sheet.getRow(2).getCell(0);
                    assertTrue(middle == null || middle.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
                    assertEquals("c", sheet.getRow(3).getCell(0).getStringCellValue());
                }
            }
        }

        @Test
        void thePolicyMustNotBeNull() throws IOException {
            try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = w.<Row>list(0, 1);
                assertThrows(NullPointerException.class, () -> list.writeErrorPolicy(null));
            }
        }
    }

    @Test
    void aMissingStreamIsRefusedBeforeAnythingElse() throws IOException {
        try (var w = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
            var list = w.<Row>list(0, 1);
            assertThrows(NullPointerException.class, () -> list.write((java.util.stream.Stream<Row>) null));
            assertThrows(NullPointerException.class, () -> list.writeWithHeaders((java.util.stream.Stream<Row>) null));
            assertThrows(NullPointerException.class, () -> list.write((Iterable<Row>) null));
        }
    }
}
