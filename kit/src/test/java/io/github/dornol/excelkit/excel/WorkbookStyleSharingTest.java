package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

/**
 * The sheets of a workbook that look alike share their cell styles: a workbook of many sheets must not hold one more
 * style per sheet for a header color, a stripe, a row color, a subtotal row or a type inferred from the data.
 */
class WorkbookStyleSharingTest {
    record Row(String name, int amount, long big, double ratio, LocalDate day, BigDecimal money) {}

    private static final List<Row> ROWS = List.of(
            new Row("a", 1, 10L, 0.1, LocalDate.of(2026, 1, 1), BigDecimal.ONE),
            new Row("a", 2, 20L, 0.2, LocalDate.of(2026, 1, 2), BigDecimal.TEN),
            new Row("b", 3, 30L, 0.3, LocalDate.of(2026, 1, 3), BigDecimal.TEN));

    /** A sheet of untyped columns: their numeric and date formats are inferred from the values. */
    private static ExcelSheetWriter<Row> columns(ExcelSheetWriter<Row> sheet) {
        return sheet.column("Name", Row::name).column("Amount", Row::amount).column("Big", Row::big)
                .column("Ratio", Row::ratio).column("Day", Row::day).column("Money", Row::money);
    }

    private static int styles(int sheets, Consumer<ExcelSheetWriter<Row>> configure) throws IOException {
        var out = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            for (int i = 0; i < sheets; i++) {
                var sheet = columns(workbook.<Row>sheet("S" + i));
                configure.accept(sheet);
                sheet.write(ROWS);
            }
            workbook.finish().writeTo(out);
        }
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return wb.getNumCellStyles() * 1000 + wb.getNumberOfFonts();
        }
    }

    private static Map<String, Consumer<ExcelSheetWriter<Row>>> features() {
        Map<String, Consumer<ExcelSheetWriter<Row>>> f = new LinkedHashMap<>();
        f.put("nothing", s -> {});
        f.put("stripedRows", s -> s.stripedRows());
        f.put("rowColor", s -> s.rowColor(r -> r.amount() > 1 ? ExcelColor.LIGHT_YELLOW : null));
        f.put("rowStyle", s -> s.rowStyle(r -> r.amount() > 1, st -> st.bold(true).italic(true)));
        f.put("groupSubtotal", s -> s.groupSubtotal(Row::name, g -> g.label("Name", k -> "t").sum("Amount")));
        f.put("groupSubtotal with keyColumn", s -> s.groupSubtotal(Row::name, g -> g.label("Big", k -> "t").keyColumn("Name").sum("Amount")));
        f.put("emptyMessage", s -> s.emptyMessage("none"));
        f.put("summary", s -> s.summary(x -> x.label("Total").sum("Amount")));
        f.put("protectSheet", s -> s.protectSheet("pw"));
        f.put("defaultStyle", s -> s.defaultStyle(d -> d.bold(true)));
        return f;
    }

    @Nested
    class Sheets {
        @Test
        void aSheetFeatureAddsNoStylesWithTheNumberOfSheets() throws IOException {
            for (var feature : features().entrySet()) {
                int one = styles(1, feature.getValue());
                int twenty = styles(20, feature.getValue());
                assertEquals(one, twenty, feature.getKey() + ": 20 sheets hold as many styles and fonts as 1");
            }
        }

        @Test
        void headerLookAndRequiredMarksAddNoStylesWithTheNumberOfSheets() throws IOException {
            assertEquals(headerLook(1), headerLook(20));
        }

        private int headerLook(int sheets) throws IOException {
            var out = new ByteArrayOutputStream();
            try (var workbook = ExcelWorkbook.create()) {
                for (int i = 0; i < sheets; i++) {
                    var sheet = workbook.<Row>sheet("S" + i)
                            .column("Name", Row::name, c -> c.requiredMark().headerBackgroundColor(ExcelColor.LIGHT_BLUE)
                                    .headerFontColor(ExcelColor.WHITE))
                            .column("Amount", Row::amount, c -> c.cellColor((v, r) -> ExcelColor.LIGHT_GREEN));
                    sheet.write(ROWS);
                }
                workbook.finish().writeTo(out);
            }
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                return wb.getNumCellStyles() * 1000 + wb.getNumberOfFonts();
            }
        }
    }

    @Nested
    class SharedLooksStayCorrect {
        @Test
        void everySheetStillGetsItsOwnStripesColorsAndSubtotals() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var workbook = ExcelWorkbook.create()) {
                columns(workbook.<Row>sheet("Plain")).write(ROWS);
                columns(workbook.<Row>sheet("Striped")).stripedRows().write(ROWS);
                columns(workbook.<Row>sheet("Subtotal")).groupSubtotal(Row::name, g -> g.label("Name", k -> "t").sum("Amount")).write(ROWS);
                columns(workbook.<Row>sheet("Striped2")).stripedRows().write(ROWS);
                workbook.finish().writeTo(out);
            }
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var plain = (org.apache.poi.xssf.usermodel.XSSFCellStyle) wb.getSheetAt(0).getRow(2).getCell(1).getCellStyle();
                var striped = (org.apache.poi.xssf.usermodel.XSSFCellStyle) wb.getSheetAt(1).getRow(2).getCell(1).getCellStyle();
                var striped2 = (org.apache.poi.xssf.usermodel.XSSFCellStyle) wb.getSheetAt(3).getRow(2).getCell(1).getCellStyle();
                assertEquals(org.apache.poi.ss.usermodel.FillPatternType.NO_FILL, plain.getFillPattern(), "a plain sheet has no stripe");
                assertEquals(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND, striped.getFillPattern());
                assertEquals(striped.getIndex(), striped2.getIndex(), "two striped sheets share the style");
                var subtotal = wb.getSheetAt(2).getRow(3).getCell(1);
                assertEquals("SUBTOTAL(9,B2:B3)", subtotal.getCellFormula());
                assertEquals(org.apache.poi.ss.usermodel.FillPatternType.NO_FILL,
                        wb.getSheetAt(0).getRow(3).getCell(1).getCellStyle().getFillPattern(), "the plain sheet stays plain");
            }
        }

        @Test
        void aTypeInferredFromTheDataKeepsItsFormatOnEverySheet() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var workbook = ExcelWorkbook.create()) {
                for (int i = 0; i < 3; i++) columns(workbook.<Row>sheet("S" + i)).write(ROWS);
                workbook.finish().writeTo(out);
            }
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                for (int s = 0; s < 3; s++) {
                    var row = wb.getSheetAt(s).getRow(1);
                    assertTrue(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(row.getCell(4)), "a date shows as a date: sheet " + s);
                    assertNotEquals(0, row.getCell(3).getCellStyle().getDataFormat(), "a double has a number format: sheet " + s);
                    assertNotEquals(0, row.getCell(4).getCellStyle().getDataFormat(), "a date has a date format: sheet " + s);
                }
            }
        }

        @Test
        void untypedColumnsOfTheSameKindShareOneStyleInASheet() throws IOException {
            record Wide(double a, double b, double c, double d, double e) {}
            int one;
            int five;
            var out1 = new ByteArrayOutputStream();
            ExcelWriter.<Wide>create().column("A", Wide::a).write(List.of(new Wide(1, 2, 3, 4, 5))).writeTo(out1);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out1.toByteArray()))) { one = wb.getNumCellStyles(); }
            var out5 = new ByteArrayOutputStream();
            ExcelWriter.<Wide>create().column("A", Wide::a).column("B", Wide::b).column("C", Wide::c).column("D", Wide::d)
                    .column("E", Wide::e).write(List.of(new Wide(1, 2, 3, 4, 5))).writeTo(out5);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out5.toByteArray()))) { five = wb.getNumCellStyles(); }
            assertEquals(one, five, "five columns of doubles need no more styles than one");
        }
    }
}
