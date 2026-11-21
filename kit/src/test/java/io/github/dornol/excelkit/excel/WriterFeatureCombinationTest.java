package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Every writer feature is checked on its own and together with every other one, by writing a small
 * workbook and reading it back. A feature's check must hold whichever other features are on: the bugs
 * this guards against were interactions, such as a row style dropping a column's font color or a preamble
 * using up the auto-width sample, which no test of either feature alone could see.
 * <p>
 * To cover a new feature, add a {@link Feature}: how to switch it on, and what the written sheet must
 * show. The checks locate rows by content, so a preamble or a group header does not disturb them.
 */
class WriterFeatureCombinationTest {

    record Row(String name, int score) {
    }

    /** One sheet as written: where its header and data rows ended up. */
    private record Layout(XSSFSheet sheet, int header, List<Integer> dataRows) {
        XSSFRow data(int index) {
            return sheet.getRow(dataRows.get(index));
        }

        int firstData() {
            return dataRows.get(0);
        }

        int lastData() {
            return dataRows.get(dataRows.size() - 1);
        }
    }

    /** What the features need of a writer, which {@link ExcelWriter} and {@link ExcelSheetWriter} both offer. */
    private interface WriterApi {
        AbstractSheetWriter<Row, ?> base();

        void maxRows(int maxRows);

        void table(String name);
    }

    private interface Check {
        void verify(XSSFWorkbook workbook, List<Layout> sheets);
    }

    /**
     * A feature switched on by configuring the writer, and/or one column ({@code nameColumn} or
     * {@code scoreColumn}). Columns are declared once, with every feature's configuration applied.
     */
    private record Feature(String name, Consumer<WriterApi> apply,
                           Consumer<ColumnStyleConfig<Row, ?>> nameColumn, Consumer<ColumnStyleConfig<Row, ?>> scoreColumn,
                           Check check) {
        static final Consumer<ColumnStyleConfig<Row, ?>> NONE = c -> { };

        static Feature writer(String name, Consumer<WriterApi> apply, Check check) {
            return new Feature(name, apply, NONE, NONE, check);
        }

        static Feature name(String name, Consumer<ColumnStyleConfig<Row, ?>> column, Check check) {
            return new Feature(name, w -> { }, column, NONE, check);
        }

        static Feature score(String name, Consumer<ColumnStyleConfig<Row, ?>> column, Check check) {
            return new Feature(name, w -> { }, NONE, column, check);
        }
    }

    private static final byte[] RED = {(byte) 255, 0, 0};
    private static final String LONG_NAME = "n3-" + "x".repeat(30);

    private static List<Row> rows() {
        return List.of(new Row("n1", 1), new Row("n2", 2), new Row(LONG_NAME, 3), new Row("n4", 4), new Row("n5", 5));
    }

    // ---- features

    private static final List<Feature> FEATURES = List.of(
            Feature.writer("preamble",
                    w -> w.base().beforeHeader(ctx -> {
                        for (int i = 0; i < 3; i++) ctx.getSheet().createRow(i).createCell(0).setCellValue("pre" + i);
                        return 3;
                    }),
                    (wb, sheets) -> sheets.forEach(s -> {
                        for (int i = 0; i < 3; i++) expect("pre" + i, text(s.sheet.getRow(i), 0), "preamble row " + i);
                    })),
            new Feature("group", w -> { }, c -> c.group("G"), c -> c.group("G"),
                    (wb, sheets) -> sheets.forEach(s ->
                            expect("G", text(s.sheet.getRow(s.header - 1), 0), "group header above the header row"))),
            Feature.writer("rowStyleBold",
                    w -> w.base().rowStyle(r -> r.score() >= 3, s -> s.bold(true)),
                    (wb, sheets) -> sheets.forEach(s -> eachData(s, (row, score) -> {
                        if (score >= 3) assertTrue(font(row.getCell(0)).getBold(), "bold row style, score " + score);
                    }))),
            Feature.score("columnFont", c -> c.fontColor(ExcelColor.RED).underline(true),
                    (wb, sheets) -> sheets.forEach(s -> eachData(s, (row, score) -> {
                        XSSFFont font = font(row.getCell(1));
                        assertTrue(font.getXSSFColor() != null && Arrays.equals(RED, font.getXSSFColor().getRGB()),
                                "column font color, score " + score);
                        assertTrue(font.getUnderline() == Font.U_SINGLE, "column underline, score " + score);
                    }))),
            Feature.writer("rowColor",
                    w -> w.base().rowColor(r -> r.score() % 2 == 0 ? ExcelColor.LIGHT_YELLOW : null),
                    (wb, sheets) -> sheets.forEach(s -> eachData(s, (row, score) -> {
                        if (score % 2 == 0) expectFill(row.getCell(0), ExcelColor.LIGHT_YELLOW, "row color, score " + score);
                    }))),
            Feature.score("cellColor", c -> c.cellColor((v, r) -> r.score() >= 4 ? ExcelColor.LIGHT_RED : null),
                    (wb, sheets) -> sheets.forEach(s -> eachData(s, (row, score) -> {
                        if (score >= 4) expectFill(row.getCell(1), ExcelColor.LIGHT_RED, "cell color, score " + score);
                    }))),
            Feature.writer("autoFilter",
                    w -> w.base().autoFilter(true),
                    (wb, sheets) -> sheets.forEach(s -> {
                        String ref = s.sheet.getCTWorksheet().getAutoFilter().getRef();
                        assertTrue(ref.startsWith("A" + (s.header + 1) + ":"), "auto filter " + ref + " starts at header row " + (s.header + 1));
                    })),
            Feature.writer("summary",
                    w -> w.base().summary(s -> s.sum("Score")),
                    (wb, sheets) -> sheets.forEach(s -> {
                        String expected = "SUM(B" + (s.firstData() + 1) + ":B" + (s.lastData() + 1) + ")";
                        boolean found = false;
                        for (int i = s.lastData() + 1; i <= s.sheet.getLastRowNum(); i++) {
                            XSSFRow row = s.sheet.getRow(i);
                            if (row == null) continue;
                            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                                if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA
                                        && expected.equals(cell.getCellFormula())) found = true;
                            }
                        }
                        assertTrue(found, "summary formula " + expected + " below the data");
                    })),
            Feature.writer("afterData",
                    w -> w.base().afterData(ctx -> {
                        ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("END");
                        return ctx.getCurrentRow() + 1;
                    }),
                    (wb, sheets) -> sheets.forEach(s -> {
                        boolean found = false;
                        for (int i = s.lastData() + 1; i <= s.sheet.getLastRowNum(); i++) {
                            if ("END".equals(text(s.sheet.getRow(i), 0))) found = true;
                        }
                        assertTrue(found, "afterData row below the data");
                    })),
            Feature.writer("autoWidth",
                    w -> { },
                    (wb, sheets) -> sheets.forEach(s -> assertTrue(s.sheet.getColumnWidth(0) >= 8000,
                            "the 33-character name widens column A, width " + s.sheet.getColumnWidth(0)))),
            Feature.writer("rollover",
                    w -> w.maxRows(3),
                    (wb, sheets) -> {
                        assertTrue(sheets.size() == 2, "5 rows with maxRows(3) fill 2 sheets, got " + sheets.size());
                        assertTrue(sheets.get(0).dataRows.size() == 3 && sheets.get(1).dataRows.size() == 2,
                                "3 + 2 data rows, got " + sheets.get(0).dataRows.size() + " + " + sheets.get(1).dataRows.size());
                    }),
            Feature.writer("table",
                    w -> w.table("Scores_T"),
                    (wb, sheets) -> sheets.forEach(s -> {
                        assertTrue(s.sheet.getTables().size() == 1, "one table per sheet");
                        assertTrue(s.sheet.getTables().get(0).getStartCellReference().getRow() == s.header,
                                "the table starts at the header row " + s.header);
                    })),
            Feature.score("dataFormat", c -> c.format("#,##0"),
                    (wb, sheets) -> sheets.forEach(s -> eachData(s, (row, score) ->
                            expect("#,##0", row.getCell(1).getCellStyle().getDataFormatString(), "number format, score " + score)))),
            Feature.writer("protectSheet",
                    w -> w.base().protectSheet("pw"),
                    (wb, sheets) -> sheets.forEach(s ->
                            assertTrue(s.sheet.getCTWorksheet().getSheetProtection().getSheet(), "sheet is protected"))),
            Feature.writer("conditionalFormatting",
                    w -> w.base().conditionalFormatting(r -> r.columns(1).greaterThan("3", ExcelColor.LIGHT_GREEN)),
                    (wb, sheets) -> sheets.forEach(s -> {
                        var formatting = s.sheet.getSheetConditionalFormatting();
                        assertTrue(formatting.getNumConditionalFormattings() >= 1, "a conditional format");
                        CellRangeAddress range = formatting.getConditionalFormattingAt(0).getFormattingRanges()[0];
                        assertTrue(range.getFirstRow() == s.firstData(),
                                "conditional format starts at the first data row " + s.firstData() + ", was " + range.getFirstRow());
                    })),
            Feature.name("comment", c -> c.comment(r -> "c-" + r.name()),
                    (wb, sheets) -> sheets.forEach(s -> {
                        for (int i = 0; i < s.dataRows.size(); i++) {
                            Comment comment = s.data(i).getCell(0).getCellComment();
                            assertTrue(comment != null, "a comment on data row " + i);
                            expect("c-" + text(s.data(i), 0), comment.getString().getString(), "comment text");
                        }
                    })),
            Feature.writer("tabColor",
                    w -> w.base().tabColor(10, 20, 30),
                    (wb, sheets) -> sheets.forEach(s -> assertTrue(
                            Arrays.equals(new byte[]{10, 20, 30}, s.sheet.getTabColor().getRGB()), "tab color"))),
            Feature.writer("namedRange",
                    w -> w.base().namedRange("Scores", 1),
                    (wb, sheets) -> assertTrue(wb.getName("Scores") != null, "named range Scores exists")),
            Feature.writer("freeze",
                    w -> w.base().freezeRows(1),
                    (wb, sheets) -> sheets.forEach(s -> {
                        var pane = s.sheet.getPaneInformation();
                        assertTrue(pane != null && pane.getHorizontalSplitPosition() == s.header + 1,
                                "the panes freeze the rows down to the header row " + s.header);
                    })),
            Feature.writer("printRepeat",
                    w -> w.base().printSetup(p -> p.repeatHeaderRows()),
                    (wb, sheets) -> sheets.forEach(s ->
                            assertTrue(s.sheet.getRepeatingRows() != null && s.sheet.getRepeatingRows().getLastRow() == s.header,
                                    "the repeated print rows end at the header row " + s.header))),
            Feature.writer("chart",
                    w -> w.base().chart(c -> c.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "S")),
                    (wb, sheets) -> {
                        Layout last = sheets.get(sheets.size() - 1);
                        var charts = last.sheet.getDrawingPatriarch().getCharts();
                        assertTrue(charts.size() == 1, "one chart on the last sheet");
                        String values = charts.get(0).getCTChart().getPlotArea().getBarChartArray(0)
                                .getSerArray(0).getVal().getNumRef().getF();
                        assertTrue(values.contains("$B$" + (last.firstData() + 1) + ":$B$" + (last.lastData() + 1)),
                                "the series covers the data rows B" + (last.firstData() + 1) + ":B" + (last.lastData() + 1)
                                        + ", was " + values);
                    }),
            Feature.score("validation", c -> c.validation(ExcelValidation.integerBetween(0, 10)),
                    (wb, sheets) -> sheets.forEach(s -> {
                        var found = s.sheet.getDataValidations().stream()
                                .filter(v -> v.getRegions().getCellRangeAddress(0).getFirstColumn() == 1).findFirst();
                        assertTrue(found.isPresent(), "a validation on the Score column");
                        assertTrue(found.get().getRegions().getCellRangeAddress(0).getFirstRow() == s.firstData(),
                                "the validation starts at the first data row " + s.firstData());
                    })),
            Feature.name("dropdown", c -> c.dropdown("n1", "n2", "n4", "n5"),
                    (wb, sheets) -> sheets.forEach(s -> {
                        var validations = s.sheet.getDataValidations();
                        assertTrue(!validations.isEmpty(), "a dropdown validation");
                        assertTrue(validations.get(0).getRegions().getCellRangeAddress(0).getFirstRow() == s.firstData(),
                                "dropdown starts at the first data row " + s.firstData());
                    })));

    // ---- running

    @Test
    void everyFeatureHoldsAloneAndWithEveryOtherFeature() throws Exception {
        List<String> failures = new ArrayList<>();
        for (int i = 0; i < FEATURES.size(); i++) {
            run(failures, FEATURES.get(i));
            for (int j = i + 1; j < FEATURES.size(); j++) run(failures, FEATURES.get(i), FEATURES.get(j));
        }
        run(failures, FEATURES.toArray(Feature[]::new));
        assertTrue(failures.isEmpty(), failures.size() + " feature combination problem(s):\n  "
                + String.join("\n  ", failures));
    }

    private static void run(List<String> failures, Feature... features) throws Exception {
        String names = String.join(" + ", Arrays.stream(features).map(Feature::name).toList());
        for (String writer : List.of("ExcelWriter", "ExcelSheetWriter")) {
            String label = writer + ": " + names;
            byte[] file;
            try {
                file = writer.equals("ExcelWriter") ? writeWithExcelWriter(features) : writeWithSheetWriter(features);
            } catch (RuntimeException e) {
                failures.add("[" + label + "] failed to write: " + e);
                continue;
            }
            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                List<Layout> sheets = layouts(workbook);
                schemaProblems(workbook).forEach(problem -> failures.add("[" + label + "] " + problem));
                for (Feature feature : features) {
                    try {
                        feature.check().verify(workbook, sheets);
                    } catch (AssertionError | RuntimeException e) {
                        failures.add("[" + label + "] " + feature.name() + ": " + e.getMessage());
                    }
                }
            }
        }
    }

    private static byte[] writeWithExcelWriter(Feature... features) {
        ExcelWriter<Row> writer = ExcelWriter.create();
        writer.column("Name", Row::name, c -> {
            for (Feature feature : features) feature.nameColumn().accept(c);
        });
        writer.column("Score", Row::score, c -> {
            for (Feature feature : features) feature.scoreColumn().accept(c);
        });
        WriterApi api = new WriterApi() {
            @Override public AbstractSheetWriter<Row, ?> base() { return writer; }
            @Override public void maxRows(int maxRows) { writer.maxRows(maxRows); }
            @Override public void table(String name) { writer.table(name); }
        };
        for (Feature feature : features) feature.apply().accept(api);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(rows()).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] writeWithSheetWriter(Feature... features) {
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            ExcelSheetWriter<Row> writer = workbook.sheet("Sheet");
            writer.column("Name", Row::name, c -> {
                for (Feature feature : features) feature.nameColumn().accept(c);
            });
            writer.column("Score", Row::score, c -> {
                for (Feature feature : features) feature.scoreColumn().accept(c);
            });
            WriterApi api = new WriterApi() {
                @Override public AbstractSheetWriter<Row, ?> base() { return writer; }
                @Override public void maxRows(int maxRows) { writer.maxRows(maxRows); }
                @Override public void table(String name) { writer.table(name); }
            };
            for (Feature feature : features) feature.apply().accept(api);
            writer.write(rows());
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.finish().writeTo(out);
            return out.toByteArray();
        }
    }

    // ---- the XML itself

    /**
     * The parts of the workbook that must be valid SpreadsheetML: its worksheets, its styles and the workbook
     * part. (Tables, charts and drawings need schemas that POI's lite artifact does not carry.)
     * <p>
     * POI writes a color given as three bytes as a six digit {@code rgb}, where the schema asks for eight
     * (alpha first). Excel reads the short form, as it has for every program that goes through POI's
     * {@code new XSSFColor(byte[3])}, so that one error is not counted.
     */
    private static List<String> schemaProblems(XSSFWorkbook workbook) {
        java.util.Map<String, org.apache.xmlbeans.XmlObject> parts = new java.util.LinkedHashMap<>();
        parts.put("the styles", workbook.getStylesSource().getCTStylesheet());
        parts.put("the workbook", workbook.getCTWorkbook());
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            parts.put("sheet " + i, workbook.getSheetAt(i).getCTWorksheet());
        }
        List<String> problems = new ArrayList<>();
        parts.forEach((name, xml) -> {
            List<org.apache.xmlbeans.XmlError> errors = new ArrayList<>();
            xml.validate(new org.apache.xmlbeans.XmlOptions().setErrorListener(errors));
            for (org.apache.xmlbeans.XmlError error : errors) {
                if (!String.valueOf(error.getMessage()).contains("ST_UnsignedIntHex")) {
                    problems.add(name + " is not valid SpreadsheetML: " + error.getMessage());
                }
            }
        });
        return problems;
    }

    // ---- reading back

    private static List<Layout> layouts(XSSFWorkbook workbook) {
        List<Layout> layouts = new ArrayList<>();
        for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
            XSSFSheet sheet = workbook.getSheetAt(s);
            int header = -1;
            List<Integer> data = new ArrayList<>();
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                String first = text(sheet.getRow(r), 0);
                if (header < 0 && "Name".equals(first)) header = r;
                else if (header >= 0 && first != null && first.matches("n\\d.*")) data.add(r);
            }
            if (header >= 0 && !data.isEmpty()) layouts.add(new Layout(sheet, header, data));
        }
        return layouts;
    }

    private static void eachData(Layout layout, java.util.function.BiConsumer<XSSFRow, Integer> action) {
        for (int index = 0; index < layout.dataRows.size(); index++) {
            XSSFRow row = layout.data(index);
            action.accept(row, (int) row.getCell(1).getNumericCellValue());
        }
    }

    private static String text(XSSFRow row, int column) {
        if (row == null || row.getCell(column) == null) return null;
        XSSFCell cell = row.getCell(column);
        return cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING ? cell.getStringCellValue() : null;
    }

    private static XSSFFont font(XSSFCell cell) {
        return cell.getCellStyle().getFont();
    }

    private static void expect(Object expected, Object actual, String what) {
        assertTrue(java.util.Objects.equals(expected, actual), what + ": expected <" + expected + "> but was <" + actual + ">");
    }

    private static void expectFill(XSSFCell cell, ExcelColor color, String what) {
        CellStyle style = cell.getCellStyle();
        XSSFColor fill = (XSSFColor) style.getFillForegroundColorColor();
        assertTrue(style.getFillPattern() == FillPatternType.SOLID_FOREGROUND && fill != null
                        && Arrays.equals(new byte[]{(byte) color.getR(), (byte) color.getG(), (byte) color.getB()}, fill.getRGB()),
                what + ": expected fill " + color.getR() + "," + color.getG() + "," + color.getB());
    }
}
