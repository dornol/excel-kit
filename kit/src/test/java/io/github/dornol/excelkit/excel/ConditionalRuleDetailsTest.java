package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelConditionalRule.IconSetType;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfvo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The conditional formatting XML a rule produces: ranges, thresholds, colors, priorities. */
class ConditionalRuleDetailsTest {

    record Row(String name, int a, int b, int c) {
    }

    private static XSSFSheet write(int rows, Consumer<ExcelConditionalRule> rule, Consumer<ExcelWriter<Row>> more) throws Exception {
        ExcelWriter<Row> writer = ExcelWriter.<Row>create().column("Name", Row::name).column("A", Row::a)
                .column("B", Row::b).column("C", Row::c).conditionalFormatting(rule);
        more.accept(writer);
        List<Row> data = new ArrayList<>();
        for (int i = 1; i <= rows; i++) data.add(new Row("n" + i, i, i * 2, i * 3));
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(data).writeTo(out);
        // the workbook must stay open while its sheet is read: the caller closes nothing, the JVM frees it
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray())).getSheetAt(0);
    }

    private static XSSFSheet write(int rows, Consumer<ExcelConditionalRule> rule) throws Exception {
        return write(rows, rule, w -> { });
    }

    private static List<CTConditionalFormatting> formattings(XSSFSheet sheet) {
        return List.of(sheet.getCTWorksheet().getConditionalFormattingArray());
    }

    private static String range(CTConditionalFormatting formatting) {
        return formatting.getSqref().get(0).toString();
    }

    // ---- data bars

    @Test
    void aDataBarCoversTheDataRowsOfItsColumnFromMinimumToMaximum() throws Exception {
        XSSFSheet sheet = write(4, rule -> rule.columns(1).dataBar(ExcelColor.BLUE));

        CTConditionalFormatting formatting = formattings(sheet).get(0);
        assertEquals("B2:B5", range(formatting));
        CTCfRule rule = formatting.getCfRuleArray(0);
        assertEquals(STCfType.DATA_BAR, rule.getType());
        assertEquals(2, rule.getDataBar().sizeOfCfvoArray());
        assertEquals(STCfvoType.MIN, rule.getDataBar().getCfvoArray(0).getType());
        assertEquals(STCfvoType.MAX, rule.getDataBar().getCfvoArray(1).getType());
        assertArrayEquals(new byte[]{(byte) 0xFF, 0, 0, (byte) 0xFF}, rule.getDataBar().getColor().getRgb());
    }

    // ---- icon sets

    @Test
    void anIconSetHasItsOwnTypeAndPercentThresholdsAfterTheMinimum() throws Exception {
        int[][] expected = {
                {3, 0, 33, 66}, {4, 0, 25, 50, 75}, {5, 0, 20, 40, 60, 80}};
        for (IconSetType type : IconSetType.values()) {
            XSSFSheet sheet = write(3, rule -> rule.columns(1).iconSet(type));
            CTCfRule rule = formattings(sheet).get(0).getCfRuleArray(0);
            assertEquals(STCfType.ICON_SET, rule.getType(), type.toString());
            CTCfvo[] thresholds = rule.getIconSet().getCfvoArray();
            int count = switch (type) {
                case ARROWS_3, TRAFFIC_LIGHTS_3, SIGNS_3, SYMBOLS_3, FLAGS_3 -> 3;
                case ARROWS_4, RATINGS_4 -> 4;
                case ARROWS_5, RATINGS_5, QUARTERS_5 -> 5;
            };
            assertEquals(count, thresholds.length, type.toString());
            int[] values = java.util.Arrays.stream(expected).filter(row -> row[0] == count).findFirst().orElseThrow();
            assertEquals(STCfvoType.MIN, thresholds[0].getType(), type + " starts at the minimum");
            for (int i = 1; i < count; i++) {
                assertEquals(STCfvoType.PERCENT, thresholds[i].getType(), type + " threshold " + i);
                assertEquals(String.valueOf(values[i + 1]), thresholds[i].getVal(), type + " threshold " + i);
            }
        }
    }

    @Test
    void everyIconSetTypeIsAskedForByName() throws Exception {
        List<String> names = new ArrayList<>();
        for (IconSetType type : IconSetType.values()) {
            XSSFSheet sheet = write(2, rule -> rule.columns(1).iconSet(type));
            names.add(formattings(sheet).get(0).getCfRuleArray(0).getIconSet().getIconSet().toString());
        }
        assertEquals(List.of("3Arrows", "4Arrows", "5Arrows", "3TrafficLights1", "3Signs", "3Symbols", "3Flags",
                "4Rating", "5Rating", "5Quarters"), names);
    }

    // ---- comparisons

    @Test
    void aComparisonRuleKeepsItsOperatorAndValues() throws Exception {
        XSSFSheet sheet = write(3, rule -> rule.columns(1)
                .greaterThan("5", ExcelColor.LIGHT_GREEN).between("1", "2", ExcelColor.LIGHT_RED)
                .lessThanOrEqual("9", ExcelColor.LIGHT_YELLOW));

        List<ConditionalFormattingRule> rules = new ArrayList<>();
        var formatting = sheet.getSheetConditionalFormatting();
        for (int i = 0; i < formatting.getNumConditionalFormattings(); i++) {
            rules.add(formatting.getConditionalFormattingAt(i).getRule(0));
        }
        assertEquals(3, rules.size());
        assertEquals(ComparisonOperator.GT, rules.get(0).getComparisonOperation());
        assertEquals("5", rules.get(0).getFormula1());
        assertEquals(ComparisonOperator.BETWEEN, rules.get(1).getComparisonOperation());
        assertEquals("1", rules.get(1).getFormula1());
        assertEquals("2", rules.get(1).getFormula2());
        assertEquals(ComparisonOperator.LE, rules.get(2).getComparisonOperation());
        assertEquals("9", rules.get(2).getFormula1());
    }

    // ---- where it applies

    @Test
    void byDefaultEveryColumnIsFormattedAndAnExplicitListOnlyThose() throws Exception {
        List<String> all = formattings(write(3, rule -> rule.dataBar(ExcelColor.BLUE))).stream()
                .map(ConditionalRuleDetailsTest::range).toList();
        assertEquals(List.of("A2:A4", "B2:B4", "C2:C4", "D2:D4"), all);

        List<String> some = formattings(write(3, rule -> rule.columns(3, 1).dataBar(ExcelColor.BLUE))).stream()
                .map(ConditionalRuleDetailsTest::range).toList();
        assertEquals(List.of("D2:D4", "B2:B4"), some);
    }

    @Test
    void theRangeStartsAfterTheHeaderOrWhereAskedAndNeverBeforeTheFirstRow() throws Exception {
        assertEquals("B2:B4", range(formattings(write(3, rule -> rule.columns(1).dataBar(ExcelColor.BLUE))).get(0)));
        assertEquals("B3:B4", range(formattings(write(3, rule -> rule.columns(1).startRow(2).dataBar(ExcelColor.BLUE))).get(0)));
        assertEquals("B1:B4", range(formattings(write(3, rule -> rule.columns(1).startRow(0).dataBar(ExcelColor.BLUE))).get(0)));
        // a start below the last row leaves a one-row range, not an upside-down one
        assertEquals("B9", range(formattings(write(3, rule -> rule.columns(1).startRow(8).dataBar(ExcelColor.BLUE))).get(0)));
    }

    @Test
    void rowsBeneathTheDataAreNotFormatted() throws Exception {
        XSSFSheet sheet = write(3, rule -> rule.columns(1).dataBar(ExcelColor.BLUE),
                w -> w.summary(s -> s.sum("A")));
        assertEquals("B2:B4", range(formattings(sheet).get(0)));
    }

    @Test
    void everyRuleGetsAPriorityOfItsOwnInOrder() throws Exception {
        XSSFSheet sheet = write(3, rule -> rule.columns(1, 2).dataBar(ExcelColor.BLUE).iconSet(IconSetType.ARROWS_3));
        List<Long> priorities = new ArrayList<>();
        for (CTConditionalFormatting formatting : formattings(sheet)) {
            for (CTCfRule rule : formatting.getCfRuleArray()) priorities.add((long) rule.getPriority());
        }
        assertEquals(4, priorities.size());
        assertEquals(List.of(1L, 2L, 3L, 4L), priorities);
        assertFalse(priorities.isEmpty());
        assertTrue(priorities.stream().distinct().count() == priorities.size());
    }

    // ---- the XML is valid

    private static void assertValid(XSSFSheet sheet, String what) {
        List<org.apache.xmlbeans.XmlError> errors = new ArrayList<>();
        boolean valid = sheet.getCTWorksheet().validate(new org.apache.xmlbeans.XmlOptions().setErrorListener(errors));
        assertTrue(valid, what + " is not valid SpreadsheetML: " + errors);
    }

    @Test
    void everyKindOfRuleWritesValidSpreadsheetML() throws Exception {
        assertValid(write(3, rule -> rule.dataBar(ExcelColor.BLUE)), "a data bar");
        for (IconSetType type : IconSetType.values()) {
            assertValid(write(3, rule -> rule.iconSet(type)), "icon set " + type);
        }
        assertValid(write(3, rule -> rule.greaterThan("1", ExcelColor.LIGHT_GREEN).lessThan("9", ExcelColor.LIGHT_RED)
                .between("2", "3", ExcelColor.LIGHT_YELLOW).notBetween("4", "5", ExcelColor.LIGHT_RED)
                .equalTo("6", ExcelColor.LIGHT_GREEN).notEqualTo("7", ExcelColor.LIGHT_RED)
                .greaterThanOrEqual("8", ExcelColor.LIGHT_GREEN).lessThanOrEqual("9", ExcelColor.LIGHT_RED)), "comparisons");
        assertValid(write(3, rule -> rule.dataBar(ExcelColor.BLUE).iconSet(IconSetType.ARROWS_3)
                .greaterThan("1", ExcelColor.LIGHT_GREEN)), "all of them together");
    }
}
