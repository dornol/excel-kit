package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for data bar and icon set conditional formatting — verifies
 * actual CT XML structures in the generated Excel file.
 */
class DataBarIconSetTest {

    record Item(String name, int value) {}

    private static Stream<Item> testData() {
        return Stream.of(
                new Item("A", 10), new Item("B", 50),
                new Item("C", 80), new Item("D", 30), new Item("E", 90));
    }

    private CTWorksheet writeAndGetWorksheet(java.util.function.Consumer<ExcelConditionalRule> cfConfig) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cfConfig)
                .write(testData())
                .consumeOutputStream(out);

        var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
        return wb.getSheetAt(0).getCTWorksheet();
    }

    // ============================================================
    // Data Bar
    // ============================================================
    @Test
    void dataBar_shouldCreateDataBarRule() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1).dataBar(ExcelColor.BLUE));

        boolean found = false;
        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.DATA_BAR) {
                    found = true;
                    assertTrue(rule.isSetDataBar(), "Rule should have data bar config");
                    assertNotNull(rule.getDataBar().getColor(), "Data bar should have color");
                    assertEquals(2, rule.getDataBar().sizeOfCfvoArray(), "Should have min and max thresholds");
                }
            }
        }
        assertTrue(found, "Should contain a DATA_BAR conditional formatting rule");
    }

    @Test
    void dataBar_colorShouldMatchInput() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1).dataBar(ExcelColor.GREEN));

        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.DATA_BAR) {
                    byte[] rgb = rule.getDataBar().getColor().getRgb();
                    // ARGB: [0xFF, R, G, B]
                    assertEquals((byte) 0x00, rgb[1], "Red component should be 0 for GREEN");
                    assertEquals((byte) 0x80, rgb[2], "Green component should be 128 for GREEN");
                    assertEquals((byte) 0x00, rgb[3], "Blue component should be 0 for GREEN");
                }
            }
        }
    }

    // ============================================================
    // Icon Set
    // ============================================================
    @Test
    void iconSet_shouldCreateIconSetRule() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .iconSet(ExcelConditionalRule.IconSetType.ARROWS_3));

        boolean found = false;
        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.ICON_SET) {
                    found = true;
                    assertTrue(rule.isSetIconSet(), "Rule should have icon set config");
                    assertEquals(3, rule.getIconSet().sizeOfCfvoArray(), "ARROWS_3 should have 3 thresholds");
                }
            }
        }
        assertTrue(found, "Should contain an ICON_SET conditional formatting rule");
    }

    @ParameterizedTest
    @EnumSource(ExcelConditionalRule.IconSetType.class)
    void iconSet_allTypes_shouldCreateValidRules(ExcelConditionalRule.IconSetType type) throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1).iconSet(type));

        boolean found = false;
        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.ICON_SET) {
                    found = true;
                    assertTrue(rule.getIconSet().sizeOfCfvoArray() >= 3,
                            type + " should have at least 3 thresholds");
                }
            }
        }
        assertTrue(found, type + " should create ICON_SET rule");
    }

    @Test
    void iconSet_5type_shouldHave5Thresholds() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .iconSet(ExcelConditionalRule.IconSetType.ARROWS_5));

        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.ICON_SET) {
                    assertEquals(5, rule.getIconSet().sizeOfCfvoArray());
                }
            }
        }
    }

    @Test
    void iconSet_4type_shouldHave4Thresholds() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .iconSet(ExcelConditionalRule.IconSetType.RATINGS_4));

        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.ICON_SET) {
                    assertEquals(4, rule.getIconSet().sizeOfCfvoArray());
                }
            }
        }
    }

    // ============================================================
    // 2-color gradient data bar
    // ============================================================
    @Test
    void dataBar_twoColor_shouldSetGradientLengthsAndColor() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .dataBar(ExcelColor.RED, ExcelColor.GREEN));

        boolean found = false;
        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.DATA_BAR) {
                    found = true;
                    var db = rule.getDataBar();

                    // Color should be the minColor (RED)
                    byte[] rgb = db.getColor().getRgb();
                    assertEquals((byte) 0xFF, rgb[1], "Red component should be 255 for RED");
                    assertEquals((byte) 0x00, rgb[2], "Green component should be 0 for RED");
                    assertEquals((byte) 0x00, rgb[3], "Blue component should be 0 for RED");

                    // Gradient: minLength=0, maxLength=100
                    assertEquals(0, db.getMinLength(), "Gradient should set minLength=0");
                    assertEquals(100, db.getMaxLength(), "Gradient should set maxLength=100");

                    // 2 thresholds (min, max)
                    assertEquals(2, db.sizeOfCfvoArray());
                }
            }
        }
        assertTrue(found, "Should have DATA_BAR rule");
    }

    @Test
    void dataBar_singleColor_shouldNotOverrideLengths() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .dataBar(ExcelColor.BLUE));

        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.DATA_BAR) {
                    var db = rule.getDataBar();
                    // Single color should NOT set 0/100 (that's the gradient marker)
                    assertFalse(db.getMinLength() == 0 && db.getMaxLength() == 100,
                            "Single color should not have gradient lengths (0/100)");
                }
            }
        }
    }

    @Test
    void dataBar_twoColor_sameColor_shouldStillWork() throws IOException {
        // Degenerate case: same color for min and max
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf.columns(1)
                .dataBar(ExcelColor.BLUE, ExcelColor.BLUE));

        boolean found = ws.getConditionalFormattingList().stream()
                .flatMap(cf -> cf.getCfRuleList().stream())
                .anyMatch(r -> r.getType() == STCfType.DATA_BAR);
        assertTrue(found);
    }

    // ============================================================
    // Combined rules
    // ============================================================
    @Test
    void dataBar_andCellValueRules_bothApplied() throws IOException {
        CTWorksheet ws = writeAndGetWorksheet(cf -> cf
                .columns(1)
                .greaterThan("50", ExcelColor.LIGHT_RED)
                .dataBar(ExcelColor.BLUE));

        boolean hasDataBar = false;
        boolean hasCellValue = false;
        for (CTConditionalFormatting cf : ws.getConditionalFormattingList()) {
            for (var rule : cf.getCfRuleList()) {
                if (rule.getType() == STCfType.DATA_BAR) hasDataBar = true;
                if (rule.getType() == STCfType.CELL_IS) hasCellValue = true;
            }
        }
        assertTrue(hasDataBar, "Should have data bar rule");
        assertTrue(hasCellValue, "Should have cell-value rule");
    }

    // ============================================================
    // Via ExcelSheetWriter
    // ============================================================
    @Test
    void dataBar_viaExcelSheetWriter_shouldBeApplied() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Data")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .conditionalFormatting(cf -> cf.columns(1).dataBar(ExcelColor.ORANGE))
                    .write(testData());
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var ws = wb.getSheetAt(0).getCTWorksheet();
            boolean found = ws.getConditionalFormattingList().stream()
                    .flatMap(cf -> cf.getCfRuleList().stream())
                    .anyMatch(r -> r.getType() == STCfType.DATA_BAR);
            assertTrue(found);
        }
    }

    @Test
    void iconSet_viaExcelSheetWriter_shouldBeApplied() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Data")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .conditionalFormatting(cf -> cf.columns(1)
                            .iconSet(ExcelConditionalRule.IconSetType.TRAFFIC_LIGHTS_3))
                    .write(testData());
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var ws = wb.getSheetAt(0).getCTWorksheet();
            boolean found = ws.getConditionalFormattingList().stream()
                    .flatMap(cf -> cf.getCfRuleList().stream())
                    .anyMatch(r -> r.getType() == STCfType.ICON_SET);
            assertTrue(found);
        }
    }

    // ============================================================
    // Enum coverage
    // ============================================================
    @Test
    void enumValues_coverage() {
        assertEquals(10, ExcelConditionalRule.IconSetType.values().length);
        for (var t : ExcelConditionalRule.IconSetType.values()) {
            assertEquals(t, ExcelConditionalRule.IconSetType.valueOf(t.name()));
        }
    }
}
