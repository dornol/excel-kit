package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for data bar and icon set conditional formatting.
 */
class DataBarIconSetTest {

    record Item(String name, int value) {}

    private static Stream<Item> testData() {
        return Stream.of(
                new Item("A", 10), new Item("B", 50),
                new Item("C", 80), new Item("D", 30), new Item("E", 90));
    }

    @Test
    void dataBar_shouldBeApplied() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .dataBar(ExcelColor.BLUE))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            var cfCount = sheet.getCTWorksheet().sizeOfConditionalFormattingArray();
            assertTrue(cfCount > 0, "Should have conditional formatting");
        }
    }

    @Test
    void dataBar_withDifferentColors() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .dataBar(ExcelColor.GREEN))
                .write(testData())
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void iconSet_arrows3_shouldBeApplied() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
                .write(testData())
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cfCount = wb.getSheetAt(0).getCTWorksheet().sizeOfConditionalFormattingArray();
            assertTrue(cfCount > 0, "Should have conditional formatting");
        }
    }

    @ParameterizedTest
    @EnumSource(ExcelConditionalRule.IconSetType.class)
    void iconSet_allTypes_shouldWork(ExcelConditionalRule.IconSetType type) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .iconSet(type))
                .write(testData())
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void dataBar_andRules_combined() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("50", ExcelColor.LIGHT_RED)
                        .dataBar(ExcelColor.BLUE))
                .write(testData())
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void iconSet_viaExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Data")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .conditionalFormatting(cf -> cf
                            .columns(1)
                            .iconSet(ExcelConditionalRule.IconSetType.TRAFFIC_LIGHTS_3))
                    .write(testData());
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void dataBar_viaExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Item>sheet("Data")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .conditionalFormatting(cf -> cf
                            .columns(1)
                            .dataBar(ExcelColor.ORANGE))
                    .write(testData());
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.size() > 0);
    }

    @Test
    void enumValues_coverage() {
        assertEquals(10, ExcelConditionalRule.IconSetType.values().length);
        for (var t : ExcelConditionalRule.IconSetType.values()) {
            assertEquals(t, ExcelConditionalRule.IconSetType.valueOf(t.name()));
        }
    }
}
