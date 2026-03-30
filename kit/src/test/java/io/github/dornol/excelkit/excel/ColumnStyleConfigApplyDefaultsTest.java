package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ColumnStyleConfig#applyDefaults(ColumnStyleConfig)} to cover
 * all null-check branches where defaults are applied.
 */
class ColumnStyleConfigApplyDefaultsTest {

    private ExcelSheetWriter.ColumnConfig<String> config() {
        return new ExcelSheetWriter.ColumnConfig<>();
    }

    private ColumnStyleConfig.DefaultStyleConfig<String> defaults() {
        return new ColumnStyleConfig.DefaultStyleConfig<>();
    }

    @Test
    void applyDefaults_allFieldsNull_allDefaultsApplied() {
        var target = config();
        var def = defaults()
                .type(ExcelDataType.INTEGER)
                .format("#,##0")
                .backgroundColor(255, 0, 0)
                .bold(true)
                .fontSize(14)
                .border(ExcelBorderStyle.THICK)
                .locked(true)
                .rotation(45)
                .borderTop(ExcelBorderStyle.THIN)
                .borderBottom(ExcelBorderStyle.MEDIUM)
                .borderLeft(ExcelBorderStyle.DASHED)
                .borderRight(ExcelBorderStyle.DOTTED)
                .fontColor(0, 0, 255)
                .strikethrough(true)
                .underline(true)
                .verticalAlignment(VerticalAlignment.TOP)
                .wrapText(true)
                .fontName("Arial")
                .indentation(5)
                .alignment(HorizontalAlignment.RIGHT);

        target.applyDefaults(def);

        assertEquals(ExcelDataType.INTEGER, target.dataType);
        assertEquals("#,##0", target.dataFormat);
        assertArrayEquals(new int[]{255, 0, 0}, target.backgroundColor);
        assertEquals(true, target.bold);
        assertEquals(14, target.fontSize);
        assertEquals(ExcelBorderStyle.THICK, target.borderStyle);
        assertEquals(true, target.locked);
        assertEquals(ColumnStyleConfig.toExcelRotation(45), target.rotation, "Rotation should be 45");
        assertEquals(ExcelBorderStyle.THIN, target.borderTop);
        assertEquals(ExcelBorderStyle.MEDIUM, target.borderBottom);
        assertEquals(ExcelBorderStyle.DASHED, target.borderLeft);
        assertEquals(ExcelBorderStyle.DOTTED, target.borderRight);
        assertArrayEquals(new int[]{0, 0, 255}, target.fontColor);
        assertEquals(true, target.strikethrough);
        assertEquals(true, target.underline);
        assertEquals(VerticalAlignment.TOP, target.verticalAlignment);
        assertEquals(true, target.wrapText);
        assertEquals("Arial", target.fontName);
        assertEquals((short) 5, target.indentation);
        assertEquals(HorizontalAlignment.RIGHT, target.alignment);
        assertTrue(target.alignmentSet);
    }

    @Test
    void applyDefaults_existingFieldsNotOverridden() {
        var target = config()
                .type(ExcelDataType.STRING)
                .format("@")
                .backgroundColor(0, 255, 0)
                .bold(false)
                .fontSize(10)
                .border(ExcelBorderStyle.THIN)
                .locked(false)
                .rotation(0)
                .borderTop(ExcelBorderStyle.DOUBLE)
                .borderBottom(ExcelBorderStyle.DOUBLE)
                .borderLeft(ExcelBorderStyle.DOUBLE)
                .borderRight(ExcelBorderStyle.DOUBLE)
                .fontColor(255, 255, 0)
                .strikethrough(false)
                .underline(false)
                .verticalAlignment(VerticalAlignment.BOTTOM)
                .wrapText(false)
                .fontName("Times New Roman")
                .indentation(2)
                .alignment(HorizontalAlignment.LEFT);

        var def = defaults()
                .type(ExcelDataType.INTEGER)
                .format("#,##0")
                .backgroundColor(255, 0, 0)
                .bold(true)
                .fontSize(14)
                .border(ExcelBorderStyle.THICK)
                .locked(true)
                .rotation(90)
                .fontColor(0, 0, 255)
                .strikethrough(true)
                .underline(true)
                .verticalAlignment(VerticalAlignment.TOP)
                .wrapText(true)
                .fontName("Arial")
                .indentation(10)
                .alignment(HorizontalAlignment.RIGHT);

        target.applyDefaults(def);

        // All should retain original values
        assertEquals(ExcelDataType.STRING, target.dataType);
        assertEquals("@", target.dataFormat);
        assertArrayEquals(new int[]{0, 255, 0}, target.backgroundColor);
        assertEquals(false, target.bold);
        assertEquals(10, target.fontSize);
        assertEquals(ExcelBorderStyle.THIN, target.borderStyle);
        assertEquals(false, target.locked);
        assertEquals(VerticalAlignment.BOTTOM, target.verticalAlignment);
        assertEquals(false, target.wrapText);
        assertEquals("Times New Roman", target.fontName);
        assertEquals((short) 2, target.indentation);
        assertEquals(HorizontalAlignment.LEFT, target.alignment);
    }

    @Test
    void applyDefaults_alignmentSet_onTarget_notOverridden() {
        var target = config().alignment(HorizontalAlignment.LEFT);
        var def = defaults().alignment(HorizontalAlignment.RIGHT);

        target.applyDefaults(def);

        assertEquals(HorizontalAlignment.LEFT, target.alignment);
        assertTrue(target.alignmentSet);
    }

    @Test
    void applyDefaults_alignmentNotSet_onEither_usesDefault() {
        var target = config();
        var def = defaults();

        target.applyDefaults(def);

        assertEquals(HorizontalAlignment.CENTER, target.alignment);
        assertFalse(target.alignmentSet);
    }

    @Test
    void applyDefaults_emptyDefaults_noChange() {
        var target = config().type(ExcelDataType.DOUBLE);
        var def = defaults();

        target.applyDefaults(def);

        assertEquals(ExcelDataType.DOUBLE, target.dataType);
        assertNull(target.dataFormat);
        assertNull(target.bold);
    }
}
