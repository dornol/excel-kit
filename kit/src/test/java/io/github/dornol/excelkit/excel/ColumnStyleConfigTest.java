package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ColumnStyleConfig} — the shared base for
 * {@link ExcelColumn.ExcelColumnBuilder} and {@link ExcelSheetWriter.ColumnConfig}.
 * <p>
 * Validates all configuration methods, validation logic, and fluent chaining
 * through the concrete {@link ExcelSheetWriter.ColumnConfig} subclass.
 */
class ColumnStyleConfigTest {

    /**
     * Creates a fresh ColumnConfig for testing.
     */
    private ExcelSheetWriter.ColumnConfig<String> config() {
        return new ExcelSheetWriter.ColumnConfig<>();
    }

    // ============================================================
    // Fluent chaining
    // ============================================================
    @Nested
    class FluentChainingTests {

        @Test
        void allSetters_returnSameInstance() {
            var c = config();
            assertSame(c, c.type(ExcelDataType.STRING));
            assertSame(c, c.format("#,##0"));
            assertSame(c, c.alignment(HorizontalAlignment.LEFT));
            assertSame(c, c.backgroundColor(255, 0, 0));
            assertSame(c, c.backgroundColor(ExcelColor.RED));
            assertSame(c, c.bold(true));
            assertSame(c, c.fontSize(12));
            assertSame(c, c.width(5000));
            assertSame(c, c.minWidth(1000));
            assertSame(c, c.maxWidth(10000));
            assertSame(c, c.dropdown("A", "B"));
            assertSame(c, c.cellColor((v, r) -> null));
            assertSame(c, c.group("G"));
            assertSame(c, c.outline(1));
            assertSame(c, c.comment(r -> null));
            assertSame(c, c.border(ExcelBorderStyle.THIN));
            assertSame(c, c.borderTop(ExcelBorderStyle.THICK));
            assertSame(c, c.borderBottom(ExcelBorderStyle.MEDIUM));
            assertSame(c, c.borderLeft(ExcelBorderStyle.DASHED));
            assertSame(c, c.borderRight(ExcelBorderStyle.DOTTED));
            assertSame(c, c.locked(true));
            assertSame(c, c.hidden());
            assertSame(c, c.hidden(false));
            assertSame(c, c.rotation(45));
            assertSame(c, c.fontColor(0, 0, 255));
            assertSame(c, c.fontColor(ExcelColor.BLUE));
            assertSame(c, c.strikethrough());
            assertSame(c, c.strikethrough(false));
            assertSame(c, c.underline());
            assertSame(c, c.underline(false));
            assertSame(c, c.verticalAlignment(VerticalAlignment.TOP));
            assertSame(c, c.wrapText());
            assertSame(c, c.wrapText(false));
            assertSame(c, c.fontName("Arial"));
            assertSame(c, c.indentation(2));
            assertSame(c, c.validation(ExcelValidation.integerBetween(1, 10)));
        }

        @Test
        void methodChaining_multipleSettings() {
            var c = config()
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .alignment(HorizontalAlignment.RIGHT)
                    .backgroundColor(ExcelColor.LIGHT_YELLOW)
                    .bold(true)
                    .fontSize(14)
                    .rotation(45)
                    .fontColor(255, 0, 0)
                    .strikethrough()
                    .underline()
                    .borderTop(ExcelBorderStyle.THICK)
                    .borderBottom(ExcelBorderStyle.THIN)
                    .validation(ExcelValidation.integerBetween(1, 100));

            // Verify fields are set
            assertEquals(ExcelDataType.INTEGER, c.dataType);
            assertEquals("#,##0", c.dataFormat);
            assertEquals(HorizontalAlignment.RIGHT, c.alignment);
            assertNotNull(c.backgroundColor);
            assertEquals(true, c.bold);
            assertEquals(14, c.fontSize);
            assertNotNull(c.rotation);
            assertNotNull(c.fontColor);
            assertEquals(true, c.strikethrough);
            assertEquals(true, c.underline);
            assertEquals(ExcelBorderStyle.THICK, c.borderTop);
            assertEquals(ExcelBorderStyle.THIN, c.borderBottom);
            assertNotNull(c.validation);
        }
    }

    // ============================================================
    // Validation logic
    // ============================================================
    @Nested
    class ValidationTests {

        @Test
        void fontSize_positive_accepted() {
            var c = config().fontSize(1);
            assertEquals(1, c.fontSize);
            c.fontSize(100);
            assertEquals(100, c.fontSize);
        }

        @Test
        void fontSize_zero_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().fontSize(0));
        }

        @Test
        void fontSize_negative_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().fontSize(-1));
        }

        @Test
        void outline_validRange_accepted() {
            for (int i = 0; i <= 7; i++) {
                var c = config().outline(i);
                assertEquals(i, c.outlineLevel);
            }
        }

        @Test
        void outline_negative_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().outline(-1));
        }

        @Test
        void outline_tooHigh_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().outline(8));
        }

        @Test
        void rotation_validRange_accepted() {
            config().rotation(-90);
            config().rotation(0);
            config().rotation(90);
        }

        @Test
        void rotation_91_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().rotation(91));
        }

        @Test
        void rotation_negative91_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().rotation(-91));
        }
    }

    // ============================================================
    // Rotation conversion
    // ============================================================
    @Nested
    class RotationConversionTests {

        @Test
        void toExcelRotation_positive_unchanged() {
            assertEquals(0, ColumnStyleConfig.toExcelRotation(0));
            assertEquals(45, ColumnStyleConfig.toExcelRotation(45));
            assertEquals(90, ColumnStyleConfig.toExcelRotation(90));
        }

        @Test
        void toExcelRotation_negative_converted() {
            // -1 → 91, -45 → 135, -90 → 180
            assertEquals(91, ColumnStyleConfig.toExcelRotation(-1));
            assertEquals(135, ColumnStyleConfig.toExcelRotation(-45));
            assertEquals(180, ColumnStyleConfig.toExcelRotation(-90));
        }

        @Test
        void rotation_setsConvertedValue() {
            var c = config().rotation(45);
            assertEquals((short) 45, c.rotation);

            c.rotation(-30);
            assertEquals((short) 120, c.rotation);
        }
    }

    // ============================================================
    // Field defaults
    // ============================================================
    @Nested
    class DefaultValueTests {

        @Test
        void defaults_areCorrect() {
            var c = config();
            assertNull(c.dataType);
            assertNull(c.dataFormat);
            assertEquals(HorizontalAlignment.CENTER, c.alignment);
            assertNull(c.backgroundColor);
            assertNull(c.bold);
            assertNull(c.fontSize);
            assertEquals(0, c.minWidth);
            assertEquals(0, c.maxWidth);
            assertFalse(c.fixedWidth);
            assertNull(c.dropdownOptions);
            assertNull(c.cellColorFunction);
            assertNull(c.groupName);
            assertEquals(0, c.outlineLevel);
            assertNull(c.commentFunction);
            assertNull(c.borderStyle);
            assertNull(c.locked);
            assertFalse(c.hidden);
            assertNull(c.validation);
            assertNull(c.rotation);
            assertNull(c.borderTop);
            assertNull(c.borderBottom);
            assertNull(c.borderLeft);
            assertNull(c.borderRight);
            assertNull(c.fontColor);
            assertNull(c.strikethrough);
            assertNull(c.underline);
            assertNull(c.verticalAlignment);
            assertNull(c.wrapText);
            assertNull(c.fontName);
            assertNull(c.indentation);
        }
    }

    // ============================================================
    // Width methods
    // ============================================================
    @Nested
    class WidthTests {

        @Test
        void width_setsFixedWidth() {
            var c = config().width(5000);
            assertTrue(c.fixedWidth);
            assertEquals(5000, c.minWidth);
        }

        @Test
        void minWidth_setsMinOnly() {
            var c = config().minWidth(2000);
            assertFalse(c.fixedWidth);
            assertEquals(2000, c.minWidth);
        }

        @Test
        void maxWidth_setsMaxOnly() {
            var c = config().maxWidth(8000);
            assertEquals(8000, c.maxWidth);
        }
    }

    // ============================================================
    // Hidden methods
    // ============================================================
    @Nested
    class HiddenTests {

        @Test
        void hidden_noArg_setsTrue() {
            var c = config().hidden();
            assertTrue(c.hidden);
        }

        @Test
        void hidden_true() {
            var c = config().hidden(true);
            assertTrue(c.hidden);
        }

        @Test
        void hidden_false() {
            var c = config().hidden(false);
            assertFalse(c.hidden);
        }

        @Test
        void hidden_toggle() {
            var c = config().hidden().hidden(false);
            assertFalse(c.hidden);
        }
    }

    // ============================================================
    // Color methods
    // ============================================================
    @Nested
    class ColorTests {

        @Test
        void backgroundColor_rgb() {
            var c = config().backgroundColor(100, 150, 200);
            assertArrayEquals(new int[]{100, 150, 200}, c.backgroundColor);
        }

        @Test
        void backgroundColor_preset() {
            var c = config().backgroundColor(ExcelColor.RED);
            assertArrayEquals(new int[]{255, 0, 0}, c.backgroundColor);
        }

        @Test
        void fontColor_rgb() {
            var c = config().fontColor(0, 128, 255);
            assertArrayEquals(new int[]{0, 128, 255}, c.fontColor);
        }

        @Test
        void fontColor_preset() {
            var c = config().fontColor(ExcelColor.BLUE);
            assertArrayEquals(new int[]{0, 0, 255}, c.fontColor);
        }
    }

    // ============================================================
    // Strikethrough / Underline
    // ============================================================
    @Nested
    class FontDecorationTests {

        @Test
        void strikethrough_noArg_setsTrue() {
            var c = config().strikethrough();
            assertEquals(true, c.strikethrough);
        }

        @Test
        void strikethrough_boolean() {
            assertEquals(true, config().strikethrough(true).strikethrough);
            assertEquals(false, config().strikethrough(false).strikethrough);
        }

        @Test
        void underline_noArg_setsTrue() {
            var c = config().underline();
            assertEquals(true, c.underline);
        }

        @Test
        void underline_boolean() {
            assertEquals(true, config().underline(true).underline);
            assertEquals(false, config().underline(false).underline);
        }
    }

    // ============================================================
    // Border methods
    // ============================================================
    @Nested
    class BorderTests {

        @Test
        void border_setsUniform() {
            var c = config().border(ExcelBorderStyle.MEDIUM);
            assertEquals(ExcelBorderStyle.MEDIUM, c.borderStyle);
        }

        @Test
        void perSideBorders_setIndependently() {
            var c = config()
                    .borderTop(ExcelBorderStyle.THICK)
                    .borderBottom(ExcelBorderStyle.THIN)
                    .borderLeft(ExcelBorderStyle.DASHED)
                    .borderRight(ExcelBorderStyle.DOTTED);

            assertEquals(ExcelBorderStyle.THICK, c.borderTop);
            assertEquals(ExcelBorderStyle.THIN, c.borderBottom);
            assertEquals(ExcelBorderStyle.DASHED, c.borderLeft);
            assertEquals(ExcelBorderStyle.DOTTED, c.borderRight);
            assertNull(c.borderStyle, "Uniform border should remain null");
        }
    }

    // ============================================================
    // Vertical alignment
    // ============================================================
    @Nested
    class VerticalAlignmentTests {

        @Test
        void verticalAlignment_setsValue() {
            var c = config().verticalAlignment(VerticalAlignment.TOP);
            assertEquals(VerticalAlignment.TOP, c.verticalAlignment);
        }

        @Test
        void verticalAlignment_allValues() {
            for (VerticalAlignment va : VerticalAlignment.values()) {
                var c = config().verticalAlignment(va);
                assertEquals(va, c.verticalAlignment);
            }
        }
    }

    // ============================================================
    // Text wrapping
    // ============================================================
    @Nested
    class WrapTextTests {

        @Test
        void wrapText_noArg_setsTrue() {
            var c = config().wrapText();
            assertEquals(true, c.wrapText);
        }

        @Test
        void wrapText_boolean() {
            assertEquals(true, config().wrapText(true).wrapText);
            assertEquals(false, config().wrapText(false).wrapText);
        }

        @Test
        void wrapText_toggle() {
            var c = config().wrapText().wrapText(false);
            assertEquals(false, c.wrapText);
        }
    }

    // ============================================================
    // Font name
    // ============================================================
    @Nested
    class FontNameTests {

        @Test
        void fontName_setsValue() {
            var c = config().fontName("Arial");
            assertEquals("Arial", c.fontName);
        }

        @Test
        void fontName_korean() {
            var c = config().fontName("맑은 고딕");
            assertEquals("맑은 고딕", c.fontName);
        }

        @Test
        void fontName_overwrite() {
            var c = config().fontName("Arial").fontName("Times New Roman");
            assertEquals("Times New Roman", c.fontName);
        }
    }

    // ============================================================
    // Indentation
    // ============================================================
    @Nested
    class IndentationTests {

        @Test
        void indentation_validRange() {
            var c = config().indentation(0);
            assertEquals((short) 0, c.indentation);

            c.indentation(5);
            assertEquals((short) 5, c.indentation);

            c.indentation(250);
            assertEquals((short) 250, c.indentation);
        }

        @Test
        void indentation_negative_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().indentation(-1));
        }

        @Test
        void indentation_tooHigh_throws() {
            assertThrows(IllegalArgumentException.class, () -> config().indentation(251));
        }
    }

    // ============================================================
    // ExcelColumnBuilder inherits correctly
    // ============================================================
    @Nested
    class InheritanceTests {

        @Test
        void excelColumnBuilder_inheritsAllMethods() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            var builder = writer.column("Col", s -> s)
                    .type(ExcelDataType.INTEGER)
                    .format("#,##0")
                    .alignment(HorizontalAlignment.RIGHT)
                    .backgroundColor(ExcelColor.LIGHT_YELLOW)
                    .bold(true)
                    .fontSize(12)
                    .rotation(30)
                    .fontColor(255, 0, 0)
                    .strikethrough()
                    .underline()
                    .borderTop(ExcelBorderStyle.THICK)
                    .border(ExcelBorderStyle.MEDIUM)
                    .locked(true)
                    .hidden()
                    .validation(ExcelValidation.integerBetween(1, 100))
                    .dropdown("A", "B");

            // Verify it's still an ExcelColumnBuilder (not ColumnConfig)
            assertInstanceOf(ExcelColumn.ExcelColumnBuilder.class, builder);
        }

        @Test
        void columnConfig_inheritsAllMethods() {
            var c = config()
                    .type(ExcelDataType.DOUBLE)
                    .format("#,##0.00")
                    .alignment(HorizontalAlignment.LEFT)
                    .backgroundColor(100, 200, 100)
                    .bold(false)
                    .fontSize(10)
                    .rotation(-45)
                    .fontColor(ExcelColor.GREEN)
                    .strikethrough(true)
                    .underline(true)
                    .borderBottom(ExcelBorderStyle.DOUBLE)
                    .locked(false)
                    .hidden(true)
                    .validation(ExcelValidation.decimalBetween(0, 100));

            assertInstanceOf(ExcelSheetWriter.ColumnConfig.class, c);
        }
    }
}
