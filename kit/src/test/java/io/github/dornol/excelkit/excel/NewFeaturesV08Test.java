package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for v0.8 features:
 * 1. Text Rotation
 * 2. Individual Border Control
 * 3. Row Grouping/Outlining
 * 4. Advanced Data Validation
 * 5. Font Color/Strikethrough/Underline
 * 6. Sheet Tab Color
 * 7. Additional Chart Types (SCATTER, AREA, DOUGHNUT)
 */
class NewFeaturesV08Test {

    // ============================================================
    // Feature 1: Text Rotation
    // ============================================================
    @Nested
    class TextRotationTests {

        @Test
        void rotation_positive45_appliesCorrectly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Rotated", s -> s, c -> c.rotation(45))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(45, style.getRotation(), "Rotation should be 45 degrees");
            }
        }

        @Test
        void rotation_negative45_convertsToPOIFormat() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Rotated", s -> s, c -> c.rotation(-45))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                // POI: -45 degrees → 90 + 45 = 135
                assertEquals(135, style.getRotation(), "Negative rotation should be converted to POI format");
            }
        }

        @Test
        void rotation_zero_noRotation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("NoRotation", s -> s, c -> c.rotation(0))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(0, style.getRotation(), "Rotation should be 0");
            }
        }

        @Test
        void rotation_90_vertical() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Vertical", s -> s, c -> c.rotation(90))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(90, style.getRotation());
            }
        }

        @Test
        void rotation_negative90_vertical() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Vertical", s -> s, c -> c.rotation(-90))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(180, style.getRotation());
            }
        }

        @Test
        void rotation_negative1_conversionBoundary() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("BoundaryNeg", s -> s, c -> c.rotation(-1))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                // -1 → 90 + 1 = 91
                assertEquals(91, style.getRotation());
            }
        }

        @Test
        void rotation_outOfRange_throws() {
            ExcelWriter<String> writer = ExcelWriter.<String>builder().build();
            assertThrows(IllegalArgumentException.class, () ->
                    writer.column("Col", s -> s, c -> c.rotation(91)),
                    "Rotation 91 should throw IllegalArgumentException");
            assertThrows(IllegalArgumentException.class, () ->
                    writer.column("Col2", s -> s, c -> c.rotation(-91)),
                    "Rotation -91 should throw IllegalArgumentException");
        }

        @Test
        void rotation_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Rotated", s -> s, c -> c.rotation(60))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(60, style.getRotation());
            }
        }

        @Test
        void rotation_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Test")
                        .column("Rotated", s -> s, c -> c.rotation(30))
                        .write(Stream.of("test"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(30, style.getRotation());
            }
        }

        @Test
        void rotation_columnConfigOutOfRange_throws() {
            ColumnConfig<String> config = new ColumnConfig<>();
            assertThrows(IllegalArgumentException.class, () -> config.rotation(91));
            assertThrows(IllegalArgumentException.class, () -> config.rotation(-91));
        }

        @Test
        void rotation_withOtherStyling() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("StyledRotated", s -> s, c -> c
                        .rotation(45)
                        .bold(true)
                        .backgroundColor(ExcelColor.LIGHT_YELLOW))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(45, style.getRotation(), "Rotation should be preserved with other styles");
                Font font = wb.getFontAt(style.getFontIndex());
                assertTrue(font.getBold(), "Bold should also be applied");
            }
        }

        @Test
        void rotation_multipleColumns_differentAngles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Col1", s -> s, c -> c.rotation(0))
                    .column("Col2", s -> s, c -> c.rotation(45))
                    .column("Col3", s -> s, c -> c.rotation(-45))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertEquals(0, row.getCell(0).getCellStyle().getRotation());
                assertEquals(45, row.getCell(1).getCellStyle().getRotation());
                assertEquals(135, row.getCell(2).getCellStyle().getRotation());
            }
        }
    }

    // ============================================================
    // Feature 2: Individual Border Control
    // ============================================================
    @Nested
    class IndividualBorderTests {

        @Test
        void perSideBorders_allDifferentStyles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Borders", s -> s, c -> c
                        .borderTop(ExcelBorderStyle.THICK)
                        .borderBottom(ExcelBorderStyle.THIN)
                        .borderLeft(ExcelBorderStyle.DASHED)
                        .borderRight(ExcelBorderStyle.DOTTED))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.THICK, style.getBorderTop());
                assertEquals(BorderStyle.THIN, style.getBorderBottom());
                assertEquals(BorderStyle.DASHED, style.getBorderLeft());
                assertEquals(BorderStyle.DOTTED, style.getBorderRight());
            }
        }

        @Test
        void perSideBorders_partialOverride_uniformFallback() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Borders", s -> s, c -> c
                        .border(ExcelBorderStyle.MEDIUM)
                        .borderTop(ExcelBorderStyle.THICK))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.THICK, style.getBorderTop(), "Top should use per-side override");
                assertEquals(BorderStyle.MEDIUM, style.getBorderBottom(), "Bottom should use uniform border");
                assertEquals(BorderStyle.MEDIUM, style.getBorderLeft(), "Left should use uniform border");
                assertEquals(BorderStyle.MEDIUM, style.getBorderRight(), "Right should use uniform border");
            }
        }

        @Test
        void perSideBorders_noUniform_defaultThin() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Borders", s -> s, c -> c
                        .borderTop(ExcelBorderStyle.DOUBLE))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.DOUBLE, style.getBorderTop(), "Top should use per-side");
                assertEquals(BorderStyle.THIN, style.getBorderBottom(), "Bottom should default to THIN");
                assertEquals(BorderStyle.THIN, style.getBorderLeft(), "Left should default to THIN");
                assertEquals(BorderStyle.THIN, style.getBorderRight(), "Right should default to THIN");
            }
        }

        @Test
        void perSideBorders_noneOnSomeSides() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Borders", s -> s, c -> c
                        .borderTop(ExcelBorderStyle.NONE)
                        .borderBottom(ExcelBorderStyle.NONE)
                        .borderLeft(ExcelBorderStyle.THICK)
                        .borderRight(ExcelBorderStyle.THICK))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.NONE, style.getBorderTop());
                assertEquals(BorderStyle.NONE, style.getBorderBottom());
                assertEquals(BorderStyle.THICK, style.getBorderLeft());
                assertEquals(BorderStyle.THICK, style.getBorderRight());
            }
        }

        @Test
        void perSideBorders_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Col", s -> s, c -> c
                            .borderTop(ExcelBorderStyle.MEDIUM)
                            .borderBottom(ExcelBorderStyle.DASHED))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.MEDIUM, style.getBorderTop());
                assertEquals(BorderStyle.DASHED, style.getBorderBottom());
            }
        }

        @Test
        void perSideBorders_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Test")
                        .column("Col", s -> s, c -> c
                                .borderTop(ExcelBorderStyle.DOUBLE)
                                .borderBottom(ExcelBorderStyle.HAIR))
                        .write(Stream.of("test"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(BorderStyle.DOUBLE, style.getBorderTop());
                assertEquals(BorderStyle.HAIR, style.getBorderBottom());
                assertEquals(BorderStyle.THIN, style.getBorderLeft(), "Default THIN for unset sides");
                assertEquals(BorderStyle.THIN, style.getBorderRight(), "Default THIN for unset sides");
            }
        }

        @Test
        void perSideBorders_multipleColumns_mixedStyles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("A", s -> s, c -> c
                        .borderTop(ExcelBorderStyle.THICK))
                    .column("B", s -> s, c -> c
                        .border(ExcelBorderStyle.MEDIUM))
                    .column("C", s -> s, c -> c
                        .borderLeft(ExcelBorderStyle.DOTTED)
                        .borderRight(ExcelBorderStyle.DOTTED))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertEquals(BorderStyle.THICK, row.getCell(0).getCellStyle().getBorderTop());
                assertEquals(BorderStyle.MEDIUM, row.getCell(1).getCellStyle().getBorderTop());
                assertEquals(BorderStyle.DOTTED, row.getCell(2).getCellStyle().getBorderLeft());
                assertEquals(BorderStyle.DOTTED, row.getCell(2).getCellStyle().getBorderRight());
                assertEquals(BorderStyle.THIN, row.getCell(2).getCellStyle().getBorderTop());
            }
        }
    }

    // ============================================================
    // Feature 3: Row Grouping/Outlining
    // ============================================================
    @Nested
    class RowGroupingTests {

        @Test
        void groupRows_basic() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 3);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be grouped");
                assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be grouped");
                assertTrue(sheet.getRow(3).getOutlineLevel() > 0, "Row 3 should be grouped");
                assertEquals(0, sheet.getRow(4).getOutlineLevel(), "Row 4 should not be grouped");
                assertEquals(0, sheet.getRow(5).getOutlineLevel(), "Row 5 should not be grouped");
            }
        }

        @Test
        void groupRows_collapsed() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 3, true);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be grouped");
                assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be grouped");
                assertTrue(sheet.getRow(3).getOutlineLevel() > 0, "Row 3 should be grouped");
            }
        }

        @Test
        void groupRows_multipleGroups() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 2);
                        ctx.groupRows(4, 5);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e", "f"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = wb.getSheetAt(0);
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0, "Row 1 should be in first group");
                assertTrue(sheet.getRow(2).getOutlineLevel() > 0, "Row 2 should be in first group");
                assertEquals(0, sheet.getRow(3).getOutlineLevel(), "Row 3 should not be grouped");
                assertTrue(sheet.getRow(4).getOutlineLevel() > 0, "Row 4 should be in second group");
                assertTrue(sheet.getRow(5).getOutlineLevel() > 0, "Row 5 should be in second group");
            }
        }

        @Test
        void groupRows_chainingReturnsContext() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Data", s -> s)
                    .afterData(ctx -> {
                        // Verify chaining works
                        SheetContext result = ctx.groupRows(1, 2).groupRows(3, 4);
                        assertSame(ctx, result, "groupRows should return same context for chaining");
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertNotNull(wb.getSheetAt(0));
            }
        }

        @Test
        void groupRows_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Test")
                        .column("Data", s -> s)
                        .afterData(ctx -> {
                            ctx.groupRows(1, 3);
                            return ctx.getCurrentRow();
                        })
                        .write(Stream.of("a", "b", "c", "d"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Sheet sheet = xwb.getSheetAt(0);
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0);
                assertTrue(sheet.getRow(3).getOutlineLevel() > 0);
                assertEquals(0, sheet.getRow(4).getOutlineLevel());
            }
        }
    }

    // ============================================================
    // Feature 4: Advanced Data Validation
    // ============================================================
    @Nested
    class AdvancedValidationTests {

        @Test
        void integerBetween_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Score", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)))
                    .write(Stream.of("50"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have at least one validation");
            }
        }

        @Test
        void decimalBetween_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("GPA", s -> s, c -> c
                        .validation(ExcelValidation.decimalBetween(0.0, 4.0)))
                    .write(Stream.of("3.5"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have decimal validation");
            }
        }

        @Test
        void textLength_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Name", s -> s, c -> c
                        .validation(ExcelValidation.textLength(1, 50)))
                    .write(Stream.of("John"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have text length validation");
            }
        }

        @Test
        void formula_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Value", s -> s, c -> c
                        .validation(ExcelValidation.formula("AND(A2>0,A2<100)")))
                    .write(Stream.of("50"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have formula validation");
            }
        }

        @Test
        void integerGreaterThan_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Qty", s -> s, c -> c
                        .validation(ExcelValidation.integerGreaterThan(0)))
                    .write(Stream.of("5"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have integer > validation");
            }
        }

        @Test
        void integerLessThan_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerLessThan(1000)))
                    .write(Stream.of("500"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have integer < validation");
            }
        }

        @Test
        void dateRange_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Date", s -> s, c -> c
                        .validation(ExcelValidation.dateRange(
                                LocalDate.of(2024, 1, 1),
                                LocalDate.of(2024, 12, 31))))
                    .write(Stream.of("2024-06-15"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have date range validation");
            }
        }

        @Test
        void validation_withErrorMessage() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Age", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(0, 150)
                                .errorTitle("Invalid Age")
                                .errorMessage("Please enter a value between 0 and 150")))
                    .write(Stream.of("25"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have validation with error message");
            }
        }

        @Test
        void validation_withOnlyErrorMessage() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Field", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 10)
                                .errorMessage("Must be 1–10")))
                    .write(Stream.of("5"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertEquals(1, validations.size(), "Should have exactly 1 validation rule");
            }
        }

        @Test
        void validation_showErrorDisabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Soft", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100).showError(false)))
                    .write(Stream.of("50"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertEquals(1, validations.size(), "Should have exactly 1 validation rule");
            }
        }

        @Test
        void validation_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Score", s -> s, c -> c
                            .validation(ExcelValidation.integerBetween(1, 100)))
                    .write(Stream.of("50"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have validation via addColumn");
            }
        }

        @Test
        void validation_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Test")
                        .column("Score", s -> s, c -> c
                                .validation(ExcelValidation.integerBetween(1, 100)))
                        .write(Stream.of("50"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = xwb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have validation via SheetWriter");
            }
        }

        @Test
        void validation_withDropdown_coexist() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Status", s -> s, c -> c
                        .dropdown("Active", "Inactive"))
                    .column("Score", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)))
                    .write(Stream.of("Active"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() >= 2, "Should have both dropdown and integer validations");
            }
        }

        @Test
        void validation_multipleColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Age", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(0, 150)))
                    .column("GPA", s -> s, c -> c
                        .validation(ExcelValidation.decimalBetween(0.0, 4.0)))
                    .column("Name", s -> s, c -> c
                        .validation(ExcelValidation.textLength(1, 100)))
                    .write(Stream.of("25"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() >= 3, "Should have 3 different validations");
            }
        }
    }

    // ============================================================
    // Feature 5: Font Color/Strikethrough/Underline
    // ============================================================
    @Nested
    class FontStylingTests {

        @Test
        void fontColor_rgb_appliesCorrectly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Colored", s -> s, c -> c.fontColor(255, 0, 0))
                    .write(Stream.of("red text"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
                XSSFColor color = xssfFont.getXSSFColor();
                assertNotNull(color, "Font color should not be null");
                byte[] rgb = color.getRGB();
                assertNotNull(rgb);
                assertEquals((byte) 255, rgb[0], "Red component");
                assertEquals((byte) 0, rgb[1], "Green component");
                assertEquals((byte) 0, rgb[2], "Blue component");
            }
        }

        @Test
        void fontColor_preset_appliesCorrectly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Blue", s -> s, c -> c.fontColor(ExcelColor.BLUE))
                    .write(Stream.of("blue text"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
                XSSFColor color = xssfFont.getXSSFColor();
                assertNotNull(color);
                byte[] rgb = color.getRGB();
                assertNotNull(rgb);
                assertEquals((byte) 0, rgb[0]);
                assertEquals((byte) 0, rgb[1]);
                assertEquals((byte) 255, rgb[2]);
            }
        }

        @Test
        void fontColor_green_rgb() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Green", s -> s, c -> c.fontColor(0, 128, 0))
                    .write(Stream.of("green"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
                byte[] rgb = xssfFont.getXSSFColor().getRGB();
                assertEquals((byte) 0, rgb[0]);
                assertEquals((byte) 128, rgb[1]);
                assertEquals((byte) 0, rgb[2]);
            }
        }

        @Test
        void strikethrough_enabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Strike", s -> s, c -> c.strikethrough())
                    .write(Stream.of("struck out"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout(), "Font should be strikethrough");
            }
        }

        @Test
        void strikethrough_explicitTrue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Strike", s -> s, c -> c.strikethrough(true))
                    .write(Stream.of("struck"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout());
            }
        }

        @Test
        void strikethrough_disabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("NoStrike", s -> s, c -> c.strikethrough(false))
                    .write(Stream.of("normal text"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertFalse(font.getStrikeout(), "Font should not be strikethrough");
            }
        }

        @Test
        void underline_enabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Underlined", s -> s, c -> c.underline())
                    .write(Stream.of("underlined text"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_SINGLE, font.getUnderline(), "Font should be underlined");
            }
        }

        @Test
        void underline_explicitTrue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("UL", s -> s, c -> c.underline(true))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_SINGLE, font.getUnderline());
            }
        }

        @Test
        void underline_disabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("NoUL", s -> s, c -> c.underline(false))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_NONE, font.getUnderline());
            }
        }

        @Test
        void combined_allFontStyling() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Styled", s -> s, c -> c
                        .fontColor(255, 0, 0)
                        .bold(true)
                        .fontSize(14)
                        .underline()
                        .strikethrough())
                    .write(Stream.of("fully styled"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getBold(), "Font should be bold");
                assertTrue(font.getStrikeout(), "Font should be strikethrough");
                assertEquals(Font.U_SINGLE, font.getUnderline(), "Font should be underlined");
                assertEquals(14, font.getFontHeightInPoints(), "Font size should be 14");
                XSSFFont xssfFont = (XSSFFont) font;
                assertNotNull(xssfFont.getXSSFColor(), "Font color should be set");
            }
        }

        @Test
        void fontStyling_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Styled", s -> s, c -> c
                            .fontColor(ExcelColor.RED)
                            .strikethrough()
                            .underline())
                    .write(Stream.of("styled"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout());
                assertEquals(Font.U_SINGLE, font.getUnderline());
            }
        }

        @Test
        void fontStyling_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Test")
                        .column("Colored", s -> s, c -> c
                                .fontColor(0, 128, 0)
                                .underline()
                                .strikethrough())
                        .write(Stream.of("styled"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = xwb.getFontAt(xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout());
                assertEquals(Font.U_SINGLE, font.getUnderline());
                XSSFFont xf = (XSSFFont) font;
                assertNotNull(xf.getXSSFColor());
            }
        }

        @Test
        void fontColor_withBackgroundColor_bothApply() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Both", s -> s, c -> c
                        .fontColor(ExcelColor.RED)
                        .backgroundColor(ExcelColor.LIGHT_YELLOW))
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                XSSFFont xssfFont = (XSSFFont) wb.getFontAt(style.getFontIndex());
                assertNotNull(xssfFont.getXSSFColor(), "Font color should be set");
                assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
            }
        }

        @Test
        void multipleColumns_differentFontStyles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Red", s -> s, c -> c.fontColor(255, 0, 0))
                    .column("Strike", s -> s, c -> c.strikethrough())
                    .column("Under", s -> s, c -> c.underline())
                    .column("Plain", s -> s)
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                // Red column has color
                XSSFFont f0 = (XSSFFont) wb.getFontAt(row.getCell(0).getCellStyle().getFontIndex());
                assertNotNull(f0.getXSSFColor());
                // Strike column
                Font f1 = wb.getFontAt(row.getCell(1).getCellStyle().getFontIndex());
                assertTrue(f1.getStrikeout());
                // Underline column
                Font f2 = wb.getFontAt(row.getCell(2).getCellStyle().getFontIndex());
                assertEquals(Font.U_SINGLE, f2.getUnderline());
            }
        }
    }

    // ============================================================
    // Feature 6: Sheet Tab Color
    // ============================================================
    @Nested
    class TabColorTests {

        @Test
        void tabColor_rgb_appliesCorrectly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .tabColor(255, 0, 0)
                    .column("Data", s -> s)
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
                assertNotNull(tabColor, "Tab color should be set");
                byte[] rgb = tabColor.getRGB();
                assertNotNull(rgb);
                assertEquals((byte) 255, rgb[0], "Red component");
                assertEquals((byte) 0, rgb[1], "Green component");
                assertEquals((byte) 0, rgb[2], "Blue component");
            }
        }

        @Test
        void tabColor_preset_appliesCorrectly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .tabColor(ExcelColor.BLUE)
                    .column("Data", s -> s)
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
                assertNotNull(tabColor, "Tab color should be set");
                byte[] rgb = tabColor.getRGB();
                assertNotNull(rgb);
                assertEquals((byte) 0, rgb[0]);
                assertEquals((byte) 0, rgb[1]);
                assertEquals((byte) 255, rgb[2]);
            }
        }

        @Test
        void tabColor_noColor_tabColorIsNull() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .column("Data", s -> s)
                    .write(Stream.of("test"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
                assertNull(tabColor, "Tab color should be null when not set");
            }
        }

        @Test
        void tabColor_multipleSheets_applied() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().maxRows(2).build()
                    .tabColor(0, 255, 0)
                    .column("Data", s -> s)
                    .write(Stream.of("a", "b", "c"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                // Tab color should be applied to all rollover sheets
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    XSSFColor tabColor = wb.getSheetAt(i).getTabColor();
                    assertNotNull(tabColor, "Tab color should be set on sheet " + i);
                }
            }
        }

        @Test
        void tabColor_viaSheetWriter() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("TabTest")
                        .tabColor(ExcelColor.RED)
                        .column("Data", s -> s)
                        .write(Stream.of("test"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFColor tabColor = xwb.getSheetAt(0).getTabColor();
                assertNotNull(tabColor, "Tab color should be set via SheetWriter");
                byte[] rgb = tabColor.getRGB();
                assertEquals((byte) 255, rgb[0]);
                assertEquals((byte) 0, rgb[1]);
                assertEquals((byte) 0, rgb[2]);
            }
        }

        @Test
        void tabColor_differentPerSheet() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Red")
                        .tabColor(ExcelColor.RED)
                        .column("Data", s -> s)
                        .write(Stream.of("a"));
                wb.<String>sheet("Blue")
                        .tabColor(ExcelColor.BLUE)
                        .column("Data", s -> s)
                        .write(Stream.of("b"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                byte[] red = xwb.getSheetAt(0).getTabColor().getRGB();
                byte[] blue = xwb.getSheetAt(1).getTabColor().getRGB();
                assertEquals((byte) 255, red[0]);
                assertEquals((byte) 0, blue[0]);
                assertEquals((byte) 255, blue[2]);
            }
        }
    }

    // ============================================================
    // Feature 7: Additional Chart Types
    // ============================================================
    @Nested
    class AdditionalChartTests {

        record Product(String name, int sales, double price) {}

        @Test
        void scatterChart_createsSuccessfully() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("A", 100, 10.5),
                    new Product("B", 200, 20.3),
                    new Product("C", 150, 15.7)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .title("Sales vs Price")
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .showDataLabels(true))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
                assertNotNull(drawing, "Drawing should exist for scatter chart");
                assertFalse(drawing.getCharts().isEmpty(), "Should have at least one chart");
                // Verify it is a scatter chart
                var scatterList = drawing.getCharts().get(0).getCTChart()
                        .getPlotArea().getScatterChartList();
                assertFalse(scatterList.isEmpty(), "Should have a scatter chart");
            }
        }

        @Test
        void areaChart_createsSuccessfully() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("Q1", 100, 10.5),
                    new Product("Q2", 200, 20.3),
                    new Product("Q3", 150, 15.7)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Quarter", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .title("Sales Trend")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
                assertNotNull(drawing, "Drawing should exist for area chart");
                assertFalse(drawing.getCharts().isEmpty());
                var areaList = drawing.getCharts().get(0).getCTChart()
                        .getPlotArea().getAreaChartList();
                assertFalse(areaList.isEmpty(), "Should have an area chart");
            }
        }

        @Test
        void doughnutChart_createsSuccessfully() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("Phone", 300, 0),
                    new Product("Laptop", 200, 0),
                    new Product("Tablet", 100, 0)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Category", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.DOUGHNUT)
                            .title("Sales Distribution")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .showDataLabels(true))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFDrawing drawing = wb.getSheetAt(0).getDrawingPatriarch();
                assertNotNull(drawing, "Drawing should exist for doughnut chart");
                assertFalse(drawing.getCharts().isEmpty());
                var doughnutList = drawing.getCharts().get(0).getCTChart()
                        .getPlotArea().getDoughnutChartList();
                assertFalse(doughnutList.isEmpty(), "Should have a doughnut chart");
            }
        }

        @Test
        void areaChart_withAxisTitles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("Q1", 100, 10.5),
                    new Product("Q2", 200, 20.3)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Quarter", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .title("Sales Trend")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .categoryAxisTitle("Quarter")
                            .valueAxisTitle("Sales Amount"))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
                var catAxis = chart.getCTChart().getPlotArea().getCatAxList();
                assertFalse(catAxis.isEmpty(), "Chart should have a category axis");
                assertTrue(catAxis.get(0).isSetTitle(), "Category axis should have a title");
            }
        }

        @Test
        void scatterChart_withAxisTitles() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("A", 100, 10.5),
                    new Product("B", 200, 20.3)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .categoryAxisTitle("Sales")
                            .valueAxisTitle("Price"))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
                // Scatter uses value axes (not category axes)
                var valAxes = chart.getCTChart().getPlotArea().getValAxList();
                assertTrue(valAxes.size() >= 2, "Scatter chart should have 2 value axes");
            }
        }

        @Test
        void scatterChart_multiSeries() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("A", 100, 10.5),
                    new Product("B", 200, 20.3)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .column("Volume", p -> p.sales() * 2, c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .title("Multi-Series Scatter")
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .valueColumn(2, "Volume"))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFChart chart = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0);
                var scatterList = chart.getCTChart().getPlotArea().getScatterChartList();
                assertFalse(scatterList.isEmpty());
                // Should have 2 series
                assertEquals(2, scatterList.get(0).getSerList().size(), "Should have 2 scatter series");
            }
        }

        @Test
        void areaChart_multipleSeries() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            List<Product> data = List.of(
                    new Product("Q1", 100, 10.5),
                    new Product("Q2", 200, 20.3)
            );

            ExcelWriter.<Product>builder().build()
                    .column("Quarter", Product::name)
                    .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .valueColumn(2, "Price"))
                    .write(data.stream())
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var areaList = wb.getSheetAt(0).getDrawingPatriarch().getCharts()
                        .get(0).getCTChart().getPlotArea().getAreaChartList();
                assertEquals(2, areaList.get(0).getSerList().size(), "Should have 2 area series");
            }
        }

        @Test
        void newChartTypes_inExcelWorkbook() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<Product>sheet("Scatter")
                        .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                        .column("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                        .chart(chart -> chart
                                .type(ExcelChartConfig.ChartType.SCATTER)
                                .categoryColumn(0)
                                .valueColumn(1, "Price"))
                        .write(Stream.of(new Product("A", 100, 10.5)));

                wb.<Product>sheet("Area")
                        .column("Name", Product::name)
                        .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                        .chart(chart -> chart
                                .type(ExcelChartConfig.ChartType.AREA)
                                .categoryColumn(0)
                                .valueColumn(1, "Sales"))
                        .write(Stream.of(new Product("Q1", 100, 0)));

                wb.<Product>sheet("Doughnut")
                        .column("Name", Product::name)
                        .column("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                        .chart(chart -> chart
                                .type(ExcelChartConfig.ChartType.DOUGHNUT)
                                .categoryColumn(0)
                                .valueColumn(1, "Sales"))
                        .write(Stream.of(new Product("A", 60, 0)));

                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(3, xwb.getNumberOfSheets());
                for (int i = 0; i < 3; i++) {
                    assertNotNull(xwb.getSheetAt(i).getDrawingPatriarch(),
                            "Sheet " + i + " should have a chart");
                }
            }
        }

        @Test
        void chartType_enum_coverage() {
            assertEquals(6, ExcelChartConfig.ChartType.values().length);
            assertNotNull(ExcelChartConfig.ChartType.valueOf("BAR"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("LINE"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("PIE"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("SCATTER"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("AREA"));
            assertNotNull(ExcelChartConfig.ChartType.valueOf("DOUGHNUT"));
        }
    }

    // ============================================================
    // Feature combination tests
    // ============================================================
    @Nested
    class CombinationTests {

        @Test
        void allNewFeatures_viaSheetWriter_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
                wb.<String>sheet("Styled")
                        .tabColor(ExcelColor.RED)
                        .column("Rotated", s -> s, c -> c
                                .rotation(45)
                                .fontColor(ExcelColor.RED)
                                .strikethrough()
                                .underline()
                                .bold(true)
                                .fontSize(14)
                                .borderTop(ExcelBorderStyle.THICK)
                                .borderBottom(ExcelBorderStyle.THIN)
                                .borderLeft(ExcelBorderStyle.DASHED)
                                .borderRight(ExcelBorderStyle.DOTTED)
                                .backgroundColor(ExcelColor.LIGHT_YELLOW))
                        .column("Validated", s -> s, c -> c
                                .validation(ExcelValidation.integerBetween(1, 100))
                                .type(ExcelDataType.INTEGER))
                        .column("Hidden", s -> s, c -> c.hidden())
                        .column("Dropdown", s -> s, c -> c.dropdown("A", "B", "C"))
                        .afterData(ctx -> {
                            ctx.groupRows(1, 2);
                            return ctx.getCurrentRow();
                        })
                        .write(Stream.of("10", "20", "30"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = xwb.getSheetAt(0);

                // Tab color
                assertNotNull(sheet.getTabColor());
                assertEquals((byte) 255, sheet.getTabColor().getRGB()[0]);

                // Rotation + font styling + borders on first column
                CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
                assertEquals(45, style.getRotation());
                assertEquals(BorderStyle.THICK, style.getBorderTop());
                assertEquals(BorderStyle.THIN, style.getBorderBottom());
                assertEquals(BorderStyle.DASHED, style.getBorderLeft());
                assertEquals(BorderStyle.DOTTED, style.getBorderRight());
                Font font = xwb.getFontAt(style.getFontIndex());
                assertTrue(font.getBold());
                assertTrue(font.getStrikeout());
                assertEquals(Font.U_SINGLE, font.getUnderline());
                assertEquals(14, font.getFontHeightInPoints());

                // Validation + dropdown
                assertTrue(sheet.getDataValidations().size() >= 2);

                // Hidden column
                assertTrue(sheet.isColumnHidden(2));

                // Row grouping
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0);
                assertTrue(sheet.getRow(2).getOutlineLevel() > 0);
            }
        }

        @Test
        void sheetWriter_noConfig_defaultsWork() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<String>sheet("Default")
                        .column("Plain", s -> s)
                        .write(Stream.of("a", "b"));
                wb.finish().write(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = xwb.getSheetAt(0);
                assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("b", sheet.getRow(2).getCell(0).getStringCellValue());
            }
        }

        @Test
        void allNewFeatures_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>builder().build()
                    .tabColor(ExcelColor.STEEL_BLUE)
                    .column("Rotated", s -> s, c -> c
                        .rotation(45)
                        .fontColor(ExcelColor.RED)
                        .strikethrough()
                        .underline()
                        .bold(true)
                        .borderTop(ExcelBorderStyle.THICK)
                        .borderBottom(ExcelBorderStyle.THIN))
                    .column("Validated", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100))
                        .backgroundColor(ExcelColor.LIGHT_GREEN))
                    .afterData(ctx -> {
                        ctx.groupRows(1, 2);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("10", "20", "30"))
                    .write(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);

                // Tab color
                assertNotNull(sheet.getTabColor(), "Tab color should be set");

                // Rotation + font styling on first column
                CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
                assertEquals(45, style.getRotation());
                Font font = wb.getFontAt(style.getFontIndex());
                assertTrue(font.getBold());
                assertTrue(font.getStrikeout());
                assertEquals(Font.U_SINGLE, font.getUnderline());
                assertEquals(BorderStyle.THICK, style.getBorderTop());
                assertEquals(BorderStyle.THIN, style.getBorderBottom());

                // Validation on second column
                assertFalse(sheet.getDataValidations().isEmpty(), "Should have validation rules");

                // Row grouping
                assertEquals(1, sheet.getRow(1).getOutlineLevel(), "Grouped row should have outline level 1");
                assertEquals(1, sheet.getRow(2).getOutlineLevel(), "Grouped row should have outline level 1");
            }
        }
    }
}
