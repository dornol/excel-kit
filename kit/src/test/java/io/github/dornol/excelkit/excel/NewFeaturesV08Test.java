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
            new ExcelWriter<String>()
                    .column("Rotated", s -> s).rotation(45)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(45, style.getRotation(), "Rotation should be 45 degrees");
            }
        }

        @Test
        void rotation_negative45_convertsToPOIFormat() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Rotated", s -> s).rotation(-45)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                // POI: -45 degrees → 90 + 45 = 135
                assertEquals(135, style.getRotation(), "Negative rotation should be converted to POI format");
            }
        }

        @Test
        void rotation_zero_noRotation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("NoRotation", s -> s).rotation(0)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(0, style.getRotation(), "Rotation should be 0");
            }
        }

        @Test
        void rotation_90_vertical() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Vertical", s -> s).rotation(90)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(90, style.getRotation());
            }
        }

        @Test
        void rotation_negative90_vertical() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Vertical", s -> s).rotation(-90)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(180, style.getRotation());
            }
        }

        @Test
        void rotation_negative1_conversionBoundary() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("BoundaryNeg", s -> s).rotation(-1)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                // -1 → 90 + 1 = 91
                assertEquals(91, style.getRotation());
            }
        }

        @Test
        void rotation_outOfRange_throws() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            var builder = writer.column("Col", s -> s);
            assertThrows(IllegalArgumentException.class, () -> builder.rotation(91),
                    "Rotation 91 should throw IllegalArgumentException");
            assertThrows(IllegalArgumentException.class, () -> builder.rotation(-91),
                    "Rotation -91 should throw IllegalArgumentException");
        }

        @Test
        void rotation_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Rotated", s -> s, c -> c.rotation(60))
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                CellStyle style = xwb.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
                assertEquals(30, style.getRotation());
            }
        }

        @Test
        void rotation_columnConfigOutOfRange_throws() {
            ExcelSheetWriter.ColumnConfig<String> config = new ExcelSheetWriter.ColumnConfig<>();
            assertThrows(IllegalArgumentException.class, () -> config.rotation(91));
            assertThrows(IllegalArgumentException.class, () -> config.rotation(-91));
        }

        @Test
        void rotation_withOtherStyling() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("StyledRotated", s -> s)
                        .rotation(45)
                        .bold(true)
                        .backgroundColor(ExcelColor.LIGHT_YELLOW)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Col1", s -> s).rotation(0)
                    .column("Col2", s -> s).rotation(45)
                    .column("Col3", s -> s).rotation(-45)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Borders", s -> s)
                        .borderTop(ExcelBorderStyle.THICK)
                        .borderBottom(ExcelBorderStyle.THIN)
                        .borderLeft(ExcelBorderStyle.DASHED)
                        .borderRight(ExcelBorderStyle.DOTTED)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Borders", s -> s)
                        .border(ExcelBorderStyle.MEDIUM)
                        .borderTop(ExcelBorderStyle.THICK)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Borders", s -> s)
                        .borderTop(ExcelBorderStyle.DOUBLE)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Borders", s -> s)
                        .borderTop(ExcelBorderStyle.NONE)
                        .borderBottom(ExcelBorderStyle.NONE)
                        .borderLeft(ExcelBorderStyle.THICK)
                        .borderRight(ExcelBorderStyle.THICK)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Col", s -> s, c -> c
                            .borderTop(ExcelBorderStyle.MEDIUM)
                            .borderBottom(ExcelBorderStyle.DASHED))
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
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
            new ExcelWriter<String>()
                    .column("A", s -> s)
                        .borderTop(ExcelBorderStyle.THICK)
                    .column("B", s -> s)
                        .border(ExcelBorderStyle.MEDIUM)
                    .column("C", s -> s)
                        .borderLeft(ExcelBorderStyle.DOTTED)
                        .borderRight(ExcelBorderStyle.DOTTED)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 3);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 3, true);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Data", s -> s)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 2);
                        ctx.groupRows(4, 5);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e", "f"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Data", s -> s)
                    .afterData(ctx -> {
                        // Verify chaining works
                        SheetContext result = ctx.groupRows(1, 2).groupRows(3, 4);
                        assertSame(ctx, result, "groupRows should return same context for chaining");
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("a", "b", "c", "d", "e"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
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
            new ExcelWriter<String>()
                    .column("Score", s -> s)
                        .validation(ExcelValidation.integerBetween(1, 100))
                    .write(Stream.of("50"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have at least one validation");
            }
        }

        @Test
        void decimalBetween_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("GPA", s -> s)
                        .validation(ExcelValidation.decimalBetween(0.0, 4.0))
                    .write(Stream.of("3.5"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have decimal validation");
            }
        }

        @Test
        void textLength_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Name", s -> s)
                        .validation(ExcelValidation.textLength(1, 50))
                    .write(Stream.of("John"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have text length validation");
            }
        }

        @Test
        void formula_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Value", s -> s)
                        .validation(ExcelValidation.formula("AND(A2>0,A2<100)"))
                    .write(Stream.of("50"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have formula validation");
            }
        }

        @Test
        void integerGreaterThan_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Qty", s -> s)
                        .validation(ExcelValidation.integerGreaterThan(0))
                    .write(Stream.of("5"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have integer > validation");
            }
        }

        @Test
        void integerLessThan_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Val", s -> s)
                        .validation(ExcelValidation.integerLessThan(1000))
                    .write(Stream.of("500"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have integer < validation");
            }
        }

        @Test
        void dateRange_appliesValidation() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Date", s -> s)
                        .validation(ExcelValidation.dateRange(
                                LocalDate.of(2024, 1, 1),
                                LocalDate.of(2024, 12, 31)))
                    .write(Stream.of("2024-06-15"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have date range validation");
            }
        }

        @Test
        void validation_withErrorMessage() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Age", s -> s)
                        .validation(ExcelValidation.integerBetween(0, 150)
                                .errorTitle("Invalid Age")
                                .errorMessage("Please enter a value between 0 and 150"))
                    .write(Stream.of("25"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have validation with error message");
            }
        }

        @Test
        void validation_withOnlyErrorMessage() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Field", s -> s)
                        .validation(ExcelValidation.integerBetween(1, 10)
                                .errorMessage("Must be 1–10"))
                    .write(Stream.of("5"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0);
            }
        }

        @Test
        void validation_showErrorDisabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Soft", s -> s)
                        .validation(ExcelValidation.integerBetween(1, 100).showError(false))
                    .write(Stream.of("50"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0);
            }
        }

        @Test
        void validation_viaAddColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .addColumn("Score", s -> s, c -> c
                            .validation(ExcelValidation.integerBetween(1, 100)))
                    .write(Stream.of("50"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
            }

            try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = xwb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() > 0, "Should have validation via SheetWriter");
            }
        }

        @Test
        void validation_withDropdown_coexist() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Status", s -> s)
                        .dropdown("Active", "Inactive")
                    .column("Score", s -> s)
                        .validation(ExcelValidation.integerBetween(1, 100))
                    .write(Stream.of("Active"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var validations = wb.getSheetAt(0).getDataValidations();
                assertTrue(validations.size() >= 2, "Should have both dropdown and integer validations");
            }
        }

        @Test
        void validation_multipleColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Age", s -> s)
                        .validation(ExcelValidation.integerBetween(0, 150))
                    .column("GPA", s -> s)
                        .validation(ExcelValidation.decimalBetween(0.0, 4.0))
                    .column("Name", s -> s)
                        .validation(ExcelValidation.textLength(1, 100))
                    .write(Stream.of("25"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Colored", s -> s).fontColor(255, 0, 0)
                    .write(Stream.of("red text"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Blue", s -> s).fontColor(ExcelColor.BLUE)
                    .write(Stream.of("blue text"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Green", s -> s).fontColor(0, 128, 0)
                    .write(Stream.of("green"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Strike", s -> s).strikethrough()
                    .write(Stream.of("struck out"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout(), "Font should be strikethrough");
            }
        }

        @Test
        void strikethrough_explicitTrue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Strike", s -> s).strikethrough(true)
                    .write(Stream.of("struck"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertTrue(font.getStrikeout());
            }
        }

        @Test
        void strikethrough_disabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("NoStrike", s -> s).strikethrough(false)
                    .write(Stream.of("normal text"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertFalse(font.getStrikeout(), "Font should not be strikethrough");
            }
        }

        @Test
        void underline_enabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Underlined", s -> s).underline()
                    .write(Stream.of("underlined text"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_SINGLE, font.getUnderline(), "Font should be underlined");
            }
        }

        @Test
        void underline_explicitTrue() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("UL", s -> s).underline(true)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_SINGLE, font.getUnderline());
            }
        }

        @Test
        void underline_disabled() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("NoUL", s -> s).underline(false)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                Font font = wb.getFontAt(wb.getSheetAt(0).getRow(1).getCell(0).getCellStyle().getFontIndex());
                assertEquals(Font.U_NONE, font.getUnderline());
            }
        }

        @Test
        void combined_allFontStyling() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .column("Styled", s -> s)
                        .fontColor(255, 0, 0)
                        .bold(true)
                        .fontSize(14)
                        .underline()
                        .strikethrough()
                    .write(Stream.of("fully styled"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Styled", s -> s, c -> c
                            .fontColor(ExcelColor.RED)
                            .strikethrough()
                            .underline())
                    .write(Stream.of("styled"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
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
            new ExcelWriter<String>()
                    .column("Both", s -> s)
                        .fontColor(ExcelColor.RED)
                        .backgroundColor(ExcelColor.LIGHT_YELLOW)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .column("Red", s -> s).fontColor(255, 0, 0)
                    .column("Strike", s -> s).strikethrough()
                    .column("Under", s -> s).underline()
                    .column("Plain", s -> s)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .tabColor(255, 0, 0)
                    .addColumn("Data", s -> s)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .tabColor(ExcelColor.BLUE)
                    .addColumn("Data", s -> s)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

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
            new ExcelWriter<String>()
                    .addColumn("Data", s -> s)
                    .write(Stream.of("test"))
                    .consumeOutputStream(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                XSSFColor tabColor = wb.getSheetAt(0).getTabColor();
                assertNull(tabColor, "Tab color should be null when not set");
            }
        }

        @Test
        void tabColor_multipleSheets_applied() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>(2)
                    .tabColor(0, 255, 0)
                    .addColumn("Data", s -> s)
                    .write(Stream.of("a", "b", "c"))
                    .consumeOutputStream(out);

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
                wb.finish().consumeOutputStream(out);
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
                wb.finish().consumeOutputStream(out);
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

            new ExcelWriter<Product>()
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .title("Sales vs Price")
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .showDataLabels(true))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Quarter", Product::name)
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .title("Sales Trend")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Category", Product::name)
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.DOUGHNUT)
                            .title("Sales Distribution")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .showDataLabels(true))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Quarter", Product::name)
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .title("Sales Trend")
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .categoryAxisTitle("Quarter")
                            .valueAxisTitle("Sales Amount"))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .categoryAxisTitle("Sales")
                            .valueAxisTitle("Price"))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .addColumn("Volume", p -> p.sales() * 2, c -> c.type(ExcelDataType.INTEGER))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.SCATTER)
                            .title("Multi-Series Scatter")
                            .categoryColumn(0)
                            .valueColumn(1, "Price")
                            .valueColumn(2, "Volume"))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

            new ExcelWriter<Product>()
                    .addColumn("Quarter", Product::name)
                    .addColumn("Sales", p -> p.sales(), c -> c.type(ExcelDataType.INTEGER))
                    .addColumn("Price", p -> p.price(), c -> c.type(ExcelDataType.DOUBLE))
                    .chart(chart -> chart
                            .type(ExcelChartConfig.ChartType.AREA)
                            .categoryColumn(0)
                            .valueColumn(1, "Sales")
                            .valueColumn(2, "Price"))
                    .write(data.stream())
                    .consumeOutputStream(out);

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

                wb.finish().consumeOutputStream(out);
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
        void allNewFeatures_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<String>()
                    .tabColor(ExcelColor.STEEL_BLUE)
                    .column("Rotated", s -> s)
                        .rotation(45)
                        .fontColor(ExcelColor.RED)
                        .strikethrough()
                        .underline()
                        .bold(true)
                        .borderTop(ExcelBorderStyle.THICK)
                        .borderBottom(ExcelBorderStyle.THIN)
                    .column("Validated", s -> s)
                        .validation(ExcelValidation.integerBetween(1, 100))
                        .backgroundColor(ExcelColor.LIGHT_GREEN)
                    .afterData(ctx -> {
                        ctx.groupRows(1, 2);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("10", "20", "30"))
                    .consumeOutputStream(out);

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
                assertTrue(sheet.getDataValidations().size() > 0);

                // Row grouping
                assertTrue(sheet.getRow(1).getOutlineLevel() > 0);
                assertTrue(sheet.getRow(2).getOutlineLevel() > 0);
            }
        }
    }
}
