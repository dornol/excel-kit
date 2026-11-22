package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class AdvancedValidationTest {

    @Test
    void integerBetween_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Score", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 100)))
                .write(Stream.of("50"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have at least one validation");
        }
    }

    @Test
    void decimalBetween_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("GPA", s -> s, c -> c
                    .validation(ExcelValidation.decimalBetween(0.0, 4.0)))
                .write(Stream.of("3.5"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have decimal validation");
        }
    }

    @Test
    void textLength_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c
                    .validation(ExcelValidation.textLength(1, 50)))
                .write(Stream.of("John"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have text length validation");
        }
    }

    @Test
    void formula_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Value", s -> s, c -> c
                    .validation(ExcelValidation.formula("AND(A2>0,A2<100)")))
                .write(Stream.of("50"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have formula validation");
        }
    }

    @Test
    void integerGreaterThan_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Qty", s -> s, c -> c
                    .validation(ExcelValidation.integerGreaterThan(0)))
                .write(Stream.of("5"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have integer > validation");
        }
    }

    @Test
    void integerLessThan_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Val", s -> s, c -> c
                    .validation(ExcelValidation.integerLessThan(1000)))
                .write(Stream.of("500"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have integer < validation");
        }
    }

    @Test
    void dateRange_appliesValidation() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Date", s -> s, c -> c
                    .validation(ExcelValidation.dateRange(
                            LocalDate.of(2024, 1, 1),
                            LocalDate.of(2024, 12, 31))))
                .write(Stream.of("2024-06-15"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have date range validation");
        }
    }

    @Test
    void validation_withErrorMessage() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Age", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(0, 150)
                            .errorTitle("Invalid Age")
                            .errorMessage("Please enter a value between 0 and 150")))
                .write(Stream.of("25"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have validation with error message");
        }
    }

    @Test
    void validation_withOnlyErrorMessage() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Field", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 10)
                            .errorMessage("Must be 1–10")))
                .write(Stream.of("5"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertEquals(1, validations.size(), "Should have exactly 1 validation rule");
        }
    }

    @Test
    void validation_showErrorDisabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Soft", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 100).showError(false)))
                .write(Stream.of("50"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertEquals(1, validations.size(), "Should have exactly 1 validation rule");
        }
    }

    @Test
    void validation_viaAddColumn() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Score", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)))
                .write(Stream.of("50"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have validation via addColumn");
        }
    }

    @Test
    void validation_viaSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Test")
                    .column("Score", s -> s, c -> c
                            .validation(ExcelValidation.integerBetween(1, 100)))
                    .write(Stream.of("50"));
            wb.finish().writeTo(out);
        }

        try (var xwb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = xwb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() > 0, "Should have validation via SheetWriter");
        }
    }

    @Test
    void validation_withDropdown_coexist() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Status", s -> s, c -> c
                    .dropdown("Active", "Inactive"))
                .column("Score", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 100)))
                .write(Stream.of("Active"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() >= 2, "Should have both dropdown and integer validations");
        }
    }

    @Test
    void validation_multipleColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Age", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(0, 150)))
                .column("GPA", s -> s, c -> c
                    .validation(ExcelValidation.decimalBetween(0.0, 4.0)))
                .column("Name", s -> s, c -> c
                    .validation(ExcelValidation.textLength(1, 100)))
                .write(Stream.of("25"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertTrue(validations.size() >= 3, "Should have 3 different validations");
        }
    }
}
