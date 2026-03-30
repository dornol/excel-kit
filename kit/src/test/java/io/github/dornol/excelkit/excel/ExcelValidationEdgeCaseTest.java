package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelValidation} — verifies that validation rules
 * are actually written to the Excel file with correct types and parameters.
 */
class ExcelValidationEdgeCaseTest {

    private List<XSSFDataValidation> writeAndGetValidations(ExcelValidation validation) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c.validation(validation))
                .write(Stream.of("test"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return wb.getSheetAt(0).getDataValidations();
        }
    }

    // ============================================================
    // Validation types
    // ============================================================
    @Test
    void decimalBetween_shouldCreateDecimalValidation() throws IOException {
        var validations = writeAndGetValidations(ExcelValidation.decimalBetween(0.0, 999.99));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.DECIMAL, constraint.getValidationType());
        assertEquals("0.0", constraint.getFormula1());
        assertEquals("999.99", constraint.getFormula2());
    }

    @Test
    void textLength_shouldCreateTextLengthValidation() throws IOException {
        var validations = writeAndGetValidations(ExcelValidation.textLength(1, 50));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.TEXT_LENGTH, constraint.getValidationType());
        assertEquals("1", constraint.getFormula1());
        assertEquals("50", constraint.getFormula2());
    }

    @Test
    void dateRange_shouldCreateDateValidation() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.dateRange(LocalDate.of(2020, 1, 1), LocalDate.of(2030, 12, 31)));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.DATE, constraint.getValidationType());
        assertTrue(constraint.getFormula1().contains("2020"));
        assertTrue(constraint.getFormula2().contains("2030"));
    }

    @Test
    void formula_shouldCreateCustomValidation() throws IOException {
        var validations = writeAndGetValidations(ExcelValidation.formula("LEN(A2)>0"));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.FORMULA, constraint.getValidationType());
        assertEquals("LEN(A2)>0", constraint.getFormula1());
    }

    @Test
    void listFromRange_shouldCreateListValidation() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.listFromRange("Sheet2!$A$1:$A$10"));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.LIST, constraint.getValidationType());
        assertEquals("Sheet2!$A$1:$A$10", constraint.getFormula1());
    }

    @Test
    void integerGreaterThan_shouldCreateIntegerValidation() throws IOException {
        var validations = writeAndGetValidations(ExcelValidation.integerGreaterThan(0));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.INTEGER, constraint.getValidationType());
        assertEquals(DataValidationConstraint.OperatorType.GREATER_THAN, constraint.getOperator());
        assertEquals("0", constraint.getFormula1());
    }

    @Test
    void integerLessThan_shouldCreateIntegerValidation() throws IOException {
        var validations = writeAndGetValidations(ExcelValidation.integerLessThan(100));

        assertFalse(validations.isEmpty());
        var constraint = validations.get(0).getValidationConstraint();
        assertEquals(DataValidationConstraint.ValidationType.INTEGER, constraint.getValidationType());
        assertEquals(DataValidationConstraint.OperatorType.LESS_THAN, constraint.getOperator());
        assertEquals("100", constraint.getFormula1());
    }

    // ============================================================
    // Error dialog combinations
    // ============================================================
    @Test
    void errorTitleAndMessage_shouldBeWritten() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.integerBetween(1, 100)
                        .errorTitle("Error Title")
                        .errorMessage("Error Message"));

        assertFalse(validations.isEmpty());
        var v = validations.get(0);
        assertTrue(v.getShowErrorBox());
        assertEquals("Error Title", v.getErrorBoxTitle());
        assertEquals("Error Message", v.getErrorBoxText());
    }

    @Test
    void errorTitle_only_shouldUseEmptyMessage() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.integerBetween(1, 100)
                        .errorTitle("Bad Value"));

        assertFalse(validations.isEmpty());
        assertEquals("Bad Value", validations.get(0).getErrorBoxTitle());
    }

    @Test
    void errorMessage_only_shouldUseDefaultTitle() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.integerBetween(1, 100)
                        .errorMessage("Must be 1-100"));

        assertFalse(validations.isEmpty());
        assertEquals("Validation Error", validations.get(0).getErrorBoxTitle());
        assertEquals("Must be 1-100", validations.get(0).getErrorBoxText());
    }

    @Test
    void noErrorInfo_shouldStillShowErrorBox() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.integerBetween(1, 100));

        assertFalse(validations.isEmpty());
        assertTrue(validations.get(0).getShowErrorBox());
    }

    @Test
    void showError_false_shouldNotShowErrorBox() throws IOException {
        var validations = writeAndGetValidations(
                ExcelValidation.integerBetween(1, 100).showError(false));

        assertFalse(validations.isEmpty());
        assertFalse(validations.get(0).getShowErrorBox());
    }
}
