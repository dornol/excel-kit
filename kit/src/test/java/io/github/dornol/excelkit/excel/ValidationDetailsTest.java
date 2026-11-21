package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The data validations a column writes, and what is refused when it is set up instead of when Excel opens the file. */
class ValidationDetailsTest {

    record Row(String name, int value) {
    }

    private static List<CTDataValidation> validations(Consumer<ColumnStyleConfig<Row, ?>> column) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().column("name", Row::name, column::accept).column("value", Row::value)
                .write(List.of(new Row("a", 1), new Row("b", 2))).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var all = workbook.getSheetAt(0).getCTWorksheet().getDataValidations();
            return all == null ? List.of() : new ArrayList<>(List.of(all.getDataValidationArray()));
        }
    }

    // ---- what is written

    @Test
    void aValidationCoversTheColumnBelowTheHeader() throws Exception {
        List<CTDataValidation> written = validations(c -> c.validation(ExcelValidation.integerBetween(1, 10)));
        assertEquals(1, written.size());
        CTDataValidation validation = written.get(0);
        assertEquals("whole", validation.getType().toString());
        assertEquals("between", validation.getOperator().toString());
        assertEquals("1", validation.getFormula1());
        assertEquals("10", validation.getFormula2());
        assertEquals("A2:A1048576", validation.getSqref().get(0).toString());
        assertTrue(validation.getShowErrorMessage());
    }

    @Test
    void everyKindWritesItsTypeOperatorAndBounds() throws Exception {
        record Case(ExcelValidation validation, String type, String operator, String first, String second) {
        }
        for (Case test : List.of(
                new Case(ExcelValidation.integerGreaterThan(3), "whole", "greaterThan", "3", null),
                new Case(ExcelValidation.integerLessThan(9), "whole", "lessThan", "9", null),
                new Case(ExcelValidation.decimalBetween(0.5, 2.5), "decimal", "between", "0.5", "2.5"),
                new Case(ExcelValidation.decimalBetween(1e-7, 1e10), "decimal", "between", "1.0E-7", "1.0E10"),
                new Case(ExcelValidation.textLength(1, 20), "textLength", "between", "1", "20"),
                new Case(ExcelValidation.dateRange(LocalDate.of(2024, 1, 31), LocalDate.of(2024, 12, 31)), "date", "between",
                        "DATE(2024,1,31)", "DATE(2024,12,31)"),
                new Case(ExcelValidation.listFromRange("$A$1:$A$5"), "list", "between", "$A$1:$A$5", null),
                new Case(ExcelValidation.listFromRange("=Sheet2!$A$1:$A$5"), "list", "between", "Sheet2!$A$1:$A$5", null),
                new Case(ExcelValidation.formula("LEN(A2)>3"), "custom", "between", "LEN(A2)>3", null),
                new Case(ExcelValidation.formula("=LEN(A2)>3"), "custom", "between", "LEN(A2)>3", null))) {
            CTDataValidation written = validations(c -> c.validation(test.validation())).get(0);
            assertEquals(test.type(), written.getType().toString(), test.first());
            assertEquals(test.operator(), written.getOperator().toString(), test.first());
            assertEquals(test.first(), written.getFormula1());
            assertEquals(test.second(), written.getFormula2());
        }
    }

    @Test
    void theErrorBoxCarriesTheTitleAndMessageGiven() throws Exception {
        CTDataValidation both = validations(c -> c.validation(ExcelValidation.integerBetween(1, 2)
                .errorTitle("Too big").errorMessage("Use 1 or 2"))).get(0);
        assertEquals("Too big", both.getErrorTitle());
        assertEquals("Use 1 or 2", both.getError());

        CTDataValidation messageOnly = validations(c -> c.validation(ExcelValidation.integerBetween(1, 2).errorMessage("bad"))).get(0);
        assertEquals("Validation Error", messageOnly.getErrorTitle());
        assertEquals("bad", messageOnly.getError());

        CTDataValidation titleOnly = validations(c -> c.validation(ExcelValidation.integerBetween(1, 2).errorTitle("Oops"))).get(0);
        assertEquals("Oops", titleOnly.getErrorTitle());
        assertEquals("", titleOnly.getError());

        CTDataValidation silent = validations(c -> c.validation(ExcelValidation.integerBetween(1, 2).showError(false))).get(0);
        assertEquals(false, silent.getShowErrorMessage());
        assertNull(silent.getErrorTitle());
    }

    @Test
    void aDropdownIsAQuotedCommaSeparatedList() throws Exception {
        CTDataValidation list = validations(c -> c.dropdown("a", "b", "c")).get(0);
        assertEquals("list", list.getType().toString());
        assertEquals("\"a,b,c\"", list.getFormula1());
        assertEquals("A2:A1048576", list.getSqref().get(0).toString());
    }

    @Test
    void aQuoteInADropdownOptionIsWrittenTwiceAsExcelWantsIt() throws Exception {
        CTDataValidation list = validations(c -> c.dropdown("say \"hi\"", "c")).get(0);
        assertEquals("\"say \"\"hi\"\",c\"", list.getFormula1());
    }

    // ---- what is refused at once

    @Test
    void anOptionWithACommaIsRefusedBecauseExcelWouldSplitIt() {
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class, () -> new ColumnConfig<Row>().dropdown("a,b", "c"));
        assertTrue(error.getMessage().contains("comma") && error.getMessage().contains("listFromRange"), error.getMessage());
    }

    @Test
    void aDropdownNeedsOptionsAndNoNullOnes() {
        assertThrows(IllegalArgumentException.class, () -> new ColumnConfig<Row>().dropdown());
        assertThrows(NullPointerException.class, () -> new ColumnConfig<Row>().dropdown("a", null));
        assertDoesNotThrow(() -> new ColumnConfig<Row>().dropdown("only"));
    }

    @Test
    void theDropdownLimitCountsTheQuotesThatAreWrittenTwice() {
        // 127 + 1 + 127 = 255 characters fit; one more does not; a quote counts for two
        assertDoesNotThrow(() -> new ColumnConfig<Row>().dropdown("x".repeat(127), "y".repeat(127)));
        assertThrows(IllegalArgumentException.class, () -> new ColumnConfig<Row>().dropdown("x".repeat(128), "y".repeat(127)));
        assertThrows(IllegalArgumentException.class, () -> new ColumnConfig<Row>().dropdown("\"".repeat(128), "y".repeat(127)));
    }

    @Test
    void boundsThatCouldNeverBeMetAreRefused() {
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.integerBetween(10, 1));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.decimalBetween(2.0, 1.0));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.textLength(5, 1));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.textLength(-1, 5));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.dateRange(LocalDate.of(2025, 1, 1), LocalDate.of(2024, 1, 1)));
        // equal bounds are a single allowed value
        assertDoesNotThrow(() -> ExcelValidation.integerBetween(5, 5));
        assertDoesNotThrow(() -> ExcelValidation.decimalBetween(1.5, 1.5));
        assertDoesNotThrow(() -> ExcelValidation.textLength(0, 0));
        assertDoesNotThrow(() -> ExcelValidation.dateRange(LocalDate.of(2024, 1, 1), LocalDate.of(2024, 1, 1)));
    }

    @Test
    void decimalBoundsMustBeNumbers() {
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.decimalBetween(Double.NaN, 1));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.decimalBetween(0, Double.POSITIVE_INFINITY));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.decimalBetween(Double.NEGATIVE_INFINITY, 1));
    }

    @Test
    void aListRangeAndAFormulaNeedText() {
        assertThrows(NullPointerException.class, () -> ExcelValidation.listFromRange(null));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.listFromRange(""));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.listFromRange("  "));
        assertThrows(NullPointerException.class, () -> ExcelValidation.formula(null));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.formula(""));
        assertThrows(NullPointerException.class, () -> ExcelValidation.dateRange(null, LocalDate.now()));
        assertThrows(NullPointerException.class, () -> ExcelValidation.dateRange(LocalDate.now(), null));
    }

    @Test
    void theErrorBoxHasExcelsLengthLimits() {
        assertDoesNotThrow(() -> ExcelValidation.integerBetween(1, 2).errorTitle("t".repeat(32)));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.integerBetween(1, 2).errorTitle("t".repeat(33)));
        assertDoesNotThrow(() -> ExcelValidation.integerBetween(1, 2).errorMessage("m".repeat(255)));
        assertThrows(IllegalArgumentException.class, () -> ExcelValidation.integerBetween(1, 2).errorMessage("m".repeat(256)));
        assertDoesNotThrow(() -> ExcelValidation.integerBetween(1, 2).errorTitle(null).errorMessage(null));
    }

    // ---- a column has one validation

    @Test
    void aDropdownAndAValidationOnOneColumnAreRefused() {
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class, () ->
                ExcelWriter.<Row>create().column("name", Row::name, c -> c.dropdown("a").validation(ExcelValidation.integerBetween(1, 2))));
        assertTrue(error.getMessage().contains("both a dropdown and a validation"), error.getMessage());
        assertThrows(IllegalArgumentException.class, () ->
                ExcelWriter.<Row>create().column("name", Row::name, c -> c.validation(ExcelValidation.integerBetween(1, 2)).dropdown("a")));
    }

    @Test
    void aDefaultOfTheOtherKindDoesNotJoinAColumnsOwnChoice() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().defaultStyle(d -> d.dropdown("x", "y"))
                .column("own", Row::name, c -> c.validation(ExcelValidation.integerBetween(1, 2)))
                .column("default", Row::name)
                .write(List.of(new Row("a", 1))).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var written = workbook.getSheetAt(0).getCTWorksheet().getDataValidations().getDataValidationArray();
            assertEquals(2, written.length);
            assertEquals("whole", written[0].getType().toString());
            assertEquals("A2:A1048576", written[0].getSqref().get(0).toString());
            assertEquals("list", written[1].getType().toString());
            assertEquals("B2:B1048576", written[1].getSqref().get(0).toString());
        }
    }

    @Test
    void everyKindOfValidationWritesValidSpreadsheetML() throws Exception {
        List<Consumer<ColumnStyleConfig<Row, ?>>> columns = List.of(
                c -> c.validation(ExcelValidation.integerBetween(1, 10)),
                c -> c.validation(ExcelValidation.integerGreaterThan(1)),
                c -> c.validation(ExcelValidation.integerLessThan(9)),
                c -> c.validation(ExcelValidation.decimalBetween(0.5, 2.5)),
                c -> c.validation(ExcelValidation.textLength(1, 20)),
                c -> c.validation(ExcelValidation.dateRange(LocalDate.of(2024, 1, 1), LocalDate.of(2024, 12, 31))),
                c -> c.validation(ExcelValidation.listFromRange("$A$1:$A$5")),
                c -> c.validation(ExcelValidation.formula("LEN(A2)>3")),
                c -> c.validation(ExcelValidation.integerBetween(1, 2).errorTitle("Title").errorMessage("Message")),
                c -> c.dropdown("a", "b", "c"),
                c -> c.dropdown("say \"hi\"", "c"));
        for (int i = 0; i < columns.size(); i++) {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            var column = columns.get(i);
            ExcelWriter.<Row>create().column("name", Row::name, column::accept)
                    .write(List.of(new Row("a", 1))).writeTo(out);
            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                List<org.apache.xmlbeans.XmlError> errors = new ArrayList<>();
                boolean valid = workbook.getSheetAt(0).getCTWorksheet().validate(
                        new org.apache.xmlbeans.XmlOptions().setErrorListener(errors));
                assertTrue(valid, "validation " + i + " is not valid SpreadsheetML: " + errors);
            }
        }
    }
}
