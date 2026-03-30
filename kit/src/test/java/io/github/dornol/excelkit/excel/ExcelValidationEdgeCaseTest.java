package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelValidation} to cover:
 * - All validation types actually applied via write (DECIMAL, TEXT_LENGTH, DATE, FORMULA, LIST_FORMULA)
 * - Error dialog combinations (title only, message only, neither)
 * - showError false
 */
class ExcelValidationEdgeCaseTest {

    @Test
    void decimalBetween_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Price", s -> s, c -> c
                        .validation(ExcelValidation.decimalBetween(0.0, 999.99)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void textLength_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s, c -> c
                        .validation(ExcelValidation.textLength(1, 50)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void dateRange_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Date", s -> s, c -> c
                        .validation(ExcelValidation.dateRange(
                                LocalDate.of(2020, 1, 1), LocalDate.of(2030, 12, 31))))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void formula_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Custom", s -> s, c -> c
                        .validation(ExcelValidation.formula("LEN(A2)>0")))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void listFromRange_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Category", s -> s, c -> c
                        .validation(ExcelValidation.listFromRange("Sheet2!$A$1:$A$10")))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    // Error dialog combinations
    @Test
    void errorTitle_only_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)
                                .errorTitle("Bad Value")))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void errorMessage_only_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)
                                .errorMessage("Must be 1-100")))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void noErrorInfo_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void showError_false_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)
                                .showError(false)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void errorTitleAndMessage_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerBetween(1, 100)
                                .errorTitle("Error")
                                .errorMessage("Bad")))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void integerGreaterThan_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerGreaterThan(0)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }

    @Test
    void integerLessThan_appliedDuringWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Val", s -> s, c -> c
                        .validation(ExcelValidation.integerLessThan(100)))
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        assertTrue(out.size() > 0);
    }
}
