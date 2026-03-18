package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExcelValidation} DSL and factory methods.
 */
class ExcelValidationTest {

    // ============================================================
    // Factory methods — verify creation does not throw
    // ============================================================
    @Nested
    class FactoryMethodTests {

        @Test
        void integerBetween_createsValidation() {
            ExcelValidation v = ExcelValidation.integerBetween(1, 100);
            assertNotNull(v);
        }

        @Test
        void integerGreaterThan_createsValidation() {
            ExcelValidation v = ExcelValidation.integerGreaterThan(0);
            assertNotNull(v);
        }

        @Test
        void integerLessThan_createsValidation() {
            ExcelValidation v = ExcelValidation.integerLessThan(50);
            assertNotNull(v);
        }

        @Test
        void decimalBetween_createsValidation() {
            ExcelValidation v = ExcelValidation.decimalBetween(0.0, 99.9);
            assertNotNull(v);
        }

        @Test
        void textLength_createsValidation() {
            ExcelValidation v = ExcelValidation.textLength(1, 255);
            assertNotNull(v);
        }

        @Test
        void dateRange_createsValidation() {
            ExcelValidation v = ExcelValidation.dateRange(
                    LocalDate.of(2024, 1, 1), LocalDate.of(2025, 12, 31));
            assertNotNull(v);
        }

        @Test
        void listFromRange_createsValidation() {
            ExcelValidation v = ExcelValidation.listFromRange("Sheet2!$A$1:$A$10");
            assertNotNull(v);
        }

        @Test
        void formula_createsValidation() {
            ExcelValidation v = ExcelValidation.formula("AND(A1>0,A1<100)");
            assertNotNull(v);
        }
    }

    // ============================================================
    // Fluent chaining
    // ============================================================
    @Nested
    class FluentChainingTests {

        @Test
        void errorTitle_returnsSameInstance() {
            ExcelValidation v = ExcelValidation.integerBetween(1, 10);
            assertSame(v, v.errorTitle("Error"));
        }

        @Test
        void errorMessage_returnsSameInstance() {
            ExcelValidation v = ExcelValidation.integerBetween(1, 10);
            assertSame(v, v.errorMessage("Invalid value"));
        }

        @Test
        void showError_returnsSameInstance() {
            ExcelValidation v = ExcelValidation.integerBetween(1, 10);
            assertSame(v, v.showError(false));
        }

        @Test
        void fullChain_works() {
            ExcelValidation v = ExcelValidation.integerBetween(1, 100)
                    .errorTitle("Out of Range")
                    .errorMessage("Please enter 1-100")
                    .showError(true);
            assertNotNull(v);
        }
    }

    // ============================================================
    // ValidationType enum
    // ============================================================
    @Nested
    class ValidationTypeTests {

        @Test
        void allTypesExist() {
            assertEquals(6, ExcelValidation.ValidationType.values().length);
            assertNotNull(ExcelValidation.ValidationType.INTEGER);
            assertNotNull(ExcelValidation.ValidationType.DECIMAL);
            assertNotNull(ExcelValidation.ValidationType.TEXT_LENGTH);
            assertNotNull(ExcelValidation.ValidationType.DATE);
            assertNotNull(ExcelValidation.ValidationType.FORMULA);
            assertNotNull(ExcelValidation.ValidationType.LIST_FORMULA);
        }

        @Test
        void valueOf_roundTrips() {
            for (ExcelValidation.ValidationType t : ExcelValidation.ValidationType.values()) {
                assertEquals(t, ExcelValidation.ValidationType.valueOf(t.name()));
            }
        }
    }

    // ============================================================
    // Integration: apply via ExcelWriter addColumn
    // ============================================================
    @Nested
    class IntegrationTests {

        @Test
        void integerBetween_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Value", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(1, 150)));
            assertNotNull(writer);
        }

        @Test
        void decimalBetween_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Price", s -> s, c -> c
                    .validation(ExcelValidation.decimalBetween(0, 9999.99)));
            assertNotNull(writer);
        }

        @Test
        void textLength_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Name", s -> s, c -> c
                    .validation(ExcelValidation.textLength(1, 50)));
            assertNotNull(writer);
        }

        @Test
        void dateRange_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Date", s -> s, c -> c
                    .validation(ExcelValidation.dateRange(
                            LocalDate.of(2020, 1, 1), LocalDate.of(2030, 12, 31))));
            assertNotNull(writer);
        }

        @Test
        void formula_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Custom", s -> s, c -> c
                    .validation(ExcelValidation.formula("LEN(A2)>0")));
            assertNotNull(writer);
        }

        @Test
        void listFromRange_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Category", s -> s, c -> c
                    .validation(ExcelValidation.listFromRange("Lists!$A$1:$A$5")));
            assertNotNull(writer);
        }

        @Test
        void validationWithErrorInfo_appliedViaColumn_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Age", s -> s, c -> c
                    .validation(ExcelValidation.integerBetween(0, 150)
                            .errorTitle("Invalid Age")
                            .errorMessage("Age must be between 0 and 150")));
            assertNotNull(writer);
        }
    }
}
