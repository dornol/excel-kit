package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExcelSummary} DSL configuration.
 */
class ExcelSummaryTest {

    // ============================================================
    // Fluent chaining
    // ============================================================
    @Nested
    class FluentChainingTests {

        @Test
        void sum_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.sum("Amount"));
        }

        @Test
        void average_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.average("Score"));
        }

        @Test
        void count_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.count("ID"));
        }

        @Test
        void min_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.min("Price"));
        }

        @Test
        void max_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.max("Price"));
        }

        @Test
        void label_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.label("Total"));
        }

        @Test
        void labelWithColumn_returnsSameInstance() {
            ExcelSummary s = new ExcelSummary();
            assertSame(s, s.label("Name", "Total"));
        }

        @Test
        void chainMultipleOps() {
            ExcelSummary s = new ExcelSummary()
                    .label("Summary")
                    .sum("Amount")
                    .average("Score")
                    .count("ID")
                    .min("Price")
                    .max("Price");

            // Should not throw — just verify chaining works
            assertNotNull(s);
        }
    }

    // ============================================================
    // toAfterDataWriter conversion
    // ============================================================
    @Nested
    class ToAfterDataWriterTests {

        @Test
        void toAfterDataWriter_returnsNonNull() {
            ExcelSummary s = new ExcelSummary().sum("Amount");
            AfterDataWriter writer = s.toAfterDataWriter();
            assertNotNull(writer);
        }
    }

    // ============================================================
    // Op enum
    // ============================================================
    @Nested
    class OpEnumTests {

        @Test
        void allOpsExist() {
            assertEquals(5, ExcelSummary.Op.values().length);
            assertNotNull(ExcelSummary.Op.SUM);
            assertNotNull(ExcelSummary.Op.AVERAGE);
            assertNotNull(ExcelSummary.Op.COUNT);
            assertNotNull(ExcelSummary.Op.MIN);
            assertNotNull(ExcelSummary.Op.MAX);
        }

        @Test
        void valueOf_roundTrips() {
            for (ExcelSummary.Op op : ExcelSummary.Op.values()) {
                assertEquals(op, ExcelSummary.Op.valueOf(op.name()));
            }
        }
    }
}
