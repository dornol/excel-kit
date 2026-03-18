package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExcelChartConfig} builder and enums.
 */
class ExcelChartConfigTest {

    // ============================================================
    // Fluent chaining
    // ============================================================
    @Nested
    class FluentChainingTests {

        @Test
        void type_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.type(ExcelChartConfig.ChartType.LINE));
        }

        @Test
        void title_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.title("My Chart"));
        }

        @Test
        void categoryColumn_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.categoryColumn(0));
        }

        @Test
        void valueColumn_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.valueColumn(1, "Sales"));
        }

        @Test
        void position_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.position(0, 10, 8, 25));
        }

        @Test
        void categoryAxisTitle_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.categoryAxisTitle("X Axis"));
        }

        @Test
        void valueAxisTitle_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.valueAxisTitle("Y Axis"));
        }

        @Test
        void legendPosition_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.legendPosition(ExcelChartConfig.LegendPosition.BOTTOM));
        }

        @Test
        void showDataLabels_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.showDataLabels(true));
        }

        @Test
        void barGrouping_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.barGrouping(ExcelChartConfig.BarGrouping.STACKED));
        }

        @Test
        void barDirection_returnsSameInstance() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertSame(c, c.barDirection(ExcelChartConfig.BarDirection.HORIZONTAL));
        }

        @Test
        void fullChain_works() {
            ExcelChartConfig c = new ExcelChartConfig()
                    .type(ExcelChartConfig.ChartType.BAR)
                    .title("Revenue")
                    .categoryColumn(0)
                    .valueColumn(1, "Q1")
                    .valueColumn(2, "Q2")
                    .position(0, 10, 10, 25)
                    .categoryAxisTitle("Product")
                    .valueAxisTitle("Amount")
                    .legendPosition(ExcelChartConfig.LegendPosition.RIGHT)
                    .barGrouping(ExcelChartConfig.BarGrouping.STACKED)
                    .barDirection(ExcelChartConfig.BarDirection.VERTICAL)
                    .showDataLabels(true);
            assertNotNull(c);
        }

        @Test
        void multipleValueColumns_accumulate() {
            ExcelChartConfig c = new ExcelChartConfig()
                    .valueColumn(1, "A")
                    .valueColumn(2, "B")
                    .valueColumn(3, "C");
            assertNotNull(c);
        }
    }

    // ============================================================
    // ChartType enum
    // ============================================================
    @Nested
    class ChartTypeTests {

        @Test
        void allTypesExist() {
            assertEquals(6, ExcelChartConfig.ChartType.values().length);
            assertNotNull(ExcelChartConfig.ChartType.BAR);
            assertNotNull(ExcelChartConfig.ChartType.LINE);
            assertNotNull(ExcelChartConfig.ChartType.PIE);
            assertNotNull(ExcelChartConfig.ChartType.SCATTER);
            assertNotNull(ExcelChartConfig.ChartType.AREA);
            assertNotNull(ExcelChartConfig.ChartType.DOUGHNUT);
        }

        @Test
        void valueOf_roundTrips() {
            for (ExcelChartConfig.ChartType t : ExcelChartConfig.ChartType.values()) {
                assertEquals(t, ExcelChartConfig.ChartType.valueOf(t.name()));
            }
        }
    }

    // ============================================================
    // LegendPosition enum
    // ============================================================
    @Nested
    class LegendPositionTests {

        @Test
        void allPositionsExist() {
            assertEquals(5, ExcelChartConfig.LegendPosition.values().length);
            assertNotNull(ExcelChartConfig.LegendPosition.BOTTOM);
            assertNotNull(ExcelChartConfig.LegendPosition.LEFT);
            assertNotNull(ExcelChartConfig.LegendPosition.RIGHT);
            assertNotNull(ExcelChartConfig.LegendPosition.TOP);
            assertNotNull(ExcelChartConfig.LegendPosition.TOP_RIGHT);
        }
    }

    // ============================================================
    // BarGrouping enum
    // ============================================================
    @Nested
    class BarGroupingTests {

        @Test
        void allGroupingsExist() {
            assertEquals(3, ExcelChartConfig.BarGrouping.values().length);
            assertNotNull(ExcelChartConfig.BarGrouping.STANDARD);
            assertNotNull(ExcelChartConfig.BarGrouping.STACKED);
            assertNotNull(ExcelChartConfig.BarGrouping.PERCENT_STACKED);
        }
    }

    // ============================================================
    // BarDirection enum
    // ============================================================
    @Nested
    class BarDirectionTests {

        @Test
        void allDirectionsExist() {
            assertEquals(2, ExcelChartConfig.BarDirection.values().length);
            assertNotNull(ExcelChartConfig.BarDirection.VERTICAL);
            assertNotNull(ExcelChartConfig.BarDirection.HORIZONTAL);
        }
    }

    // ============================================================
    // Input validation
    // ============================================================
    @Nested
    class ValidationTests {

        @Test
        void categoryColumn_negative_throws() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertThrows(IllegalArgumentException.class, () -> c.categoryColumn(-1));
        }

        @Test
        void valueColumn_negative_throws() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertThrows(IllegalArgumentException.class, () -> c.valueColumn(-1, "Sales"));
        }

        @Test
        void categoryColumn_zero_accepted() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertDoesNotThrow(() -> c.categoryColumn(0));
        }

        @Test
        void valueColumn_zero_accepted() {
            ExcelChartConfig c = new ExcelChartConfig();
            assertDoesNotThrow(() -> c.valueColumn(0, "Sales"));
        }
    }

    // ============================================================
    // Integration: chart via ExcelWriter
    // ============================================================
    @Nested
    class IntegrationTests {

        @Test
        void barChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Name", s -> s);
            writer.addColumn("Value", s -> s.length(), c -> c.type(ExcelDataType.INTEGER));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.BAR)
                    .title("Test")
                    .categoryColumn(0)
                    .valueColumn(1, "Value"));
            assertNotNull(writer);
        }

        @Test
        void lineChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("X", s -> s);
            writer.addColumn("Y", s -> s.length(), c -> c.type(ExcelDataType.INTEGER));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.LINE)
                    .categoryColumn(0)
                    .valueColumn(1, "Y"));
            assertNotNull(writer);
        }

        @Test
        void pieChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Category", s -> s);
            writer.addColumn("Amount", s -> s.length(), c -> c.type(ExcelDataType.INTEGER));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.PIE)
                    .categoryColumn(0)
                    .valueColumn(1, "Amount"));
            assertNotNull(writer);
        }

        @Test
        void scatterChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("X", s -> (double) s.length(), c -> c.type(ExcelDataType.DOUBLE));
            writer.addColumn("Y", s -> (double) s.hashCode(), c -> c.type(ExcelDataType.DOUBLE));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.SCATTER)
                    .categoryColumn(0)
                    .valueColumn(1, "Y"));
            assertNotNull(writer);
        }

        @Test
        void areaChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Month", s -> s);
            writer.addColumn("Sales", s -> s.length(), c -> c.type(ExcelDataType.INTEGER));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.AREA)
                    .categoryColumn(0)
                    .valueColumn(1, "Sales"));
            assertNotNull(writer);
        }

        @Test
        void doughnutChart_viaExcelWriter_doesNotThrow() {
            ExcelWriter<String> writer = new ExcelWriter<>();
            writer.addColumn("Type", s -> s);
            writer.addColumn("Count", s -> s.length(), c -> c.type(ExcelDataType.INTEGER));
            writer.chart(c -> c
                    .type(ExcelChartConfig.ChartType.DOUGHNUT)
                    .categoryColumn(0)
                    .valueColumn(1, "Count"));
            assertNotNull(writer);
        }
    }
}
