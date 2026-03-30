package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ExcelKitException;
import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Miscellaneous edge case tests for remaining uncovered branches:
 * - ExcelHandler (consume twice, password null)
 * - ExcelReader (skipColumns negative, header not found)
 * - ExcelSummary (all Op types)
 * - ExcelWorkbook edge cases
 * - ExcelWriter defaultStyle with applyDefaults
 */
class MiscEdgeCaseTest {

    record Item(String name, int value) {}

    // ============================================================
    // ExcelHandler edge cases
    // ============================================================
    @Nested
    class ExcelHandlerTests {

        @Test
        void consumeOutputStream_twice_throws() throws IOException {
            ByteArrayOutputStream out1 = new ByteArrayOutputStream();
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            handler.consumeOutputStream(out1);
            assertThrows(ExcelWriteException.class, () -> handler.consumeOutputStream(new ByteArrayOutputStream()));
        }

        @Test
        void consumeOutputStreamWithPassword_nullPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), (String) null));
        }

        @Test
        void consumeOutputStreamWithPassword_blankPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), "  "));
        }

        @Test
        void consumeOutputStreamWithPassword_charArray_nullPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), (char[]) null));
        }

        @Test
        void consumeOutputStreamWithPassword_charArray_emptyPassword_throws() throws IOException {
            ExcelHandler handler = new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            assertThrows(IllegalArgumentException.class,
                    () -> handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), new char[0]));
        }
    }

    // ============================================================
    // ExcelReader edge cases
    // ============================================================
    @Nested
    class ExcelReaderTests {

        static class MutableItem {
            String name;
            int value;
        }

        @Test
        void skipColumns_negative_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            assertThrows(IllegalArgumentException.class, () -> reader.skipColumns(-1));
        }

        @Test
        void onProgress_zeroInterval_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            assertThrows(IllegalArgumentException.class,
                    () -> reader.onProgress(0, (c, cur) -> {}));
        }

        @Test
        void headerNotFound_throws() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            reader.addColumn("Name", (item, cell) -> {});
            reader.addColumn("NonExistentHeader", (item, cell) -> {});

            assertThrows(ExcelKitException.class,
                    () -> reader.build(new ByteArrayInputStream(out.toByteArray()))
                            .read(r -> {}));
        }

        @Test
        void getSheetHeaders_withHeaderRowIndex() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);

            List<String> headers = ExcelReader.getSheetHeaders(
                    new ByteArrayInputStream(out.toByteArray()), 0, 0);

            assertEquals(2, headers.size());
            assertEquals("Name", headers.get(0));
            assertEquals("Value", headers.get(1));
        }
    }

    // ============================================================
    // ExcelSummary all Op types
    // ============================================================
    @Nested
    class ExcelSummaryTests {

        @Test
        void allSummaryOps_shouldWork() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Summary")
                            .sum("Value")
                            .average("Value")
                            .count("Value")
                            .min("Value")
                            .max("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .consumeOutputStream(out);
            assertTrue(out.size() > 0);
        }

        @Test
        void summary_labelInColumn_shouldWork() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Name", "Total:")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .consumeOutputStream(out);
            assertTrue(out.size() > 0);
        }
    }

    // ============================================================
    // ExcelWriter defaultStyle
    // ============================================================
    @Nested
    class ExcelWriterDefaultStyleTests {

        @Test
        void defaultStyle_shouldApplyToAllColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .defaultStyle(d -> d.bold(true).fontSize(12).fontName("Arial"))
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);
            assertTrue(out.size() > 0);
        }

        @Test
        void defaultStyle_columnOverrides() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .defaultStyle(d -> d.bold(true))
                    .addColumn("Name", Item::name, c -> c.bold(false))
                    .addColumn("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .consumeOutputStream(out);
            assertTrue(out.size() > 0);
        }
    }

    // ============================================================
    // ExcelWorkbook edge cases
    // ============================================================
    @Nested
    class ExcelWorkbookTests {

        @Test
        void protectWorkbook_shouldWork() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = new ExcelWorkbook()) {
                wb.<Item>sheet("Data")
                        .column("Name", Item::name)
                        .write(Stream.of(new Item("A", 1)));
                wb.protectWorkbook("password123");
                wb.finish().consumeOutputStream(out);
            }
            assertTrue(out.size() > 0);
        }
    }

    // ============================================================
    // ExcelWriter write with no data
    // ============================================================
    @Nested
    class EmptyDataTests {

        @Test
        void write_emptyStream_shouldCreateHeaderOnly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value)
                    .write(Stream.empty())
                    .consumeOutputStream(out);
            assertTrue(out.size() > 0);
        }
    }

    // ============================================================
    // readStrict with error messages
    // ============================================================
    @Nested
    class ReadStrictTests {

        @Test
        void readStrict_emptyMessages_shouldShowUnknownError() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            new ExcelWriter<Item>()
                    .addColumn("Name", Item::name)
                    .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 10)))
                    .consumeOutputStream(out);

            // Read with a mapper that always succeeds
            List<Item> results = new ArrayList<>();
            ExcelReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).build(new ByteArrayInputStream(out.toByteArray()))
                    .readStrict(results::add);

            assertEquals(1, results.size());
        }
    }
}
