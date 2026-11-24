package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** The column settings of a CSV writer ({@code nullValue}), the row number column, and the overloads that take them. */
class CsvColumnConfigTest {
    record Item(String name, String status, Integer qty) {}

    private static final List<Item> ITEMS = List.of(
            new Item("a", "ok", 1), new Item("b", null, null), new Item("c", "", 3));

    private static List<String> lines(CsvHandler handler) {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        String text = out.toString(StandardCharsets.UTF_8).replace("﻿", "");
        return text.isEmpty() ? List.of() : List.of(text.split("\r\n", -1)).subList(0, text.split("\r\n", -1).length - 1);
    }

    @Nested
    class NullValue {
        @Test
        void aMissingValueIsWrittenAsThePlaceholderAndOnlyThat() {
            var csv = lines(CsvWriter.<Item>create().column("Name", Item::name)
                    .column("Status", Item::status, c -> c.nullValue("N/A"))
                    .column("Qty", Item::qty, c -> c.nullValue(0))
                    .write(ITEMS));
            assertEquals(List.of("Name,Status,Qty", "a,ok,1", "b,N/A,0", "c,,3"), csv,
                    "an empty text is a value, not a missing one");
        }

        @Test
        void withoutItAMissingValueIsAnEmptyFieldAsBefore() {
            var csv = lines(CsvWriter.<Item>create().column("Name", Item::name).column("Status", Item::status).write(ITEMS));
            assertEquals("b,", csv.get(2));
        }

        @Test
        void aNullPlaceholderMeansEmpty() {
            var csv = lines(CsvWriter.<Item>create().column("Status", Item::status, c -> c.nullValue("x").nullValue(null)).write(ITEMS));
            assertEquals(List.of("Status", "ok", "", ""), csv);
        }

        @Test
        void thePlaceholderFollowsTheQuotingOfTheWriter() {
            var special = lines(CsvWriter.<Item>create().column("Status", Item::status, c -> c.nullValue("n,a \"x\"")).write(ITEMS));
            assertEquals("\"n,a \"\"x\"\"\"", special.get(2));
            var all = lines(CsvWriter.<Item>create().quoting(CsvQuoting.ALL).column("Qty", Item::qty, c -> c.nullValue(0)).write(ITEMS));
            assertEquals(List.of("\"Qty\"", "\"1\"", "\"0\"", "\"3\""), all);
            var nonNumeric = lines(CsvWriter.<Item>create().quoting(CsvQuoting.NON_NUMERIC)
                    .column("Qty", Item::qty, c -> c.nullValue(0)).column("S", Item::status, c -> c.nullValue("n")).write(ITEMS));
            assertEquals("0,\"n\"", nonNumeric.get(2));
        }

        @Test
        void theInjectionDefenseTreatsItLikeAnyOtherValue() {
            var defended = lines(CsvWriter.<Item>create().column("Status", Item::status, c -> c.nullValue("-")).write(ITEMS));
            assertEquals("'-", defended.get(2), "a placeholder that starts like a formula gets the quote, as documented");
            var plain = lines(CsvWriter.<Item>create().csvInjectionDefense(false).column("Status", Item::status, c -> c.nullValue("-")).write(ITEMS));
            assertEquals("-", plain.get(2));
            var safe = lines(CsvWriter.<Item>create().column("Status", Item::status, c -> c.nullValue("N/A")).write(ITEMS));
            assertEquals("N/A", safe.get(2));
        }

        @Test
        void aToleratedFailureIsAMissingValueToo() {
            var csv = lines(CsvWriter.<Item>create().writeErrorPolicy(io.github.dornol.excelkit.core.WriteErrorPolicy.LENIENT)
                    .column("Qty", i -> { if (i.qty() == null) throw new IllegalStateException("boom"); return i.qty(); },
                            c -> c.nullValue("?")).write(ITEMS));
            assertEquals(List.of("Qty", "1", "?", "3"), csv);
        }

        @Test
        void aFailureStillFailsWhenTheWriterIsToldTo() {
            var writer = CsvWriter.<Item>create().writeErrorPolicy(io.github.dornol.excelkit.core.WriteErrorPolicy.FAIL_FAST)
                    .column("Qty", i -> { throw new IllegalStateException("boom"); }, c -> c.nullValue("?"));
            assertThrows(CsvWriteException.class, () -> writer.write(ITEMS).writeTo(new ByteArrayOutputStream()));
        }

        @Test
        void eachColumnKeepsItsOwnPlaceholder() {
            var csv = lines(CsvWriter.<Item>create()
                    .column("A", Item::status, c -> c.nullValue("one"))
                    .column("B", Item::status, c -> c.nullValue("two"))
                    .column("C", Item::status)
                    .write(List.of(new Item("x", null, null))));
            assertEquals(List.of("A,B,C", "one,two,"), csv);
        }

        @Test
        void aRowFunctionColumnTakesItToo() {
            var csv = lines(CsvWriter.<Item>create()
                    .column("Pos", (item, cursor) -> item.qty() == null ? null : cursor.getCurrentTotal(), c -> c.nullValue("none"))
                    .write(ITEMS));
            assertEquals(List.of("Pos", "1", "none", "3"), csv);
        }
    }

    @Nested
    class RowNumber {
        @Test
        void countsFromOneAcrossTheRows() {
            var csv = lines(CsvWriter.<Item>create().rowNumberColumn("No.").column("Name", Item::name).write(ITEMS));
            assertEquals(List.of("No.,Name", "1,a", "2,b", "3,c"), csv);
        }

        @Test
        void worksFromAStreamAndAnyPositionAndWithTheOtherColumnKinds() {
            var csv = lines(CsvWriter.<Item>create().column("Name", Item::name).rowNumberColumn("No.")
                    .constColumn("K", "k").columnIf("Gone", false, Item::name).write(ITEMS.stream()));
            assertEquals(List.of("Name,No.,K", "a,1,k", "b,2,k", "c,3,k"), csv);
        }

        @Test
        void aMillionRowsCountToAMillion() {
            var out = new ByteArrayOutputStream();
            CsvWriter.<Integer>create().rowNumberColumn("n").write(java.util.stream.IntStream.range(0, 1_000_000).boxed()).writeTo(out);
            String text = out.toString(StandardCharsets.UTF_8);
            assertTrue(text.endsWith("\r\n1000000\r\n"), text.substring(text.length() - 20));
        }

        @Test
        void aHeaderOnlyFileWhenThereAreNoRows() {
            assertEquals(List.of("No."), lines(CsvWriter.<Item>create().rowNumberColumn("No.").write(List.of())));
        }

        @Test
        void aCsvWithTheColumnReadsBackWithTheNumbers() {
            var out = new ByteArrayOutputStream();
            CsvWriter.<Item>create().rowNumberColumn("No").column("Name", Item::name).write(ITEMS).writeTo(out);
            var seen = new ArrayList<String>();
            CsvReader.forMap().read(new ByteArrayInputStream(out.toByteArray()), r -> seen.add(r.data().get("No") + ":" + r.data().get("Name")));
            assertEquals(List.of("1:a", "2:b", "3:c"), seen);
        }
    }

    @Nested
    class Overloads {
        private static final Consumer<CsvColumnConfig> NA = c -> c.nullValue("NA");

        @Test
        void everyColumnMethodTakesTheSettings() {
            var csv = lines(CsvWriter.<Item>create()
                    .column("A", Item::status, NA)
                    .column("B", (item, cursor) -> item.status(), NA)
                    .columnIf("C", true, Item::status, NA)
                    .columnIf("D", true, (item, cursor) -> item.status(), NA)
                    .columnIf("Off", false, Item::status, NA)
                    .columnIf("Off2", false, (item, cursor) -> item.status(), NA)
                    .constColumn("E", null, NA)
                    .constColumnIf("F", true, null, NA)
                    .constColumnIf("Off3", false, "x", NA)
                    .write(List.of(new Item("x", null, null))));
            assertEquals(List.of("A,B,C,D,E,F", "NA,NA,NA,NA,NA,NA"), csv);
        }

        @Test
        void aNullConfigurerIsNoSettings() {
            var csv = lines(CsvWriter.<Item>create()
                    .column("A", Item::status, (Consumer<CsvColumnConfig>) null)
                    .column("B", (item, cursor) -> item.status(), null)
                    .columnIf("C", true, Item::status, null)
                    .constColumn("D", null, null)
                    .constColumnIf("E", true, "e", null)
                    .write(List.of(new Item("x", null, null))));
            assertEquals(List.of("A,B,C,D,E", ",,,,e"), csv);
        }

        @Test
        void aColumnNeedsAFunction() {
            var writer = CsvWriter.<Item>create();
            assertThrows(NullPointerException.class, () -> writer.column("A", (java.util.function.Function<Item, Object>) null, NA));
            assertThrows(NullPointerException.class, () -> writer.column("A", (io.github.dornol.excelkit.core.read.RowFunction<Item, Object>) null, NA));
        }

        @Test
        void streamsAndIterablesGiveTheSameFile() {
            var a = lines(CsvWriter.<Item>create().column("S", Item::status, NA).write(ITEMS));
            var b = lines(CsvWriter.<Item>create().column("S", Item::status, NA).write(ITEMS.stream()));
            assertEquals(a, b);
            assertEquals(Stream.of("S", "ok", "NA", "").toList(), a);
        }
    }
}
