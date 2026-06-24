package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ReadResult;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvReader#forMap()} — the v0.12.0 replacement for the
 * deleted {@code CsvMapReader}. Verifies header auto-discovery, dialect/delimiter/charset
 * support, and the mixed-mode runtime guards.
 */
class CsvReaderMapModeTest {

    @Nested
    @DisplayName("Factory")
    class Factory {

        @Test
        void forMap_returnsNonNull() {
            assertNotNull(CsvReader.forMap());
        }

        @Test
        void forMap_returnsNewInstanceEachCall() {
            assertNotSame(CsvReader.forMap(), CsvReader.forMap());
        }
    }

    @Nested
    @DisplayName("Reading — header auto-discover")
    class HeaderAutoDiscover {

        @Test
        void forMap_readsAllColumnsFromHeaderRow() {
            String csv = "Name,Age,City\nAlice,30,Seoul\nBob,25,Tokyo\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(2, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
            assertEquals("Seoul", results.get(0).get("City"));
            assertEquals("Bob", results.get(1).get("Name"));
            assertEquals("Tokyo", results.get(1).get("City"));
        }

        @Test
        void forMap_preservesHeaderOrder() {
            String csv = "Name,Age,City\nAlice,30,Seoul\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(List.of("Name", "Age", "City"),
                    new ArrayList<>(results.get(0).keySet()));
        }
    }

    @Nested
    @DisplayName("Fluent API compatibility")
    class FluentApi {

        @Test
        void forMap_dialect_TSV() {
            String tsv = "Name\tAge\nAlice\t30\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .dialect(CsvDialect.TSV)
                    .build(new ByteArrayInputStream(tsv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
        }

        @Test
        void forMap_delimiter_pipe() {
            String csv = "Name|Age\nAlice|30\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .delimiter('|')
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
        }

        @Test
        void forMap_charset() {
            // Write CJK content in EUC-KR, read it back via charset().
            String csv = "이름,도시\n앨리스,서울\n";
            byte[] bytes = csv.getBytes(java.nio.charset.Charset.forName("EUC-KR"));
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .charset(java.nio.charset.Charset.forName("EUC-KR"))
                    .build(new ByteArrayInputStream(bytes))
                    .read(r -> results.add(r.data()));

            assertEquals("앨리스", results.get(0).get("이름"));
            assertEquals("서울", results.get(0).get("도시"));
        }

        @Test
        void forMap_headerRowIndex_skipsRowsBeforeHeader() {
            String csv = "Title: report\n\nName,Age\nAlice,30\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .headerRowIndex(2)
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
        }

        @Test
        void forMap_onProgress_firesAtExpectedIntervals() {
            String csv = "n\n1\n2\n3\n4\n5\n6\n";
            AtomicLong lastCount = new AtomicLong(0);
            AtomicInteger callCount = new AtomicInteger(0);
            CsvReader.forMap()
                    .onProgress(2, (count, data) -> {
                        lastCount.set(count);
                        callCount.incrementAndGet();
                    })
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> {});

            assertEquals(3, callCount.get(), "onProgress fires at rows 2, 4, 6");
            assertEquals(6, lastCount.get());
        }

        @Test
        void forMap_readAsStream_producesSameRowsAsRead() {
            String csv = "Name,Age\nAlice,30\nBob,25\n";
            List<Map<String, String>> viaRead = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> viaRead.add(r.data()));

            List<Map<String, String>> viaStream = new ArrayList<>();
            try (Stream<ReadResult<Map<String, String>>> s = CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .readAsStream()) {
                s.forEach(r -> viaStream.add(r.data()));
            }

            assertEquals(viaRead, viaStream);
        }
    }

    @Nested
    @DisplayName("Mixed-mode runtime guards")
    class MixedModeGuards {

        @Test
        void column_setterOnly_throws() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.column((row, cell) -> {}));
        }

        @Test
        void column_nameAndSetter_throws() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.column("Name", (row, cell) -> {}));
        }

        @Test
        void columnAt_throws() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class,
                    () -> reader.columnAt(0, (row, cell) -> {}));
        }

        @Test
        void skipColumn_throws() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class, reader::skipColumn);
        }

        @Test
        void skipColumns_throws() {
            var reader = CsvReader.forMap();
            assertThrows(IllegalStateException.class, () -> reader.skipColumns(2));
        }

        @Test
        void setterMode_stillAllowsColumnCalls() {
            var reader = new CsvReader<>(Object::new, null);
            assertDoesNotThrow(() -> reader.column((row, cell) -> {}));
            assertDoesNotThrow(() -> reader.column("Name", (row, cell) -> {}));
            assertDoesNotThrow(() -> reader.columnAt(0, (row, cell) -> {}));
            assertDoesNotThrow(reader::skipColumn);
            assertDoesNotThrow(() -> reader.skipColumns(1));
        }

        @Test
        @DisplayName("guard message names the rejected method and mentions forMap()")
        void guardMessage_identifiesMethodName() {
            var reader = CsvReader.forMap();
            var ex = assertThrows(IllegalStateException.class,
                    () -> reader.column((row, cell) -> {}));
            assertTrue(ex.getMessage().contains("column(BiConsumer)"),
                    "error message should name the rejected method");
            assertTrue(ex.getMessage().contains("forMap()"),
                    "error message should point the user at forMap()");

            var ex2 = assertThrows(IllegalStateException.class,
                    () -> reader.columnAt(0, (r, c) -> {}));
            assertTrue(ex2.getMessage().contains("columnAt(int, BiConsumer)"));
        }
    }

    @Nested
    @DisplayName("Behavioral equivalence with deleted CsvMapReader")
    class BehavioralEquivalence {

        @Test
        @DisplayName("row with more cells than headers: the extras are ignored")
        void csvMapReader_moreDataColumnsThanHeaders() {
            String csv = "Name,Age\nAlice,30,extra1,extra2\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals(2, results.get(0).size(),
                    "map shape is bounded by header count, not data-cell count");
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("30", results.get(0).get("Age"));
        }

        @Test
        @DisplayName("blank header cell is treated as empty string, not null")
        void csvMapReader_blankHeaderCell() {
            // CSV parsers represent a missing field between delimiters as empty string,
            // which then becomes a header key of "" (non-null), so it's retained.
            // This documents the behavior vs Excel's null-header path.
            String csv = "Name,,City\nAlice,30,Seoul\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"));
            assertEquals("Seoul", results.get(0).get("City"));
            assertTrue(results.get(0).containsKey(""),
                    "CSV empty header becomes a '' key (different from Excel's null-header path)");
            assertEquals("30", results.get(0).get(""));
        }

        @Test
        @DisplayName("present but empty cell maps to empty string, not null")
        void csvMapReader_presentButEmptyCell_becomesEmptyString() {
            String csv = "Name,Age,City\nAlice,,Seoul\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertTrue(results.get(0).containsKey("Age"));
            assertNotNull(results.get(0).get("Age"),
                    "values in forMap() maps are never null (CellData coerces null → \"\")");
            assertEquals("", results.get(0).get("Age"));
            assertEquals("Seoul", results.get(0).get("City"));
        }

        @Test
        @DisplayName("duplicate headers: first occurrence wins (RowData.headerIndex behavior)")
        void csvMapReader_duplicateHeaders_firstWins() {
            // putIfAbsent semantics in buildHeaderIndexMap mean the first occurrence of a
            // duplicate header keeps its index; the duplicate column is not independently
            // addressable via get(name). The map output has one entry per unique header name.
            String csv = "A,B,A\nfirst,b-val,second\n";
            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals(2, results.get(0).size(),
                    "duplicate headers collapse to a single map key");
            assertEquals("first", results.get(0).get("A"));
            assertEquals("b-val", results.get(0).get("B"));
        }

        @Test
        @DisplayName("BOM-prefixed file: BOM is stripped from the first header, data reads correctly")
        void csvMapReader_bomPrefix_isStripped() {
            byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
            byte[] content = "Name,Age\nAlice,30\n".getBytes(StandardCharsets.UTF_8);
            byte[] withBom = new byte[bom.length + content.length];
            System.arraycopy(bom, 0, withBom, 0, bom.length);
            System.arraycopy(content, 0, withBom, bom.length, content.length);

            List<Map<String, String>> results = new ArrayList<>();
            CsvReader.forMap()
                    .build(new ByteArrayInputStream(withBom))
                    .read(r -> results.add(r.data()));

            assertEquals(1, results.size());
            assertEquals("Alice", results.get(0).get("Name"),
                    "first header key should be 'Name', not '\\uFEFFName'");
            assertEquals("30", results.get(0).get("Age"));
        }
    }
}
