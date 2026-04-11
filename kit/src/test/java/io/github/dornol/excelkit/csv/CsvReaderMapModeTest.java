package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadResult;
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
    }
}
