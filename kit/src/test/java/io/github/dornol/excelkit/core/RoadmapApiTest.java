package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class RoadmapApiTest {

    record Person(String name) {}
    static final class MutablePerson { String name; String later; }

    @Test
    void directReadDoesNotCloseCallerInput() {
        AtomicBoolean closed = new AtomicBoolean();
        var input = new ByteArrayInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8)) {
            @Override public void close() { closed.set(true); }
        };

        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .read(input, result -> assertTrue(result.success()));

        assertFalse(closed.get());
    }

    @Test
    void sourceReadClosesOwnedInput() {
        AtomicBoolean closed = new AtomicBoolean();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .read(() -> new ByteArrayInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8)) {
                    @Override public void close() { closed.set(true); }
                }, result -> {});
        assertTrue(closed.get());
    }

    @Test
    void readWhileStopsNormally() {
        AtomicInteger count = new AtomicInteger();
        CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .readWhile(input("Name\nA\nB\nC\n"), result -> count.incrementAndGet() < 2);
        assertEquals(2, count.get());
    }

    @Test
    void readWhilePropagatesPredicateFailure() {
        var failure = new IllegalStateException("stop failed");
        var thrown = assertThrows(IllegalStateException.class, () ->
                CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                        .readWhile(input("Name\nA\n"), result -> { throw failure; }));
        assertSame(failure, thrown);
    }

    @Test
    void maxErrorsAbortsAfterAllowedErrors() {
        List<ReadResult<Person>> delivered = new ArrayList<>();
        ReadAbortException error = assertThrows(ReadAbortException.class, () ->
                CsvReader.<Person>mapping(row -> {
                            throw new IllegalArgumentException("bad");
                        })
                        .maxErrors(1)
                        .read(input("Name\nA\nB\n"), delivered::add));

        assertEquals(1, delivered.size());
        assertEquals(ReadAbortReason.MAX_ERRORS_EXCEEDED, error.reason());
        assertEquals(2, error.errorCount());
    }

    @Test
    void maxErrorsZeroAbortsBeforeDeliveringFirstFailure() {
        List<ReadResult<Person>> delivered = new ArrayList<>();
        ReadAbortException error = assertThrows(ReadAbortException.class, () ->
                CsvReader.<Person>mapping(row -> { throw new IllegalArgumentException("bad"); })
                        .maxErrors(0)
                        .read(input("Name\nA\n"), delivered::add));
        assertTrue(delivered.isEmpty());
        assertEquals(1, error.errorCount());
    }

    @Test
    void headerNormalizerAppliesToHeadersAndLookups() {
        List<Person> people = new ArrayList<>();
        CsvReader.<Person>mapping(row -> new Person(row.get("name").asString()))
                .headerNormalizer(value -> value.trim().toLowerCase(java.util.Locale.ROOT))
                .read(input(" Name \nAlice\n"), result -> people.add(result.data()));
        assertEquals(List.of(new Person("Alice")), people);
    }

    @Test
    void normalizedDuplicateHeadersHonorPolicy() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .headerNormalizer(value -> value.trim().toLowerCase(java.util.Locale.ROOT))
                .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                .read(input("Name, name \nA,B\n"), result -> {}));
    }

    @Test
    void executionUsesColumnSnapshot() {
        CsvReader<MutablePerson> reader = new CsvReader<>(MutablePerson::new)
                .column("Name", (person, cell) -> person.name = cell.asString());
        List<ReadResult<MutablePerson>> results = new ArrayList<>();
        reader.read(input("Name,Later\nA,1\nB,2\n"), result -> {
            results.add(result);
            if (results.size() == 1) {
                reader.column("Later", (person, cell) -> person.later = cell.asString());
            }
        });
        assertEquals(2, results.size());
        assertNull(results.get(1).data().later);
    }

    @Test
    void excelMetadataDoesNotCloseCallerInputButClosesSource() throws Exception {
        byte[] workbook = excelBytes();
        AtomicBoolean callerClosed = new AtomicBoolean();
        var callerInput = new ByteArrayInputStream(workbook) {
            @Override public void close() { callerClosed.set(true); }
        };
        assertEquals("Data", ExcelReader.getSheetNames(callerInput).get(0).name());
        assertFalse(callerClosed.get());

        AtomicBoolean sourceClosed = new AtomicBoolean();
        ExcelReader.getSheetHeaders(() -> new ByteArrayInputStream(workbook) {
            @Override public void close() { sourceClosed.set(true); }
        }, 0, 0);
        assertTrue(sourceClosed.get());
    }

    @Test
    void excelReadWhilePropagatesPredicateFailure() {
        var failure = new IllegalStateException("excel predicate failed");
        var thrown = assertThrows(IllegalStateException.class, () ->
                ExcelReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                        .readWhile(new ByteArrayInputStream(excelBytes()), result -> { throw failure; }));
        assertSame(failure, thrown);
    }

    @Test
    void pathAndSourceOverloadsAreSymmetric() throws Exception {
        Path csv = Files.createTempFile("excel-kit-reader", ".csv");
        try {
            Files.writeString(csv, "Name\nA\nB\n");
            List<Person> strict = new ArrayList<>();
            CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                    .readStrict(csv, strict::add);
            assertEquals(2, strict.size());
            assertTrue(Files.exists(csv), "Path input must remain caller-owned");

            AtomicInteger count = new AtomicInteger();
            CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                    .readWhile((InputStreamSource) () -> Files.newInputStream(csv),
                            result -> count.incrementAndGet() < 1);
            assertEquals(1, count.get());
        } finally {
            Files.deleteIfExists(csv);
        }
    }

    @Test
    void iterableWriteDoesNotCloseCallerStream() throws Exception {
        AtomicBoolean closed = new AtomicBoolean();
        Stream<Person> stream = Stream.of(new Person("A")).onClose(() -> closed.set(true));
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        CsvWriter.<Person>create().column("Name", Person::name).write(stream).writeTo(output);
        assertFalse(closed.get());
        stream.close();
        assertTrue(closed.get());

        output.reset();
        CsvWriter.<Person>create().column("Name", Person::name)
                .write(List.of(new Person("B"))).writeTo(output);
        assertTrue(output.toString(StandardCharsets.UTF_8).contains("B"));
    }

    private static ByteArrayInputStream input(String csv) {
        return new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
    }

    private static byte[] excelBytes() {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create().sheetName("Data").column("Name", Person::name)
                .write(List.of(new Person("A"))).writeTo(output);
        return output.toByteArray();
    }
}
