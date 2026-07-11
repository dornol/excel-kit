package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class RoadmapApiTest {

    record Person(String name) {}

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
    void headerNormalizerAppliesToHeadersAndLookups() {
        List<Person> people = new ArrayList<>();
        CsvReader.<Person>mapping(row -> new Person(row.get("name").asString()))
                .headerNormalizer(value -> value.trim().toLowerCase(java.util.Locale.ROOT))
                .read(input(" Name \nAlice\n"), result -> people.add(result.data()));
        assertEquals(List.of(new Person("Alice")), people);
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
}
