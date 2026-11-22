package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
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

class InputOwnershipTest {

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

        AtomicBoolean excelClosed = new AtomicBoolean();
        Stream<Person> excelStream = Stream.of(new Person("C"))
                .onClose(() -> excelClosed.set(true));
        ExcelWriter.<Person>create().column("Name", Person::name)
                .write(excelStream).writeTo(new ByteArrayOutputStream());
        assertFalse(excelClosed.get());
        excelStream.close();
        assertTrue(excelClosed.get());
    }

    record Person(String name) {}

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
