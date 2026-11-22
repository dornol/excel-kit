package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.jupiter.api.Assertions.*;

class ReadWhileTest {

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
    void excelReadWhilePropagatesPredicateFailure() {
        var failure = new IllegalStateException("excel predicate failed");
        var thrown = assertThrows(IllegalStateException.class, () ->
                ExcelReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                        .readWhile(new ByteArrayInputStream(excelBytes()), result -> { throw failure; }));
        assertSame(failure, thrown);
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
