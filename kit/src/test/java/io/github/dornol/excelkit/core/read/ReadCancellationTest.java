package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import static org.junit.jupiter.api.Assertions.*;

class ReadCancellationTest {

    @Test void cancellationStopsBetweenRows() {
        AtomicBoolean cancelled = new AtomicBoolean();
        var values = new java.util.ArrayList<Person>();
        ReadSummary summary = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .cancellationToken(cancelled::get)
                .readWithSummary(csv("Name\nA\nB\n"), result -> {
                    values.add(result.data());
                    cancelled.set(true);
                });
        assertEquals(List.of(new Person("A")), values);
        assertTrue(summary.stoppedEarly());
    }

    record Person(String name) {}

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
