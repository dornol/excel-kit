package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import static org.junit.jupiter.api.Assertions.*;

class ReadSessionTest {
    record Person(String name) {}

    @Test
    void sessionReadsAllAndDoesNotCloseCallerStream() {
        TrackingInputStream input = new TrackingInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
        List<Person> people;
        try (ReadSession<Person> session = CsvReader.<Person>mapping(
                row -> new Person(row.get("Name").asString())).open(input)) {
            people = session.readAll();
        }
        assertEquals(List.of(new Person("Alice")), people);
        assertFalse(input.closed.get());
    }

    @Test
    void sourceStreamIsClosedAfterSessionInitialization() {
        TrackingInputStream input = new TrackingInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
        ReadSession<Person> session = CsvReader.<Person>mapping(
                row -> new Person(row.get("Name").asString())).open(() -> input);
        assertTrue(input.closed.get());
        session.close();
    }

    private static final class TrackingInputStream extends ByteArrayInputStream {
        private final AtomicBoolean closed = new AtomicBoolean();

        private TrackingInputStream(byte[] bytes) { super(bytes); }

        @Override public void close() throws IOException {
            closed.set(true);
            super.close();
        }
    }
}
