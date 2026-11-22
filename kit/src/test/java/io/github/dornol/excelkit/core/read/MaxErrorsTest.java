package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class MaxErrorsTest {

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

    record Person(String name) {}

    private static ByteArrayInputStream input(String csv) {
        return new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
    }
}
