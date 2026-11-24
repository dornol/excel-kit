package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A failed row reaches the caller in its {@link ReadResult}, value included, so it can be reported to whoever
 * uploaded the file. The log must not repeat that value: an upload is personal data and a bad file can fail on
 * every row.
 */
class ReadFailureLogContentTest {

    private static final String SECRET = "SECRET-4711";

    private static String captureStderr(Runnable action) {
        PrintStream original = System.err;
        ByteArrayOutputStream captured = new ByteArrayOutputStream();
        System.setErr(new PrintStream(captured, true));
        try {
            action.run();
        } finally {
            System.setErr(original);
        }
        return captured.toString();
    }

    private static ByteArrayInputStream csv() {
        return new ByteArrayInputStream(("Qty\n" + SECRET + "\n" + SECRET + "\n").getBytes(StandardCharsets.UTF_8));
    }

    public static class Item {
        int qty;
    }

    @Test
    void aCellThatCannotBeConverted_isReportedInTheResultButNotLogged() {
        List<ReadResult<Item>> results = new ArrayList<>();
        String log = captureStderr(() -> CsvReader.<Item>setter(Item::new)
                .column("Qty", (item, cell) -> item.qty = cell.asInt())
                .read(csv(), results::add));

        assertEquals(2, results.size());
        assertFalse(results.get(0).success());
        assertTrue(results.get(0).messages().toString().contains(SECRET), "the caller still gets the value");
        assertTrue(log.contains("Column mapping failed for 'Qty'"), log);
        assertFalse(log.contains(SECRET), log);
    }

    @Test
    void aRowMapperThatThrows_isReportedInTheResultButItsMessageIsNotLogged() {
        List<ReadResult<String>> results = new ArrayList<>();
        String log = captureStderr(() -> CsvReader.<String>mapping(row -> {
            throw new IllegalArgumentException("bad value " + row.get("Qty").asString());
        }).read(csv(), results::add));

        assertEquals(2, results.size());
        assertFalse(results.get(0).success());
        assertTrue(results.get(0).messages().toString().contains(SECRET));
        assertTrue(log.contains("Row mapping failed at file row"), log);
        assertFalse(log.contains(SECRET), log);
    }

    @Test
    void aCsvColumnThatKeepsFailing_logsOnceAtErrorAndTheRestAtDebug() {
        String log = captureStderr(() -> CsvWriter.<String>create()
                .writeErrorPolicy(io.github.dornol.excelkit.core.WriteErrorPolicy.LENIENT)
                .column("Boom", s -> {
                    throw new IllegalStateException("boom");
                })
                .write(java.util.stream.Stream.of("a", "b", "c"))
                .writeTo(OutputStream.nullOutputStream()));

        String marker = "Failed to apply function for column 'Boom'";
        int count = 0;
        for (int i = log.indexOf(marker); i >= 0; i = log.indexOf(marker, i + marker.length())) count++;
        assertEquals(1, count, log);
    }
}
