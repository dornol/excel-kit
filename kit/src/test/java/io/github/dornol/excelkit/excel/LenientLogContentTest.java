package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Lenient failures are logged with the column and position, never with the row's or the cell's content. */
class LenientLogContentTest {

    private record Person(String name) {
        @Override
        public String toString() {
            return "Person[secret-name=" + name + "]";
        }
    }

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

    @Test
    void failedValueFunction_logsColumnAndRowNumberButNotTheRow() {
        String log = captureStderr(() -> ExcelWriter.<Person>create()
                .writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT)
                .column("Boom", p -> {
                    throw new IllegalStateException("boom");
                })
                .write(List.of(new Person("alice")))
                .writeTo(OutputStream.nullOutputStream()));

        assertTrue(log.contains("'Boom'"), log);
        assertFalse(log.contains("secret-name"), log);
    }

    @Test
    void aColumnThatKeepsFailing_logsOnceAtErrorAndTheRestAtDebug() {
        String log = captureStderr(() -> ExcelWriter.<Person>create()
                .writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT)
                .column("Boom", p -> {
                    throw new IllegalStateException("boom");
                })
                .write(List.of(new Person("a"), new Person("b"), new Person("c")))
                .writeTo(OutputStream.nullOutputStream()));

        assertEquals(1, occurrences(log, "applyFunction exception caught for column 'Boom'"), log);
        assertTrue(log.contains("further failures of this column are logged at DEBUG"), log);
    }

    private static int occurrences(String text, String part) {
        int count = 0;
        for (int i = text.indexOf(part); i >= 0; i = text.indexOf(part, i + part.length())) count++;
        return count;
    }
}
