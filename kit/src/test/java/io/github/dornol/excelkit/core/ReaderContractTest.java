package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Verifies the format-independent reader contract for CSV and Excel. */
class ReaderContractTest {

    @Test
    void maxRowsStopsBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            AtomicInteger count = new AtomicInteger();
            read(excel, reader -> reader.maxRows(1), result -> count.incrementAndGet());
            assertEquals(1, count.get(), format(excel));
        }
    }

    @Test
    void readWhileStopsBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            AtomicInteger count = new AtomicInteger();
            ReadSummary summary = readWhile(excel, result -> count.incrementAndGet() < 1);
            assertEquals(1, count.get(), format(excel));
            assertTrue(summary.stoppedEarly(), format(excel));
        }
    }

    @Test
    void cellCharacterLimitStopsBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            assertThrows(ReadLimitExceededException.class, () ->
                    read(excel, reader -> reader.limits(new ReadLimits(-1, -1, -1, 2)), result -> fail()),
                    format(excel));
        }
    }

    @Test
    void cancellationStopsBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            AtomicBoolean cancelled = new AtomicBoolean();
            AtomicInteger count = new AtomicInteger();
            ReadSummary summary = readWithToken(excel, cancelled, result -> {
                count.incrementAndGet();
                cancelled.set(true);
            });
            assertEquals(1, count.get(), format(excel));
            assertTrue(summary.stoppedEarly(), format(excel));
        }
    }

    @Test
    void callerInputStreamRemainsOpenForBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            AtomicBoolean closed = new AtomicBoolean();
            InputStream input = new ByteArrayInputStream(bytes(excel)) {
                @Override
                public void close() {
                    closed.set(true);
                }
            };
            read(excel, reader -> {}, input, result -> {});
            assertFalse(closed.get(), format(excel));
        }
    }

    private static void read(boolean excel, Consumer<AbstractReader<String, ?>> configure,
                             Consumer<ReadResult<String>> consumer) {
        read(excel, configure, new ByteArrayInputStream(bytes(excel)), consumer);
    }

    @SuppressWarnings({"rawtypes", "unchecked"})
    private static void read(boolean excel, Consumer<AbstractReader<String, ?>> configure,
                             InputStream input, Consumer<ReadResult<String>> consumer) {
        if (excel) {
            ExcelReader<String> reader = ExcelReader.mapping(row -> row.get("Name").asString());
            configure.accept((AbstractReader) reader);
            reader.read(input, consumer);
        } else {
            CsvReader<String> reader = CsvReader.mapping(row -> row.get("Name").asString());
            configure.accept((AbstractReader) reader);
            reader.read(input, consumer);
        }
    }

    private static ReadSummary readWhile(boolean excel, java.util.function.Predicate<ReadResult<String>> predicate) {
        InputStream input = new ByteArrayInputStream(bytes(excel));
        if (excel) return ExcelReader.<String>mapping(row -> row.get("Name").asString()).readWhile(input, predicate);
        return CsvReader.<String>mapping(row -> row.get("Name").asString()).readWhile(input, predicate);
    }

    private static ReadSummary readWithToken(boolean excel, AtomicBoolean cancelled,
                                             Consumer<ReadResult<String>> consumer) {
        InputStream input = new ByteArrayInputStream(bytes(excel));
        if (excel) {
            return ExcelReader.<String>mapping(row -> row.get("Name").asString())
                    .cancellationToken(cancelled::get).readWithSummary(input, consumer);
        }
        return CsvReader.<String>mapping(row -> row.get("Name").asString())
                .cancellationToken(cancelled::get).readWithSummary(input, consumer);
    }

    private static byte[] bytes(boolean excel) {
        if (!excel) return "Name\nLong\nB\n".getBytes(StandardCharsets.UTF_8);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("Name", value -> value)
                .write(java.util.List.of("Long", "B")).writeTo(output);
        return output.toByteArray();
    }

    private static String format(boolean excel) {
        return excel ? "Excel" : "CSV";
    }

    private static void fail() {
        throw new AssertionError("oversized cell must not be delivered");
    }
}
