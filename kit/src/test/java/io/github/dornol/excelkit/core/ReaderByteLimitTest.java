package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import static org.junit.jupiter.api.Assertions.*;

class ReaderByteLimitTest {
    @TempDir Path directory;

    @Test
    void bothReadersStopAtExactlyOneByteBeyondTheLimitAndKeepCallerStreamsOpen() {
        for (boolean excel : List.of(false, true)) {
            var input = new TrackedInput(1_000_000);
            var error = assertThrows(ReadLimitExceededException.class, () -> open(excel, input, 100));
            assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, error.limit());
            assertEquals(100, error.configured());
            assertEquals(101, error.actual());
            assertEquals(101, input.consumed);
            assertFalse(input.closed);
        }
    }

    @Test
    void exactBoundaryEmptyZeroUnlimitedAndLongMaxLimitsAreAccepted() throws Exception {
        for (boolean excel : List.of(false, true)) {
            for (long maximum : new long[]{10, -1, Long.MAX_VALUE}) {
                var input = new TrackedInput(10);
                try (var session = open(excel, input, maximum)) {
                    assertEquals(10, input.consumed);
                }
                assertFalse(input.closed);
            }
            try (var session = open(excel, new TrackedInput(0), 0)) {
                assertNotNull(session);
            }
            var input = new TrackedInput(100);
            var error = assertThrows(ReadLimitExceededException.class, () -> open(excel, input, 0));
            assertEquals(1, error.actual());
            assertEquals(1, input.consumed);
        }
    }

    @Test
    void ownedSourceIsClosedWhenLimitIsExceeded() {
        for (boolean excel : List.of(false, true)) {
            var input = new TrackedInput(100);
            InputStreamSource source = () -> input;
            assertThrows(ReadLimitExceededException.class, () -> {
                if (excel) ExcelReader.forMap().limits(limits(3)).open(source);
                else CsvReader.forMap().limits(limits(3)).open(source);
            });
            assertEquals(4, input.consumed);
            assertTrue(input.closed);
        }
    }

    @Test
    void singleByteBulkAndSkipShareOneBudgetAndStopAfterFailure() throws Exception {
        var input = new TrackedInput(100);
        var limited = new ProbeReader().limits(limits(4)).limitInput(input);
        assertEquals(0, limited.read());
        assertEquals(2, limited.read(new byte[2]));
        assertEquals(1, limited.skip(1));
        var failure = assertThrows(ReadLimitExceededException.class, limited::read);
        assertEquals(5, failure.actual());
        assertEquals(5, input.consumed);
        assertSame(failure, assertThrows(ReadLimitExceededException.class, () -> limited.read(new byte[32])));
        assertSame(failure, assertThrows(ReadLimitExceededException.class, () -> limited.skip(32)));
        assertEquals(5, input.consumed);
        limited.close();
        assertFalse(input.closed);
    }

    @Test
    void skipAndReadAllBytesCannotConsumeBeyondTheSentinelByte() {
        for (boolean skip : List.of(false, true)) {
            var input = new TrackedInput(1000);
            var limited = new ProbeReader().limits(limits(3)).limitInput(input);
            var error = assertThrows(ReadLimitExceededException.class, () -> {
                if (skip) limited.skip(Long.MAX_VALUE);
                else limited.readAllBytes();
            });
            assertEquals(4, input.consumed);
            assertEquals(4, error.actual());
        }
    }

    @Test
    void zeroLengthReadsDoNotConsumeAndInvalidRangesAreRejected() throws Exception {
        var input = new TrackedInput(100);
        var limited = new ProbeReader().limits(limits(0)).limitInput(input);
        assertEquals(0, limited.read(new byte[4], 4, 0));
        assertEquals(0, limited.skip(-1));
        assertThrows(IndexOutOfBoundsException.class, () -> limited.read(new byte[4], 4, 1));
        assertEquals(0, input.consumed);
    }

    @Test
    void observedCountSaturatesInsteadOfWrappingAtLongMax() throws Exception {
        var infinite = new InputStream() {
            @Override public int read() { return 0; }
            @Override public long skip(long requested) { return requested; }
        };
        var limited = new ProbeReader().limits(limits(Long.MAX_VALUE)).limitInput(infinite);
        assertEquals(Long.MAX_VALUE, limited.skip(Long.MAX_VALUE));
        var error = assertThrows(ReadLimitExceededException.class, limited::read);
        assertEquals(Long.MAX_VALUE, error.configured());
        assertEquals(Long.MAX_VALUE, error.actual());
    }

    @Test
    void postCopyGuardPreservesSizeExceptionAndRemovesOwnedTemporaryResources() throws Exception {
        try (var handler = new ProbeHandler(3)) {
            Path file = handler.getTempFile();
            Path parent = handler.getTempDir();
            Files.write(file, new byte[4]);
            var error = assertThrows(ReadLimitExceededException.class, handler::validateReadOptions);
            assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, error.limit());
            assertEquals(3, error.configured());
            assertEquals(4, error.actual());
            assertFalse(Files.exists(file));
            assertFalse(Files.exists(parent));
        }
    }

    @Test
    void oversizedCallerOwnedPathIsNotDeleted() throws Exception {
        Path external = Files.write(directory.resolve("input.csv"), new byte[4]);
        try (var handler = new ProbeHandler(3)) {
            handler.useExternalInput(external);
            var error = assertThrows(ReadLimitExceededException.class, handler::validateReadOptions);
            assertEquals(4, error.actual());
            assertArrayEquals(new byte[4], Files.readAllBytes(external));
        }
        assertTrue(Files.exists(external));
    }

    @Test
    void failedSizeInspectionRetainsItsCauseAndCleansTheOwnedDirectory() throws Exception {
        try (var handler = new ProbeHandler(3)) {
            Path file = handler.getTempFile();
            Path parent = handler.getTempDir();
            Files.delete(file);
            var error = assertThrows(ExcelKitException.class, handler::validateReadOptions);
            assertInstanceOf(java.nio.file.NoSuchFileException.class, error.getCause());
            assertFalse(Files.exists(parent));
        }
    }

    private static ReadLimits limits(long maximum) { return new ReadLimits(maximum, -1, -1, -1); }

    private static ReadSession<?> open(boolean excel, InputStream input, long maximum) {
        return excel ? ExcelReader.forMap().limits(limits(maximum)).open(input)
                     : CsvReader.forMap().limits(limits(maximum)).open(input);
    }

    private static final class ProbeReader extends AbstractReader<String, ProbeReader> {
        ProbeReader() { super((Function<RowData, String>) row -> "", null); }
    }

    private static final class ProbeHandler extends AbstractReadHandler<String> {
        ProbeHandler(long maximum) {
            super(InputStream.nullInputStream(), null, row -> "", null, ".csv", null,
                    options(maximum));
        }
        @Override public void read(Consumer<ReadResult<String>> consumer) { }
        @Override public void readWhile(Predicate<ReadResult<String>> predicate) { }
    }

    private static ReadOptions options(long maximum) {
        var configuration = new ReaderConfiguration();
        configuration.limits = limits(maximum);
        return configuration.snapshot();
    }

    private static final class TrackedInput extends ByteArrayInputStream {
        long consumed;
        boolean closed;
        TrackedInput(int length) { super(new byte[length]); }
        @Override public synchronized int read() {
            int value = super.read();
            if (value >= 0) consumed++;
            return value;
        }
        @Override public synchronized int read(byte[] bytes, int offset, int length) {
            int read = super.read(bytes, offset, length);
            if (read > 0) consumed += read;
            return read;
        }
        @Override public synchronized long skip(long length) {
            long skipped = super.skip(length);
            consumed += skipped;
            return skipped;
        }
        @Override public synchronized long transferTo(OutputStream output) throws IOException {
            // ByteArrayInputStream's optimized implementation bypasses read().
            long copied = super.transferTo(output);
            consumed += copied;
            return copied;
        }
        @Override public void close() { closed = true; }
    }
}
