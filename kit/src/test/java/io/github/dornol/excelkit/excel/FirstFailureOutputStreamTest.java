package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * {@code ExcelHandler.FirstFailureOutputStream} wraps the stream a workbook is written to and remembers the first
 * {@link IOException} that stream throws, so that the failure of the stream (a client that went away, a full disk)
 * can be reported instead of what POI's ZIP writer throws later while it cleans up.
 */
class FirstFailureOutputStreamTest {

    /** Fails on the chosen operation, with an exception of its own for every call, and records what reached it. */
    private static final class Recording extends OutputStream {
        final ByteArrayOutputStream received = new ByteArrayOutputStream();
        final List<IOException> thrown = new ArrayList<>();
        int flushes;
        int closes;
        boolean failWrite;
        boolean failFlush;
        boolean failClose;

        private IOException failure(String what) {
            IOException e = new IOException(what + " #" + (thrown.size() + 1));
            thrown.add(e);
            return e;
        }

        @Override
        public void write(int b) throws IOException {
            if (failWrite) throw failure("write(int)");
            received.write(b);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            if (failWrite) throw failure("write(bytes)");
            received.write(b, off, len);
        }

        @Override
        public void flush() throws IOException {
            flushes++;
            if (failFlush) throw failure("flush");
        }

        @Override
        public void close() throws IOException {
            closes++;
            if (failClose) throw failure("close");
        }
    }

    @Nested
    class PassingThrough {
        @Test
        void everyByteReachesTheWrappedStreamUnchanged() throws Exception {
            Recording target = new Recording();
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            out.write(7);
            out.write(new byte[]{1, 2, 3});
            out.write(new byte[]{9, 8, 7, 6, 5}, 1, 3);

            assertArrayEquals(new byte[]{7, 1, 2, 3, 8, 7, 6}, target.received.toByteArray());
        }

        @Test
        void aLargeBufferIsWrittenInOneCallNotByteByByte() throws Exception {
            // FilterOutputStream would write byte by byte; a workbook is megabytes
            int[] calls = {0};
            OutputStream counting = new OutputStream() {
                @Override
                public void write(int b) {
                    calls[0]++;
                }

                @Override
                public void write(byte[] b, int off, int len) {
                    calls[0]++;
                }
            };
            new ExcelHandler.FirstFailureOutputStream(counting).write(new byte[100_000]);

            assertEquals(1, calls[0]);
        }

        @Test
        void flushAndCloseAreForwardedOnceEach() throws Exception {
            Recording target = new Recording();
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            out.flush();
            out.close();

            assertEquals(1, target.flushes);
            assertEquals(1, target.closes);
        }

        @Test
        void withoutAFailureTheReportedExceptionIsReturnedAsItIs() throws Exception {
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(new Recording());
            out.write(1);
            IOException reported = new IOException("from the zip writer");

            assertSame(reported, out.reportedAs(reported));
            assertEquals(0, reported.getSuppressed().length);
        }
    }

    @Nested
    class RememberingTheFirstFailure {
        @Test
        void aFailingWriteOfOneByteIsRethrownAndRemembered() {
            Recording target = new Recording();
            target.failWrite = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            IOException thrown = assertThrows(IOException.class, () -> out.write(1));

            assertSame(target.thrown.get(0), thrown);
            IOException later = new IOException("later");
            assertSame(thrown, out.reportedAs(later));
        }

        @Test
        void aFailingWriteOfABufferIsRethrownAndRemembered() {
            Recording target = new Recording();
            target.failWrite = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            IOException thrown = assertThrows(IOException.class, () -> out.write(new byte[10], 2, 5));

            assertSame(target.thrown.get(0), thrown);
            assertSame(thrown, out.reportedAs(new IOException("later")));
        }

        @Test
        void aFailingFlushIsRethrownAndRemembered() {
            Recording target = new Recording();
            target.failFlush = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            IOException thrown = assertThrows(IOException.class, out::flush);

            assertSame(target.thrown.get(0), thrown);
            assertSame(thrown, out.reportedAs(new IOException("later")));
        }

        @Test
        void aFailingCloseIsRethrownAndRemembered() {
            Recording target = new Recording();
            target.failClose = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            IOException thrown = assertThrows(IOException.class, out::close);

            assertSame(target.thrown.get(0), thrown);
            assertSame(thrown, out.reportedAs(new IOException("later")));
        }

        @Test
        void theFirstFailureWinsOverEveryLaterOne() {
            Recording target = new Recording();
            target.failWrite = true;
            target.failFlush = true;
            target.failClose = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);

            assertThrows(IOException.class, () -> out.write(1));
            assertThrows(IOException.class, out::flush);
            assertThrows(IOException.class, () -> out.write(new byte[3], 0, 3));
            assertThrows(IOException.class, out::close);

            assertEquals(4, target.thrown.size());
            assertSame(target.thrown.get(0), out.reportedAs(new IOException("later")));
        }

        @Test
        void theFailureThatWasReportedIsAttachedToTheStreamsOwn() {
            Recording target = new Recording();
            target.failWrite = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);
            assertThrows(IOException.class, () -> out.write(1));
            IOException zipError = new IOException("This archive contains unclosed entries.");

            IOException result = out.reportedAs(zipError);

            assertSame(target.thrown.get(0), result);
            assertEquals(List.of(zipError), List.of(result.getSuppressed()), "what POI reported stays visible");
        }

        @Test
        void aStreamsOwnFailureThatIsAlsoTheReportedOneIsNotAttachedToItself() {
            Recording target = new Recording();
            target.failWrite = true;
            ExcelHandler.FirstFailureOutputStream out = new ExcelHandler.FirstFailureOutputStream(target);
            IOException thrown = assertThrows(IOException.class, () -> out.write(1));

            IOException result = out.reportedAs(thrown);

            assertSame(thrown, result);
            assertEquals(0, result.getSuppressed().length, "addSuppressed(self) would throw");
        }
    }

    @Nested
    class ThroughTheHandler {
        private record Item(String text) {
        }

        private static ExcelHandler handler() {
            return ExcelWriter.<Item>create().column("Text", Item::text)
                    .write(IntStream.range(0, 3000).mapToObj(i -> new Item("value " + i)).toList());
        }

        @Test
        void aStreamThatFailsOnFlushIsTheCauseOfTheWriteFailure() {
            Recording target = new Recording();
            target.failFlush = true;

            ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> handler().writeTo(target));

            assertTrue(!target.thrown.isEmpty(), "the stream was flushed");
            assertSame(target.thrown.get(0), e.getCause(), String.valueOf(e.getCause()));
        }

        @Test
        void theWriteLeavesTheCallersStreamOpenSoItsCloseNeverFailsTheWrite() throws Exception {
            // The caller owns the stream: the handler writes the workbook, flushes at most, and never closes it.
            Recording target = new Recording();
            target.failClose = true;

            handler().writeTo(target);

            assertEquals(0, target.closes, "the stream was closed by the write");
            assertTrue(target.received.size() > 0, "the workbook arrived");
            assertEquals(0, target.thrown.size());
        }
    }
}
