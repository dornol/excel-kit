package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.testsupport.TempFiles;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** A handler that will not be written is released with {@code discard()}. */
class HandlerDiscardTest {

    private static ExcelHandler excelHandler() {
        return ExcelWriter.<String>create().column("A", s -> s).write(List.of("x"));
    }

    private static CsvHandler csvHandler() {
        return CsvWriter.<String>create().column("A", s -> s).write(List.of("x"));
    }

    @Nested
    class Excel {
        @Test
        void discardReleasesTheWorkbook() throws IOException {
            long before = TempFiles.leftoverCount();
            ExcelHandler handler = excelHandler();
            assertEquals(before + 1, TempFiles.leftoverCount());

            handler.discard();

            assertEquals(before, TempFiles.leftoverCount());
        }

        @Test
        void discardIsIdempotent() {
            ExcelHandler handler = excelHandler();
            handler.discard();
            assertDoesNotThrow(handler::discard);
        }

        @Test
        void aDiscardedHandlerCannotBeWritten() {
            ExcelHandler handler = excelHandler();
            handler.discard();
            assertThrows(ExcelWriteException.class, () -> handler.writeTo(OutputStream.nullOutputStream()));
        }

        @Test
        void discardAfterWritingDoesNothing() throws IOException {
            long before = TempFiles.leftoverCount();
            ExcelHandler handler = excelHandler();
            handler.writeTo(OutputStream.nullOutputStream());

            assertDoesNotThrow(handler::discard);

            assertEquals(before, TempFiles.leftoverCount());
        }

        @Test
        void aHandlerKeptAfterAPasswordConflictCanStillBeDiscarded() throws IOException {
            long before = TempFiles.leftoverCount();
            ExcelHandler handler = ExcelWriter.<String>create().password("first").column("A", s -> s)
                    .write(List.of("x"));
            assertThrows(IllegalStateException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), "second"));
            assertEquals(before + 1, TempFiles.leftoverCount(), "a conflict leaves the handler usable");

            handler.discard();

            assertEquals(before, TempFiles.leftoverCount());
        }
    }

    @Nested
    class Csv {
        @Test
        void discardDeletesTheStagedFile() throws IOException {
            long before = TempFiles.leftoverCount();
            CsvHandler handler = csvHandler();
            assertEquals(before + 1, TempFiles.leftoverCount());

            handler.discard();

            assertEquals(before, TempFiles.leftoverCount());
        }

        @Test
        void discardIsIdempotent() {
            CsvHandler handler = csvHandler();
            handler.discard();
            assertDoesNotThrow(handler::discard);
        }

        @Test
        void aDiscardedHandlerCannotBeWritten() {
            CsvHandler handler = csvHandler();
            handler.discard();
            assertThrows(CsvWriteException.class, () -> handler.writeTo(OutputStream.nullOutputStream()));
        }
    }

    @Nested
    class OneContract {
        /**
         * A handler is written after the method that made it has returned (a StreamingResponseBody), so
         * try-with-resources would release it before that: neither handler is AutoCloseable.
         */
        @Test
        void neitherHandlerIsAutoCloseable() {
            assertFalse(AutoCloseable.class.isAssignableFrom(ExcelHandler.class));
            assertFalse(AutoCloseable.class.isAssignableFrom(CsvHandler.class));
            assertFalse(AutoCloseable.class.isAssignableFrom(FileHandler.class));
        }

        @Test
        void bothAreReleasedThroughTheInterface() throws IOException {
            for (FileHandler handler : List.of(excelHandler(), csvHandler())) {
                long before = TempFiles.leftoverCount();
                handler.discard();
                assertEquals(before - 1, TempFiles.leftoverCount(), handler.getClass().getSimpleName());
                assertDoesNotThrow(handler::discard);
                assertThrows(io.github.dornol.excelkit.core.ExcelKitException.class,
                        () -> handler.writeTo(OutputStream.nullOutputStream()), handler.getClass().getSimpleName());
            }
        }

        @Test
        void aWrittenHandlerIgnoresDiscardAfterwards() throws IOException {
            for (FileHandler handler : List.of(excelHandler(), csvHandler())) {
                handler.writeTo(OutputStream.nullOutputStream());
                long before = TempFiles.leftoverCount();
                assertDoesNotThrow(handler::discard);
                assertEquals(before, TempFiles.leftoverCount());
            }
        }

        @Test
        void aHandlerWrittenLaterIsNotReleasedByTheMethodThatMadeIt() {
            // the pattern that try-with-resources would break: the handler is made here and written afterwards
            java.util.function.Supplier<java.util.function.Consumer<OutputStream>> controller = () -> excelHandler()::writeTo;
            var body = controller.get();
            assertDoesNotThrow(() -> body.accept(OutputStream.nullOutputStream()));
        }
    }
}
