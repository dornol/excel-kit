package io.github.dornol.excelkit.excel;

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

        @Test
        void closeDoesTheSameAsDiscard() throws IOException {
            long before = TempFiles.leftoverCount();
            CsvHandler handler = csvHandler();

            handler.close();

            assertEquals(before, TempFiles.leftoverCount());
            assertThrows(CsvWriteException.class, () -> handler.writeTo(OutputStream.nullOutputStream()));
        }

        @Test
        void tryWithResourcesAfterWritingDoesNothingExtra() throws IOException {
            long before = TempFiles.leftoverCount();
            try (CsvHandler handler = csvHandler()) {
                handler.writeTo(OutputStream.nullOutputStream());
            }
            assertEquals(before, TempFiles.leftoverCount());
        }
    }
}
