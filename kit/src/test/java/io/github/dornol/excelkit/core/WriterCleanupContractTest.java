package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelWriteException;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import java.util.Spliterators;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** Verifies writer exception boundaries and caller-owned stream lifecycle. */
class WriterCleanupContractTest {

    @Test
    void iteratorFailureIsWrappedAndInputStreamRemainsCallerOwned() {
        for (boolean excel : new boolean[]{false, true}) {
            AtomicBoolean closed = new AtomicBoolean();
            Stream<String> rows = failingRows(closed);

            if (excel) {
                assertThrows(ExcelWriteException.class, () -> ExcelWriter.<String>create()
                        .column("Name", value -> value).write(rows));
            } else {
                assertThrows(CsvWriteException.class, () -> CsvWriter.<String>create()
                        .column("Name", value -> value).write(rows));
            }
            assertFalse(closed.get(), format(excel));
            rows.close();
        }
    }

    @Test
    void outputFailureIsWrappedForBothFormats() {
        for (boolean excel : new boolean[]{false, true}) {
            OutputStream output = new FailingOutputStream();
            if (excel) {
                assertThrows(ExcelWriteException.class, () -> ExcelWriter.<String>create()
                        .column("Name", value -> value).write(java.util.List.of("A")).writeTo(output));
            } else {
                assertThrows(CsvWriteException.class, () -> CsvWriter.<String>create()
                        .column("Name", value -> value).write(java.util.List.of("A")).writeTo(output));
            }
        }
    }

    private static Stream<String> failingRows(AtomicBoolean closed) {
        Iterator<String> iterator = new Iterator<>() {
            private int index;

            @Override
            public boolean hasNext() {
                return index < 2;
            }

            @Override
            public String next() {
                if (index++ == 0) return "A";
                throw new IllegalStateException("iterator failure");
            }
        };
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(iterator, 0), false)
                .onClose(() -> closed.set(true));
    }

    private static String format(boolean excel) {
        return excel ? "Excel" : "CSV";
    }

    private static final class FailingOutputStream extends OutputStream {
        @Override
        public void write(int value) throws IOException {
            throw new IOException("output failure");
        }
    }
}
