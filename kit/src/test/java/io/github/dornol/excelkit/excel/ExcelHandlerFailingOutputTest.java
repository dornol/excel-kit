package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.OutputStream;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * When the stream a workbook is written to fails (a client that disconnected, a full disk), the
 * {@link ExcelWriteException} names that failure as its cause. The ZIP writer underneath reports a different,
 * less telling error while it cleans up ("This archive contains unclosed entries."); that must not hide it.
 */
class ExcelHandlerFailingOutputTest {

    private record Item(String text) {
    }

    private static final class FailingOutput extends OutputStream {
        final IOException failure = new IOException("Broken pipe");
        private final int allowed;
        private int written;

        FailingOutput(int allowed) {
            this.allowed = allowed;
        }

        @Override
        public void write(int b) throws IOException {
            write(new byte[]{(byte) b}, 0, 1);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            if (written + len > allowed) throw failure;
            written += len;
        }
    }

    private static ExcelWriter<Item> writer() {
        return ExcelWriter.<Item>create().column("Text", Item::text);
    }

    private static ExcelHandler handler(ExcelWriter<Item> writer) {
        return writer.write(IntStream.range(0, 5000).mapToObj(i -> new Item("value " + i)).toList());
    }

    @Test
    void theStreamsOwnFailureIsTheCauseWhereverTheWriteBreaks() {
        for (int allowed : new int[]{0, 100, 500, 2000, 6000, 20_000}) {
            FailingOutput out = new FailingOutput(allowed);

            ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> handler(writer()).writeTo(out),
                    "allowed=" + allowed);

            assertSame(out.failure, e.getCause(), "allowed=" + allowed + ", got " + e.getCause());
        }
    }

    @Test
    void theSameHoldsForAnEncryptedWorkbookWrittenToAFailingStream() {
        FailingOutput out = new FailingOutput(1000);

        ExcelWriteException e = assertThrows(ExcelWriteException.class,
                () -> handler(writer().password("pw")).writeTo(out));

        assertSame(out.failure, e.getCause(), String.valueOf(e.getCause()));
    }

    @Test
    void aPasswordGivenWhenTheResponseIsWrittenBehavesTheSame() {
        FailingOutput out = new FailingOutput(1000);

        ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> handler(writer()).writeTo(out, "pw"));

        assertSame(out.failure, e.getCause(), String.valueOf(e.getCause()));
    }

    @Test
    void aWorkbookThatSucceedsIsUnaffected() throws Exception {
        java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
        handler(writer()).writeTo(out);
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new java.io.ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(5000, wb.getSheetAt(0).getLastRowNum());
        }
    }
}
