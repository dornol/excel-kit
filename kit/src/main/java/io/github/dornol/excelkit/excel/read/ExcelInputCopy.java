package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadLimitExceededException;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/** Copies an input stream to a temp file, stopping as soon as it passes the byte limit. */
final class ExcelInputCopy {
    private ExcelInputCopy() {
    }

    /**
     * @param limit the most bytes to accept, or a negative number for no limit
     * @throws ReadLimitExceededException if the stream holds more than {@code limit} bytes; the copy stops
     *                                    at the first byte beyond it
     */
    static void copy(InputStream input, Path target, long limit) throws IOException {
        try (OutputStream out = Files.newOutputStream(target)) {
            byte[] buffer = new byte[8192];
            long total = 0;
            int read;
            while ((read = input.read(buffer)) >= 0) {
                if (limit >= 0 && read > limit - total) {
                    long observed = total > Long.MAX_VALUE - read ? Long.MAX_VALUE : total + read;
                    throw new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES, limit, observed);
                }
                out.write(buffer, 0, read);
                total += read;
            }
        }
    }
}
