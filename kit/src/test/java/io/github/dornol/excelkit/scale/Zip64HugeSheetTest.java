package io.github.dornol.excelkit.scale;

import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.BufferedOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.LongStream;
import java.util.zip.CRC32;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A worksheet whose XML is larger than 4 GiB, the limit of a plain ZIP entry. The writer asks POI for
 * {@code Zip64Mode.AsNeeded} so that Excel, LibreOffice and streaming ZIP readers can read ordinary files;
 * this checks that the same setting still produces a valid archive when an entry outgrows 32-bit sizes. The archive is
 * read back with the JDK's {@link ZipFile}, which takes sizes from the central directory, and every byte of the big
 * entry is checked against the CRC the archive recorded.
 * <p>
 * It writes about 5.5 GB of XML and takes minutes, so it has its own tag and task: {@code ./gradlew :kit:hugeTest}.
 */
@Tag("huge")
class Zip64HugeSheetTest {

    private static final long DEFAULT_ROWS = 1_000_000L;
    /** Fewer rows give a quick smoke run of the checks below; only the default rows reach 4 GiB. */
    private static final long ROWS = Long.getLong("excelkit.huge.rows", DEFAULT_ROWS);
    private static final int COLUMNS = 110;
    private static final long FOUR_GIB = 4L * 1024 * 1024 * 1024;

    @TempDir
    Path dir;

    @Test
    void aSheetBeyondFourGibibytesIsWrittenAsAValidArchive() throws Exception {
        Path file = dir.resolve("huge.xlsx");

        long writeStart = System.nanoTime();
        ExcelWriter<Long> writer = ExcelWriter.create(options -> options.compressTempFiles(true));
        for (int c = 0; c < COLUMNS; c++) {
            final int column = c;
            writer.column("C" + c, id -> "r" + id + "c" + column);
        }
        try (OutputStream out = new BufferedOutputStream(Files.newOutputStream(file), 1 << 16)) {
            writer.write(LongStream.range(0, ROWS).boxed()).writeTo(out);
        }
        System.out.printf("wrote %,d bytes in %d s%n", Files.size(file), (System.nanoTime() - writeStart) / 1_000_000_000L);

        try (ZipFile zip = new ZipFile(file.toFile())) {
            ZipEntry sheet = zip.getEntry("xl/worksheets/sheet1.xml");
            assertNotNull(sheet, "the worksheet entry");
            if (ROWS >= DEFAULT_ROWS) {
                assertTrue(sheet.getSize() > FOUR_GIB, "the test must produce an entry above 4 GiB, got " + sheet.getSize());
            }
            System.out.printf("worksheet entry: %,d bytes uncompressed, %,d compressed%n", sheet.getSize(), sheet.getCompressedSize());

            CRC32 crc = new CRC32();
            long bytes = 0;
            long rows = 0;
            byte[] tail = new byte[0];
            byte[] pattern = "</row>".getBytes(StandardCharsets.US_ASCII);
            byte[] window = new byte[(1 << 20) + pattern.length];
            int carry = 0;
            long readStart = System.nanoTime();
            try (InputStream in = zip.getInputStream(sheet)) {
                int n;
                while ((n = in.read(window, carry, window.length - carry)) > 0) {
                    crc.update(window, carry, n);
                    bytes += n;
                    int total = carry + n;
                    for (int i = 0; i + pattern.length <= total; i++) {
                        if (i + pattern.length <= carry) continue; // ended in the bytes carried over: counted already
                        if (window[i] == '<' && window[i + 1] == '/' && window[i + 2] == 'r' && window[i + 3] == 'o'
                                && window[i + 4] == 'w' && window[i + 5] == '>') rows++;
                    }
                    tail = concatTail(tail, window, total, 4096);
                    carry = Math.min(pattern.length - 1, total);
                    System.arraycopy(window, total - carry, window, 0, carry);
                }
            }
            System.out.printf("read %,d bytes back in %d s%n", bytes, (System.nanoTime() - readStart) / 1_000_000_000L);

            assertEquals(sheet.getSize(), bytes, "every byte of the entry is readable");
            assertEquals(sheet.getCrc(), crc.getValue(), "the contents match the CRC the archive recorded");
            assertEquals(ROWS + 1, rows, "the header and every data row");
            String end = new String(tail, StandardCharsets.UTF_8);
            assertTrue(end.contains(">r" + (ROWS - 1) + "c" + (COLUMNS - 1) + "<"), "the last cell of the last row is there");
        }
    }

    /** The last {@code keep} bytes of what has been read so far. */
    private static byte[] concatTail(byte[] previous, byte[] window, int total, int keep) {
        int fromWindow = Math.min(keep, total);
        int fromPrevious = Math.min(previous.length, keep - fromWindow);
        byte[] out = new byte[fromPrevious + fromWindow];
        System.arraycopy(previous, previous.length - fromPrevious, out, 0, fromPrevious);
        System.arraycopy(window, total - fromWindow, out, fromPrevious, fromWindow);
        return out;
    }
}
