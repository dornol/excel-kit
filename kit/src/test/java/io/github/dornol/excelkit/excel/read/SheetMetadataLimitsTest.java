package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadLimitExceededException;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** Reading only the sheet names or headers of an untrusted upload honours the read limits. */
class SheetMetadataLimitsTest {

    private static byte[] workbook(int sheets) {
        var out = new ByteArrayOutputStream();
        try (var wb = ExcelWorkbook.create()) {
            for (int i = 0; i < sheets; i++) {
                wb.<String>sheet("S" + i).column("A", s -> s).column("B", s -> s).column("C", s -> s)
                        .write(Stream.of("x"));
            }
            wb.finish().writeTo(out);
        }
        return out.toByteArray();
    }

    @Test
    void withinTheLimitsTheResultIsTheSameAsWithoutThem() {
        byte[] bytes = workbook(2);

        var names = ExcelReader.getSheetNames(new ByteArrayInputStream(bytes), ReadLimits.UPLOAD);
        var headers = ExcelReader.getSheetHeaders(new ByteArrayInputStream(bytes), 1, 0, ReadLimits.UPLOAD);

        assertEquals(ExcelReader.getSheetNames(new ByteArrayInputStream(bytes)), names);
        assertEquals(List.of("A", "B", "C"), headers);
    }

    @Test
    void aStreamPastTheInputLimitIsRefusedWhileItIsCopied() {
        long[] served = {0};
        InputStream endless = new InputStream() {
            @Override public int read() { served[0]++; return 0; }
            @Override public int read(byte[] b, int off, int len) { served[0] += len; return len; }
        };
        var limits = new ReadLimits(1_000, -1, -1, -1);

        var names = assertThrows(ReadLimitExceededException.class, () -> ExcelReader.getSheetNames(endless, limits));
        assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, names.limit());
        assertThrows(ReadLimitExceededException.class, () -> ExcelReader.getSheetHeaders(endless, 0, 0, limits));
        assertTrue(served[0] < 100_000, "the copy stops at the limit: " + served[0]);
    }

    @Test
    void tooManySheetsOrColumnsAreRefused() {
        byte[] bytes = workbook(3);

        var sheets = assertThrows(ReadLimitExceededException.class, () ->
                ExcelReader.getSheetNames(new ByteArrayInputStream(bytes), new ReadLimits(-1, 2, -1, -1)));
        assertEquals(ReadLimitExceededException.Limit.SHEETS, sheets.limit());
        assertEquals(3, sheets.actual());

        var columns = assertThrows(ReadLimitExceededException.class, () ->
                ExcelReader.getSheetHeaders(new ByteArrayInputStream(bytes), 0, 0, new ReadLimits(-1, -1, 2, -1)));
        assertEquals(ReadLimitExceededException.Limit.COLUMNS, columns.limit());
    }

    @Test
    void theLimitsMustNotBeNull() {
        var in = new ByteArrayInputStream(new byte[0]);
        assertThrows(NullPointerException.class, () -> ExcelReader.getSheetNames(in, (ReadLimits) null));
        assertThrows(NullPointerException.class, () -> ExcelReader.getSheetHeaders(in, 0, 0, (ReadLimits) null));
    }
}
