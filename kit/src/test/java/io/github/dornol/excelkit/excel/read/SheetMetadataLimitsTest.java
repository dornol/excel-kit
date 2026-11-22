package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadLimitExceededException;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** An inspector with limits refuses oversized uploads before it has copied or parsed them. */
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
        var limited = ExcelWorkbookInspector.create().limits(ReadLimits.UPLOAD);

        assertEquals(ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(bytes)),
                limited.sheetNames(new ByteArrayInputStream(bytes)));
        assertEquals(List.of("A", "B", "C"), limited.headers(new ByteArrayInputStream(bytes), 1, 0));
    }

    @Test
    void aStreamPastTheInputLimitIsRefusedWhileItIsCopied() {
        long[] served = {0};
        InputStream endless = new InputStream() {
            @Override public int read() { served[0]++; return 0; }
            @Override public int read(byte[] b, int off, int len) { served[0] += len; return len; }
        };
        var inspector = ExcelWorkbookInspector.create().limits(new ReadLimits(1_000, -1, -1, -1));

        var names = assertThrows(ReadLimitExceededException.class, () -> inspector.sheetNames(endless));
        assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, names.limit());
        assertThrows(ReadLimitExceededException.class, () -> inspector.headers(endless, 0, 0));
        assertTrue(served[0] < 100_000, "the copy stops at the limit: " + served[0]);
    }

    @Test
    void aFileLargerThanTheInputLimitIsRefusedWithoutBeingOpened(@TempDir Path dir) throws Exception {
        Path file = dir.resolve("big.xlsx");
        Files.write(file, workbook(1));
        var inspector = ExcelWorkbookInspector.create().limits(new ReadLimits(10, -1, -1, -1));

        var thrown = assertThrows(ReadLimitExceededException.class, () -> inspector.sheetNames(file));

        assertEquals(ReadLimitExceededException.Limit.INPUT_BYTES, thrown.limit());
        assertEquals(Files.size(file), thrown.actual());
        assertThrows(ReadLimitExceededException.class, () -> inspector.headers(file, 0, 0));
    }

    @Test
    void tooManySheetsOrColumnsAreRefused() {
        byte[] bytes = workbook(3);

        var sheets = assertThrows(ReadLimitExceededException.class, () -> ExcelWorkbookInspector.create()
                .limits(new ReadLimits(-1, 2, -1, -1)).sheetNames(new ByteArrayInputStream(bytes)));
        assertEquals(ReadLimitExceededException.Limit.SHEETS, sheets.limit());
        assertEquals(3, sheets.actual());

        var columns = assertThrows(ReadLimitExceededException.class, () -> ExcelWorkbookInspector.create()
                .limits(new ReadLimits(-1, -1, 2, -1)).headers(new ByteArrayInputStream(bytes), 0, 0));
        assertEquals(ReadLimitExceededException.Limit.COLUMNS, columns.limit());
    }

    @Test
    void aFileIsReadInPlaceAndNotCopied(@TempDir Path dir) throws Exception {
        Path file = dir.resolve("book.xlsx");
        Files.write(file, workbook(2));

        assertEquals(List.of("S0", "S1"), ExcelWorkbookInspector.create().sheetNames(file).stream()
                .map(ExcelSheetInfo::name).toList());
        try (var files = Files.list(dir)) {
            assertEquals(1, files.count());
        }
    }

    @Test
    void theSettingsRejectNull() {
        assertThrows(NullPointerException.class, () -> ExcelWorkbookInspector.create().limits(null));
        var inspector = ExcelWorkbookInspector.create();
        assertThrows(NullPointerException.class, () -> inspector.sheetNames((InputStream) null));
        assertThrows(NullPointerException.class, () -> inspector.headers((Path) null, 0, 0));
    }
}
