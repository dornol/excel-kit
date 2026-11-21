package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** A table that cannot span rollover sheets fails when the first rollover is due, not after the whole stream. */
class TableRolloverTest {

    private static final TableOptions SINGLE_SHEET = new TableOptions("Data", "TableStyleMedium2", true, false);
    private static final int MAX_ROWS = 10;

    private static IntStream countedRows(int rows, AtomicInteger consumed) {
        return IntStream.range(0, rows).peek(i -> consumed.incrementAndGet());
    }

    @Test
    void excelWriterFailsAtTheFirstRolloverNotAtTheEnd() {
        AtomicInteger consumed = new AtomicInteger();
        ExcelWriteException error = assertThrows(ExcelWriteException.class, () ->
                ExcelWriter.<Integer>create().maxRows(MAX_ROWS).table(SINGLE_SHEET).column("v", i -> i)
                        .write(countedRows(1_000, consumed).boxed()).writeTo(OutputStream.nullOutputStream()));

        assertTrue(error.getMessage().contains("perRolloverSheet"), error.getMessage());
        assertTrue(error.getMessage().contains("'Data'"), error.getMessage());
        assertEquals(MAX_ROWS + 1, consumed.get(), "stopped when the row that needs a second sheet arrived");
    }

    @Test
    void sheetWriterFailsAtTheFirstRolloverNotAtTheEnd() {
        AtomicInteger consumed = new AtomicInteger();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            ExcelWriteException error = assertThrows(ExcelWriteException.class, () ->
                    workbook.<Integer>sheet("S").maxRows(MAX_ROWS).table(SINGLE_SHEET).column("v", i -> i)
                            .write(countedRows(1_000, consumed).boxed()));

            assertTrue(error.getMessage().contains("perRolloverSheet"), error.getMessage());
            assertEquals(MAX_ROWS + 1, consumed.get());
        }
    }

    @Test
    void aTableThatFitsOnOneSheetIsFine() {
        assertDoesNotThrow(() ->
                ExcelWriter.<Integer>create().maxRows(MAX_ROWS).table(SINGLE_SHEET).column("v", i -> i)
                        .write(IntStream.range(0, MAX_ROWS).boxed()).writeTo(OutputStream.nullOutputStream()));
    }

    @Test
    void perRolloverSheetTablesRollOver() {
        assertDoesNotThrow(() ->
                ExcelWriter.<Integer>create().maxRows(MAX_ROWS).table("Data").column("v", i -> i)
                        .write(IntStream.range(0, MAX_ROWS * 3).boxed()).writeTo(OutputStream.nullOutputStream()));
    }
}
