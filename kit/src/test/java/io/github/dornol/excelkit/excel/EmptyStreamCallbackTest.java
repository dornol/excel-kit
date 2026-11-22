package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class EmptyStreamCallbackTest {

    // ========================================================================
    // Empty stream + callbacks
    // ========================================================================
    @Test
    void emptyStream_withAfterData_shouldStillCallback() throws IOException {
        boolean[] called = {false};
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .afterData(ctx -> {
                    called[0] = true;
                    return ctx.getCurrentRow();
                })
                .write(Stream.empty())
                .writeTo(out);

        assertTrue(called[0], "afterData should be called even with empty stream");
    }

    @Test
    void emptyStream_withBeforeHeader_shouldStillCallback() throws IOException {
        boolean[] called = {false};
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .beforeHeader(ctx -> {
                    called[0] = true;
                    return ctx.getCurrentRow();
                })
                .write(Stream.empty())
                .writeTo(out);

        assertTrue(called[0], "beforeHeader should be called even with empty stream");
    }
}
