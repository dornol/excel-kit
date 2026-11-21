package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The auto-width sample counts data rows, wherever the header sits. */
class AutoWidthSampleRowsTest {

    private static final String LONG = "a-very-very-very-very-long-value-here";

    private static int width(ExcelWriter<String> writer, List<String> rows) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.column("v", s -> s).write(rows).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return workbook.getSheetAt(0).getColumnWidth(0);
        }
    }

    @Test
    void rowsWrittenBeforeTheHeaderDoNotUseUpTheSample() throws Exception {
        int plain = width(ExcelWriter.<String>create().autoWidthSampleRows(10), List.of(LONG));
        int withPreamble = width(ExcelWriter.<String>create().autoWidthSampleRows(10)
                .beforeHeader(context -> 25), List.of(LONG));

        assertTrue(plain > 5000, "the long value must widen the column: " + plain);
        assertEquals(plain, withPreamble);
    }

    @Test
    void exactlyTheConfiguredNumberOfDataRowsIsSampled() throws Exception {
        int narrow = width(ExcelWriter.<String>create().autoWidthSampleRows(2), List.of("a", "b", LONG));
        int wide = width(ExcelWriter.<String>create().autoWidthSampleRows(3), List.of("a", "b", LONG));

        assertTrue(wide > narrow, "third row is sampled only when the sample is 3: " + narrow + " vs " + wide);
    }

    @Test
    void zeroDisablesSampling() throws Exception {
        int none = width(ExcelWriter.<String>create().autoWidthSampleRows(0), List.of(LONG));
        int sampled = width(ExcelWriter.<String>create().autoWidthSampleRows(5), List.of(LONG));

        assertTrue(sampled > none);
    }
}
