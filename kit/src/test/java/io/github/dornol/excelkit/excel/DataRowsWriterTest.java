package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.*;

/**
 * The settings of how rows are written live in one base class ({@link DataRowsWriter}), so a sheet and a list written
 * into a template take and check them the same way.
 */
class DataRowsWriterTest {
    record Row(int n) {}

    /** The same settings, written once for both kinds of writer. */
    private static <W extends DataRowsWriter<Row, W>> W settings(W w) {
        return w.column("N", Row::n, c -> c.type(ExcelDataType.INTEGER))
                .rowHeight(30)
                .rowColor(r -> r.n() % 2 == 0 ? ExcelColor.LIGHT_BLUE : null)
                .summary(s -> s.label("Total").sum("N"))
                .autoWidthSampleRows(10)
                .writeErrorPolicy(WriteErrorPolicy.LENIENT);
    }

    @Test
    void theSameSettingsWriteTheSameRowsOnASheetAndOnATemplateList() throws IOException {
        var rows = List.of(new Row(1), new Row(2), new Row(3));
        var sheetOut = new ByteArrayOutputStream();
        settings(ExcelWriter.<Row>create()).write(rows).writeTo(sheetOut);

        var templateOut = new ByteArrayOutputStream();
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            wb.createSheet("S");
            wb.write(out);
            try (var template = new ExcelTemplateWriter(new ByteArrayInputStream(out.toByteArray()))) {
                settings(template.<Row>list(0)).writeWithHeaders(rows);
                template.finish().writeTo(templateOut);
            }
        }
        try (var a = new XSSFWorkbook(new ByteArrayInputStream(sheetOut.toByteArray()));
             var b = new XSSFWorkbook(new ByteArrayInputStream(templateOut.toByteArray()))) {
            assertEquals(dump(a), dump(b));
            assertEquals(30f, b.getSheetAt(0).getRow(1).getHeightInPoints(), 0.01);
        }
    }

    private static String dump(XSSFWorkbook wb) {
        var sheet = wb.getSheetAt(0);
        var fmt = new org.apache.poi.ss.usermodel.DataFormatter(java.util.Locale.ROOT);
        var sb = new StringBuilder();
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            var row = sheet.getRow(r);
            sb.append(r).append(':').append(row.getHeightInPoints()).append(':')
                    .append(java.util.stream.IntStream.range(0, row.getLastCellNum())
                            .mapToObj(c -> fmt.formatCellValue(row.getCell(c))).collect(Collectors.joining(","))).append('\n');
        }
        return sb.toString();
    }

    @Test
    void aTemplateListChecksItsSettingsAsASheetDoes() throws IOException {
        try (var wb = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            wb.createSheet("S");
            wb.write(out);
            try (var template = new ExcelTemplateWriter(new ByteArrayInputStream(out.toByteArray()))) {
                var list = template.<Row>list(0);
                assertThrows(IllegalArgumentException.class, () -> list.rowHeight(0));
                assertThrows(IllegalArgumentException.class, () -> list.rowHeight(-3));
                assertThrows(IllegalArgumentException.class, () -> list.onProgress(0, (total, cursor) -> {}));
                assertThrows(IllegalArgumentException.class, () -> list.autoWidthSampleRows(-1));
                assertThrows(NullPointerException.class, () -> list.afterData(null));
                assertThrows(NullPointerException.class, () -> list.table((TableOptions) null));
            }
        }
    }
}
