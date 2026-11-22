package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** {@link ReadLimits#maxColumns()} stops a file with more columns than allowed, and says how many it has. */
class ColumnLimitTest {

    private static byte[] csv(int columns) {
        String header = String.join(",", java.util.stream.IntStream.range(0, columns).mapToObj(i -> "h" + i).toList());
        String row = String.join(",", java.util.Collections.nCopies(columns, "v"));
        return (header + "\n" + row + "\n").getBytes(StandardCharsets.UTF_8);
    }

    private static byte[] xlsx(int columns) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("S");
            for (int r = 0; r < 2; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < columns; c++) row.createCell(c).setCellValue(r == 0 ? "h" + c : "v");
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static ReadLimits limit(int maxColumns) {
        return new ReadLimits(-1, -1, maxColumns, -1, -1);
    }

    @Test
    void csvWithTooManyColumnsIsRejectedWithTheCount() {
        ReadLimitExceededException error = assertThrows(ReadLimitExceededException.class,
                () -> CsvReader.forMap().limits(limit(2)).read(new ByteArrayInputStream(csv(3)), r -> { }));
        assertEquals(ReadLimitExceededException.Limit.COLUMNS, error.limit());
        assertEquals(2, error.configured());
        assertEquals(3, error.actual());
    }

    @Test
    void xlsxWithTooManyColumnsIsRejectedWithTheCount() throws Exception {
        byte[] file = xlsx(5);
        ReadLimitExceededException error = assertThrows(ReadLimitExceededException.class,
                () -> ExcelReader.forMap().limits(limit(4)).read(new ByteArrayInputStream(file), r -> { }));
        assertEquals(ReadLimitExceededException.Limit.COLUMNS, error.limit());
        assertEquals(4, error.configured());
        assertEquals(5, error.actual());
    }

    @Test
    void exactlyTheAllowedNumberOfColumnsIsFine() throws Exception {
        List<Object> rows = new ArrayList<>();
        assertDoesNotThrow(() -> CsvReader.forMap().limits(limit(3)).read(new ByteArrayInputStream(csv(3)), rows::add));
        assertDoesNotThrow(() -> ExcelReader.forMap().limits(limit(3)).read(new ByteArrayInputStream(xlsx(3)), rows::add));
        assertEquals(2, rows.size());
    }
}
