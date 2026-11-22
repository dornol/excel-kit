package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** Sheet rollover at small {@code maxRows} values: every sheet holds at most maxRows data rows. */
class MaxRowsBoundaryTest {

    /** Data rows (below the one header row) of every sheet. */
    private static List<Integer> dataRowsPerSheet(int maxRows, int rows) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Integer>create().column("N", n -> n).maxRows(maxRows)
                .write(IntStream.rangeClosed(1, rows).boxed().toList()).writeTo(out);
        List<Integer> counts = new ArrayList<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            for (org.apache.poi.ss.usermodel.Sheet sheet : wb) counts.add(sheet.getLastRowNum());
        }
        return counts;
    }

    @Test
    void maxRowsOne_rollsOverEveryRow() throws Exception {
        assertEquals(List.of(1, 1, 1), dataRowsPerSheet(1, 3));
    }

    @Test
    void maxRowsTwo_rollsOverEveryTwoRows() throws Exception {
        assertEquals(List.of(2, 2, 1), dataRowsPerSheet(2, 5));
    }

    @Test
    void exactMultiple_doesNotCreateAnEmptyTrailingSheet() throws Exception {
        assertEquals(List.of(2, 2), dataRowsPerSheet(2, 4));
    }
}
