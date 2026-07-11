package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/** Regression tests for map-mode reader behavior. */
class MapReaderStreamTest {

    record Item(String name, int value) {}

    private byte[] writeTestExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .write(Stream.of(new Item("A", 10), new Item("B", 20), new Item("C", 30)))
                .writeTo(out);
        return out.toByteArray();
    }

    @Test
    void read_withSparseRow_shouldFillGaps() throws IOException {
        // Write Excel with gap in data (col 0 has value, col 1 empty, col 2 has value)
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = wb.createSheet("Test");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("A");
            header.createCell(1).setCellValue("B");
            header.createCell(2).setCellValue("C");
            var data = sheet.createRow(1);
            data.createCell(0).setCellValue("x");
            // Skip col 1
            data.createCell(2).setCellValue("z");
            wb.write(out);
        }

        var results = new java.util.ArrayList<Map<String, String>>();
        ExcelReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("x", results.get(0).get("A"));
        assertEquals("z", results.get(0).get("C"));
    }


}
