package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelWorkbookInspector;
import io.github.dornol.excelkit.excel.read.ExcelReadException;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.read.ExcelSheetInfo;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookMultiSheetReadTest {
    @Test
    void getSheetNames_rolloverSheets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().maxRows(3)
                .sheetName("Data")
                .column("Name", s -> s)
                .write(Stream.of("A", "B", "C", "D", "E", "F", "G", "H", "I"))
                .writeTo(out);

        var sheets = ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(out.toByteArray()));
        assertTrue(sheets.size() >= 3, "Should have multiple sheets");
        assertEquals("Data", sheets.get(0).name());
    }

    @Test
    void getSheetHeaders_emptySheet() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s)
                .write(Stream.empty())
                .writeTo(out);

        var headers = ExcelWorkbookInspector.create().headers(
                new ByteArrayInputStream(out.toByteArray()), 0, 0);
        assertEquals(List.of("A", "B"), headers);
    }

    @Test
    void getSheetNames_exception_invalidStream() {
        assertThrows(ExcelReadException.class, () ->
                ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(new byte[]{1, 2, 3})));
    }

    @Test
    void getSheetHeaders_exception_invalidStream() {
        assertThrows(ExcelReadException.class, () ->
                ExcelWorkbookInspector.create().headers(new ByteArrayInputStream(new byte[]{1, 2, 3}), 0, 0));
    }

    @Test
    void readAllSheets_iterateByName() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<String>sheet("Users")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice", "Bob"));
            wb.<String>sheet("Items")
                    .column("Item", s -> s)
                    .write(Stream.of("Widget"));
            wb.finish().writeTo(out);
        }

        byte[] data = out.toByteArray();
        var sheets = ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(data));

        // Read each sheet by name
        Map<String, List<String>> allData = new LinkedHashMap<>();
        for (ExcelSheetInfo info : sheets) {
            List<String> values = new ArrayList<>();
            new ExcelReader<>(Holder::new, null)
                    .sheetIndex(info.index())
                    .column((h, c) -> h.value = c.asString())
                    .read(new ByteArrayInputStream(data), r -> values.add(r.data().value));
            allData.put(info.name(), values);
        }

        assertEquals(List.of("Alice", "Bob"), allData.get("Users"));
        assertEquals(List.of("Widget"), allData.get("Items"));
    }

    public static class Holder {
        String value;
    }
}
