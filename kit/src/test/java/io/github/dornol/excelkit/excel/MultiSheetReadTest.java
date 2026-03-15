package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class MultiSheetReadTest {

    @Test
    void getSheetNames_shouldReturnAllSheetNames() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Users")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"));
            wb.<String>sheet("Orders")
                    .column("Item", s -> s)
                    .write(Stream.of("Widget"));
            wb.finish().consumeOutputStream(out);
        }

        List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(
                new ByteArrayInputStream(out.toByteArray()));

        assertEquals(2, sheets.size());
        assertEquals("Users", sheets.get(0).name());
        assertEquals(0, sheets.get(0).index());
        assertEquals("Orders", sheets.get(1).name());
        assertEquals(1, sheets.get(1).index());
    }

    @Test
    void getSheetHeaders_shouldReturnHeaderNames() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("First", s -> s)
                .addColumn("Second", s -> s)
                .addColumn("Third", s -> s)
                .write(Stream.of("data"))
                .consumeOutputStream(out);

        List<String> headers = ExcelReader.getSheetHeaders(
                new ByteArrayInputStream(out.toByteArray()), 0, 0);

        assertEquals(List.of("First", "Second", "Third"), headers);
    }

    @Test
    void readSpecificSheet_shouldReadCorrectSheet() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Users")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"));
            wb.<String>sheet("Orders")
                    .column("Item", s -> s)
                    .write(Stream.of("Widget", "Gadget"));
            wb.finish().consumeOutputStream(out);
        }

        // Read second sheet
        var results = new java.util.ArrayList<String>();
        new ExcelReader<>(Holder::new, null)
                .sheetIndex(1)
                .addColumn((h, c) -> h.value = c.asString())
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(r -> results.add(r.data().value));

        assertEquals(List.of("Widget", "Gadget"), results);
    }

    @Test
    void getSheetNames_singleSheet() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .sheetName("Data")
                .addColumn("Name", s -> s)
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(
                new ByteArrayInputStream(out.toByteArray()));

        assertEquals(1, sheets.size());
        assertEquals("Data", sheets.get(0).name());
    }

    @Test
    void getSheetHeaders_withHeaderRowIndex() throws IOException {
        // Create a file with content before the header
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .beforeHeader(ctx -> {
                    var row = ctx.getSheet().createRow(0);
                    row.createCell(0).setCellValue("Title Row");
                    return 1;
                })
                .addColumn("Name", s -> s)
                .addColumn("Age", s -> "30")
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        List<String> headers = ExcelReader.getSheetHeaders(
                new ByteArrayInputStream(out.toByteArray()), 0, 1);

        assertEquals(List.of("Name", "Age"), headers);
    }

    public static class Holder {
        String value;
    }
}
