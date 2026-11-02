package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;

class UploadResultTest {

    @Test
    void read_collectsRowsAndStructuredErrors() {
        String csv = "Name,Price\nNotebook,1200\nPen,not-a-number\n";
        var reader = CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Price", (p, cell) -> p.price = cell.asInt());

        UploadResult<Product> result = UploadResult.read("CSV", consumer ->
                reader.read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), consumer));

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, result.errorCount());
        assertEquals("Notebook", result.rows().get(0).name);
        assertEquals(2, result.errors().get(0).rowNum());
        assertEquals(3, result.errors().get(0).fileRowNum());
        assertEquals("Price", result.errors().get(0).cellErrors().get(0).headerName());
        assertEquals("not-a-number", result.errors().get(0).cellErrors().get(0).cellValue());
        assertEquals(2, result.summary().totalRows());
        assertEquals("not-a-number", result.errors().get(0).rawValues().get(1));
    }

    static class Product {
        String name;
        Integer price;
    }
}
