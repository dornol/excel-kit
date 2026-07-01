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
        var handler = CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Price", (p, cell) -> p.price = cell.asInt())
                .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)));

        UploadResult<Product> result = UploadResult.read("CSV", handler);

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, result.errorCount());
        assertEquals("Notebook", result.rows().getFirst().name);
        assertEquals(2, result.errors().getFirst().rowNum());
        assertEquals(3, result.errors().getFirst().fileRowNum());
        assertEquals("Price", result.errors().getFirst().cellErrors().getFirst().headerName());
        assertEquals("not-a-number", result.errors().getFirst().cellErrors().getFirst().cellValue());
        assertEquals(2, result.summary().totalRows());
        assertEquals("not-a-number", result.errors().getFirst().rawValues().get(1));
    }

    static class Product {
        String name;
        Integer price;
    }
}
