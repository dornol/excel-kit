package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelKitUploadTest {

    @Test
    void csv_opensMultipartFileAndCollectsUploadResult() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nNotebook,1200\nPen,bad\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.csv(file, inputStream -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Price", (p, cell) -> p.price = cell.asInt())
                .build(inputStream));

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, result.errorCount());
        assertEquals("Notebook", result.rows().getFirst().name);
    }

    @Test
    void read_usesCustomTypeLabel() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name\nNotebook\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.read("Products", file, inputStream -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .build(inputStream));

        assertEquals("Products", result.type());
        assertEquals(1, result.successCount());
        assertEquals(0, result.errorCount());
    }

    static class Product {
        String name;
        Integer price;
    }
}
