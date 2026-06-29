package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitSchema;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitTemplateResponseTest {

    private static final ExcelKitSchema<Product> SCHEMA = ExcelKitSchema.<Product>builder()
            .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
            .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt())
            .build();

    @Test
    void excel_setsDownloadHeaders() {
        var response = ExcelKitTemplateResponse.excel(SCHEMA, "template");

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("template.xlsx"));
    }

    @Test
    void csv_setsDownloadHeaders() {
        var response = ExcelKitTemplateResponse.csv(SCHEMA, "template");

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("template.csv"));
    }

    @Test
    void excel_acceptsSampleRows() {
        var response = ExcelKitTemplateResponse.excel(SCHEMA, "template",
                List.of(new Product("Notebook", 1200)));

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
    }

    @Test
    void csv_acceptsSampleRows() {
        var response = ExcelKitTemplateResponse.csv(SCHEMA, "template",
                List.of(new Product("Notebook", 1200)));

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
    }

    static class Product {
        String name;
        Integer price;

        Product(String name, Integer price) {
            this.name = name;
            this.price = price;
        }
    }
}
