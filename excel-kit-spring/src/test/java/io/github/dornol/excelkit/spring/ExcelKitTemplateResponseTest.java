package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;

import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitTemplateResponseTest {

    private static final ExcelKitSchema<Product> SCHEMA = ExcelKitSchema.<Product>builder()
            .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
            .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt())
            .build();

    @Test
    void excel_setsDownloadHeaders() throws IOException {
        var response = ExcelKitTemplateResponse.excel(SCHEMA, "template");

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("template.xlsx"));
        try (var workbook = StreamedBody.workbook(response)) {
            var sheet = workbook.getSheet("Template");
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Price", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals(0, sheet.getLastRowNum(), "an empty template has only the header row");
        }
    }

    @Test
    void csv_setsDownloadHeaders() throws IOException {
        var response = ExcelKitTemplateResponse.csv(SCHEMA, "template");

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("template.csv"));
        String header = StreamedBody.text(response).lines().findFirst().orElseThrow();
        assertTrue(header.contains("Name") && header.contains("Price"), header);
    }

    @Test
    void excel_acceptsSampleRows() throws IOException {
        var response = ExcelKitTemplateResponse.excel(SCHEMA, "template",
                List.of(new Product("Notebook", 1200)));

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        try (var workbook = StreamedBody.workbook(response)) {
            var row = workbook.getSheet("Template").getRow(1);
            assertEquals("Notebook", row.getCell(0).getStringCellValue());
            assertEquals(1200, (int) row.getCell(1).getNumericCellValue());
        }
    }

    @Test
    void excelWithGuidance_setsDownloadHeaders() throws IOException {
        var response = ExcelKitTemplateResponse.excelWithGuidance(SCHEMA, "template");

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        try (var workbook = StreamedBody.workbook(response)) {
            var sheet = workbook.getSheet("Template");
            assertTrue(sheet.getRow(0).getCell(0).getStringCellValue().startsWith("Required columns:"));
            assertEquals("Name", sheet.getRow(1).getCell(0).getStringCellValue(),
                    "the header row follows the guidance row");
        }
    }

    @Test
    void csv_acceptsSampleRows() throws IOException {
        var response = ExcelKitTemplateResponse.csv(SCHEMA, "template",
                List.of(new Product("Notebook", 1200)));

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        String text = StreamedBody.text(response);
        assertTrue(text.contains("Notebook") && text.contains("1200"), text);
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
