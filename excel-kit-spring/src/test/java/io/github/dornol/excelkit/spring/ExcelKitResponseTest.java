package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitResponseTest {

    @Test
    void excelBuilder_setsDownloadHeaders() {
        var response = ExcelKitResponse.excel("report").build();

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertEquals("max-age=10", response.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("report.xlsx"));
    }

    @Test
    void csvBuilder_setsDownloadHeaders() {
        var response = ExcelKitResponse.csv("report").build();

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertEquals("max-age=10", response.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("report.csv"));
    }
}
