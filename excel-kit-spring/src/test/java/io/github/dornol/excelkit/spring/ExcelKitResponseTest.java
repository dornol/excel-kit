package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;

import java.time.Duration;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelKitResponseTest {

    @Test
    void excelBuilder_setsDownloadHeaders() {
        var response = ExcelKitResponse.excel("report").build();

        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertEquals("private, max-age=10", response.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("report.xlsx"));
    }

    @Test
    void csvBuilder_setsDownloadHeaders() {
        var response = ExcelKitResponse.csv("report").build();

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertEquals("private, max-age=10", response.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("report.csv"));
    }

    @Test
    void builder_allowsCustomCacheMaxAge() {
        var response = ExcelKitResponse.excel("report", Duration.ofMinutes(5)).build();

        assertEquals("private, max-age=300", response.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
    }

    @Test
    void builder_rejectsNegativeOrNullCacheMaxAge() {
        assertThrows(IllegalArgumentException.class,
                () -> ExcelKitResponse.csv("report", Duration.ofSeconds(-1)));
        assertThrows(IllegalArgumentException.class,
                () -> ExcelKitResponse.csv("report", null));
    }

    private static io.github.dornol.excelkit.excel.ExcelHandler handler() {
        return io.github.dornol.excelkit.excel.ExcelWriter.<String>create()
                .column("A", value -> value)
                .write(java.util.List.of("x"));
    }

    @Test
    void encryptedExcel_rejectsBlankPasswordBeforeBuildingTheResponse() {
        for (String password : new String[]{null, "", "   "}) {
            var handler = handler();

            assertThrows(IllegalArgumentException.class,
                    () -> ExcelKitResponse.excel(handler, "report", password));

            // The rejected handler released its workbook, so it is already consumed.
            assertThrows(io.github.dornol.excelkit.excel.ExcelWriteException.class,
                    () -> handler.writeTo(java.io.OutputStream.nullOutputStream()));
        }
    }

    @Test
    void encryptedExcel_withPasswordStillStreamsLater() throws Exception {
        var handler = handler();

        var response = ExcelKitResponse.excel(handler, "report", "pw1234");

        assertEquals(200, response.getStatusCode().value());
        // Nothing is consumed until the body runs.
        var out = new java.io.ByteArrayOutputStream();
        response.getBody().writeTo(out);
        assertTrue(out.size() > 0);
    }
}
