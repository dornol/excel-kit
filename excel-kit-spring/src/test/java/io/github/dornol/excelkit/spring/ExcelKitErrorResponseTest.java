package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;

import java.time.Duration;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitErrorResponseTest {

    @Test
    void reportRows_flattensCellErrors() {
        UploadError error = new UploadError(1, 2,
                io.github.dornol.excelkit.core.RowError.Type.MAPPING,
                List.of("failed"),
                List.of(new CellErrorResponse(2, "Price", "bad", "Failed to set column")),
                List.of("Notebook", "bad"));

        List<ErrorReportRow> rows = ExcelKitErrorResponse.reportRows(List.of(error)).toList();

        assertEquals(1, rows.size());
        assertEquals(1, rows.getFirst().rowNum());
        assertEquals(2, rows.getFirst().fileRowNum());
        assertEquals(2, rows.getFirst().columnIndex());
        assertEquals("Price", rows.getFirst().headerName());
        assertEquals("bad", rows.getFirst().cellValue());
        assertEquals("Notebook | bad", rows.getFirst().rawValues());
    }

    @Test
    void csv_setsDownloadHeaders() {
        UploadResult<Object> result = new UploadResult<>("CSV", 0, 1, List.of(),
                List.of(new UploadError(1, 2,
                        io.github.dornol.excelkit.core.RowError.Type.VALIDATION,
                        List.of("required"), List.of())));

        var response = ExcelKitErrorResponse.csv(result, "errors");

        assertEquals("text/csv; charset=UTF-8", response.getHeaders().getFirst(HttpHeaders.CONTENT_TYPE));
        assertTrue(response.getHeaders().getFirst(HttpHeaders.CONTENT_DISPOSITION).contains("errors.csv"));
    }

    @Test
    void errorReports_allowCustomCacheMaxAge() {
        var errors = List.of(new UploadError(1, 2,
                io.github.dornol.excelkit.core.RowError.Type.VALIDATION,
                List.of("required"), List.of()));

        var csv = ExcelKitErrorResponse.csv(errors, "errors", Duration.ofMinutes(1));
        var excel = ExcelKitErrorResponse.excel(errors, "errors", Duration.ofMinutes(1));

        assertEquals("max-age=60", csv.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertEquals("max-age=60", excel.getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
    }
}
