package io.github.dornol.excelkit.example.app.book.adapter.in.web;

import io.github.dornol.excelkit.example.app.book.application.port.in.ExportBookUseCase;
import io.github.dornol.excelkit.example.app.book.application.port.out.StreamingContent;
import io.github.dornol.excelkit.example.app.common.DownloadResponse;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

@Controller
public class BookExportController {

    private final ExportBookUseCase exportBookUseCase;

    public BookExportController(ExportBookUseCase exportBookUseCase) {
        this.exportBookUseCase = exportBookUseCase;
    }

    @GetMapping("/download-excel")
    public ResponseEntity<StreamingResponseBody> downloadExcel() {
        StreamingContent content = exportBookUseCase.exportExcel();
        return DownloadResponse.excel("book list excel")
                .body(content::writeTo);
    }

    @GetMapping("/download-excel-with-password")
    public ResponseEntity<StreamingResponseBody> downloadExcelWithPassword(
            @RequestParam(required = false, defaultValue = "1234") String password) {
        StreamingContent content = exportBookUseCase.exportExcelWithPassword(password);
        return DownloadResponse.excel("book list excel with password")
                .body(content::writeTo);
    }

    @GetMapping("/download-csv")
    public ResponseEntity<StreamingResponseBody> downloadCsv() {
        StreamingContent content = exportBookUseCase.exportCsv();
        return DownloadResponse.csv("book list csv")
                .body(content::writeTo);
    }

}
