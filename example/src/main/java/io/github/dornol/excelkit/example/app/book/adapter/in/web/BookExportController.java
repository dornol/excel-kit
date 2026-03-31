package io.github.dornol.excelkit.example.app.book.adapter.in.web;

import io.github.dornol.excelkit.example.app.book.application.port.in.ExportBookUseCase;
import io.github.dornol.excelkit.example.app.book.application.port.out.StreamingContent;
import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
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
        return DownloadUtil.builder("book list excel", DownloadFileType.EXCEL)
                .body(content::writeTo);
    }

    @GetMapping("/download-excel-with-password")
    public ResponseEntity<StreamingResponseBody> downloadExcelWithPassword(
            @RequestParam(required = false, defaultValue = "1234") String password) {
        StreamingContent content = exportBookUseCase.exportExcelWithPassword(password);
        return DownloadUtil.builder("book list excel with password", DownloadFileType.EXCEL)
                .body(content::writeTo);
    }

    @GetMapping("/download-csv")
    public ResponseEntity<StreamingResponseBody> downloadCsv() {
        StreamingContent content = exportBookUseCase.exportCsv();
        return DownloadUtil.builder("book list csv", DownloadFileType.CSV)
                .body(content::writeTo);
    }

}
