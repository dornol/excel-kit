package io.github.dornol.excelkit.example.app.book.adapter.in.web;

import io.github.dornol.excelkit.example.app.book.application.port.in.ImportBookUseCase;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@Controller
public class BookImportController {

    private final ImportBookUseCase importBookUseCase;

    public BookImportController(ImportBookUseCase importBookUseCase) {
        this.importBookUseCase = importBookUseCase;
    }

    @PostMapping("/read-excel")
    public String readExcel(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream()) {
            importBookUseCase.importExcel(inputStream);
        }
        return "redirect:/";
    }

    @PostMapping("/read-and-save")
    public String readAndSaveExcel(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream()) {
            importBookUseCase.importAndSaveExcel(inputStream);
        }
        return "redirect:/";
    }

    @PostMapping("/read-csv")
    public String readCsv(MultipartFile file) throws IOException {
        try (InputStream inputStream = file.getInputStream()) {
            importBookUseCase.importCsv(inputStream);
        }
        return "redirect:/";
    }

}
