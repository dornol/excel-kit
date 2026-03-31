package io.github.dornol.excelkit.example.app.book.application.port.out;

import io.github.dornol.excelkit.example.app.book.domain.BookReadDto;
import io.github.dornol.excelkit.example.app.book.domain.ImportResult;

import java.io.InputStream;
import java.util.function.Consumer;

public interface FileImportPort {

    void readExcel(InputStream is, Consumer<ImportResult<BookReadDto>> consumer);

    void readCsv(InputStream is, Consumer<ImportResult<BookReadDto>> consumer);

}
