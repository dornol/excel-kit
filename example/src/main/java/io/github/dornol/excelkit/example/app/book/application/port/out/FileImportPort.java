package io.github.dornol.excelkit.example.app.book.application.port.out;

import java.io.InputStream;
import java.util.function.Consumer;

public interface FileImportPort {

    void readExcel(InputStream is, Consumer<ImportResult<BookReadDto>> consumer);

    void readCsv(InputStream is, Consumer<ImportResult<BookReadDto>> consumer);

}
