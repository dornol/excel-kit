package io.github.dornol.excelkit.example.app.book.application.port.in;

import java.io.InputStream;

public interface ImportBookUseCase {

    void importExcel(InputStream is);

    void importAndSaveExcel(InputStream is);

    void importCsv(InputStream is);

}
