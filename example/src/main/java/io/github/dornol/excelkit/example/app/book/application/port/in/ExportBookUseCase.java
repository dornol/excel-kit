package io.github.dornol.excelkit.example.app.book.application.port.in;

import io.github.dornol.excelkit.example.app.book.application.port.out.StreamingContent;

public interface ExportBookUseCase {

    StreamingContent exportExcel();

    StreamingContent exportExcelWithPassword(String password);

    StreamingContent exportCsv();

}
