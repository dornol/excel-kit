package io.github.dornol.excelkit.example.app.book.application.port.out;

import java.util.stream.Stream;

public interface FileExportPort {

    StreamingContent exportExcel(Stream<BookDto> data);

    StreamingContent exportExcelWithPassword(Stream<BookDto> data, String password);

    StreamingContent exportCsv(Stream<BookDto> data);

}
