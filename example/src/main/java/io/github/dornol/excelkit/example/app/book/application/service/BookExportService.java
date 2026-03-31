package io.github.dornol.excelkit.example.app.book.application.service;

import io.github.dornol.excelkit.example.app.book.application.port.in.ExportBookUseCase;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookQueryPort;
import io.github.dornol.excelkit.example.app.book.application.port.out.FileExportPort;
import io.github.dornol.excelkit.example.app.book.domain.StreamingContent;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

@Service
class BookExportService implements ExportBookUseCase {

    private final BookQueryPort bookQueryPort;
    private final FileExportPort fileExportPort;

    public BookExportService(BookQueryPort bookQueryPort, FileExportPort fileExportPort) {
        this.bookQueryPort = bookQueryPort;
        this.fileExportPort = fileExportPort;
    }

    @Transactional(readOnly = true)
    @Override
    public StreamingContent exportExcel() {
        return fileExportPort.exportExcel(bookQueryPort.streamAll());
    }

    @Transactional(readOnly = true)
    @Override
    public StreamingContent exportExcelWithPassword(String password) {
        return fileExportPort.exportExcelWithPassword(bookQueryPort.streamAll(), password);
    }

    @Transactional(readOnly = true)
    @Override
    public StreamingContent exportCsv() {
        return fileExportPort.exportCsv(bookQueryPort.streamAll());
    }

}
