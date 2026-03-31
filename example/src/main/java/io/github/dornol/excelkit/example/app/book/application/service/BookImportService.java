package io.github.dornol.excelkit.example.app.book.application.service;

import io.github.dornol.excelkit.example.app.book.application.port.in.ImportBookUseCase;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookPersistencePort;
import io.github.dornol.excelkit.example.app.book.application.port.out.FileImportPort;
import io.github.dornol.excelkit.example.app.book.domain.Book;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
class BookImportService implements ImportBookUseCase {

    private static final Logger log = LoggerFactory.getLogger(BookImportService.class);

    private final BookPersistencePort bookPersistencePort;
    private final FileImportPort fileImportPort;

    public BookImportService(BookPersistencePort bookPersistencePort, FileImportPort fileImportPort) {
        this.bookPersistencePort = bookPersistencePort;
        this.fileImportPort = fileImportPort;
    }

    @Override
    public void importExcel(InputStream is) {
        fileImportPort.readExcel(is, result ->
                log.info("success: {}, error: {}, dto: {}", result.success(), result.messages(), result.data()));
    }

    @Transactional
    @Override
    public void importAndSaveExcel(InputStream is) {
        long start = System.currentTimeMillis();
        final int[] counts = {0};
        final int[] successCounts = {0};
        final int[] errorCounts = {0};
        List<Book> books = new ArrayList<>();

        fileImportPort.readExcel(is, result -> {
            if (result.success()) {
                var data = result.data();
                books.add(new Book(null, data.getTitle(), data.getSubtitle(),
                        data.getAuthor(), data.getPublisher(), data.getIsbn(), data.getDescription()));
                successCounts[0]++;
            } else {
                errorCounts[0]++;
                log.warn("failed to read excel: {}", result.messages());
            }
            counts[0]++;
            if (counts[0] > 1000) {
                bookPersistencePort.saveAll(books);
                bookPersistencePort.flushAndClear();
                books.clear();
                counts[0] = 0;
            }
        });

        if (!books.isEmpty()) {
            bookPersistencePort.saveAll(books);
        }

        long end = System.currentTimeMillis();
        log.info("read and save excel finished: {}ms, success: {}, error: {}", end - start, successCounts[0], errorCounts[0]);
    }

    @Override
    public void importCsv(InputStream is) {
        fileImportPort.readCsv(is, result ->
                log.info("success: {}, error: {}, dto: {}", result.success(), result.messages(), result.data()));
    }

}
