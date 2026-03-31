package io.github.dornol.excelkit.example.app.book.application.port.out;

import io.github.dornol.excelkit.example.app.book.domain.Book;

import java.util.List;

public interface BookPersistencePort {

    void saveAll(List<Book> books);

    void flushAndClear();

}
