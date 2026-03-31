package io.github.dornol.excelkit.example.app.book.adapter.out.persistence;

import io.github.dornol.excelkit.example.app.book.application.port.out.BookPersistencePort;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookQueryPort;
import io.github.dornol.excelkit.example.app.book.domain.Book;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookDto;
import jakarta.persistence.EntityManager;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.stream.Stream;

@Repository
class BookJpaAdapter implements BookQueryPort, BookPersistencePort {

    private final BookJpaRepository bookJpaRepository;
    private final EntityManager em;

    public BookJpaAdapter(BookJpaRepository bookJpaRepository, EntityManager em) {
        this.bookJpaRepository = bookJpaRepository;
        this.em = em;
    }

    @Override
    public Stream<BookDto> streamAll() {
        return bookJpaRepository.getStream();
    }

    @Override
    public void saveAll(List<Book> books) {
        bookJpaRepository.saveAll(books);
    }

    @Override
    public void flushAndClear() {
        em.flush();
        em.clear();
    }

}
