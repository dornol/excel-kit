package io.github.dornol.excelkit.example.app.book.application.port.out;

import java.util.stream.Stream;

public interface BookQueryPort {

    Stream<BookDto> streamAll();

}
