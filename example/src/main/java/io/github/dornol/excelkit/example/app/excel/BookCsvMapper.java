package io.github.dornol.excelkit.example.app.excel;


import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvReadHandler;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.example.app.dto.BookDto;
import io.github.dornol.excelkit.example.app.dto.BookReadDto;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.stream.Stream;

public final class BookCsvMapper {
    private BookCsvMapper() {
        /* empty */
    }

    public static CsvHandler getHandler(Stream<BookDto> stream) {
        return new CsvWriter<BookDto>()
                .column("no", (rowData, cursor) -> cursor.getCurrentTotal())
                .column("id", BookDto::id)
                .column("title", BookDto::title)
                .column("subtitle", BookDto::subtitle)
                .column("author", BookDto::author)
                .column("publisher", BookDto::publisher)
                .column("isbn", BookDto::isbn)
                .column("description", BookDto::description)
                .write(stream);
    }

    public static CsvReadHandler<BookReadDto> getReadHandler(InputStream inputStream, Validator validator) {
        return new CsvReader<>(BookReadDto::new, validator)
                .column((r, d) -> {})
                .column((r, d) -> r.setId(d.asLong()))
                .column((r, d) -> r.setTitle(d.asString()))
                .column((r, d) -> r.setSubtitle(d.asString()))
                .column((r, d) -> r.setAuthor(d.asString()))
                .column((r, d) -> r.setPublisher(d.asString()))
                .column((r, d) -> r.setIsbn(d.asString()))
                .column((r, d) -> r.setDescription(d.asString()))
                .build(inputStream);
    }

}
