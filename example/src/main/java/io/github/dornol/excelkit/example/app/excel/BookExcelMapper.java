package io.github.dornol.excelkit.example.app.excel;

import io.github.dornol.excelkit.example.app.dto.BookDto;
import io.github.dornol.excelkit.example.app.dto.BookReadDto;
import io.github.dornol.excelkit.excel.*;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.stream.Stream;

public class BookExcelMapper {

    private BookExcelMapper() {
        /* empty */
    }

    public static ExcelHandler getHandler(Stream<BookDto> stream) {
        return new ExcelWriter<BookDto>(0xCC, 0xFF, 0x99)
                .column("no", (rowData, cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.INTEGER)
                .column("id", BookDto::id).type(ExcelDataType.LONG)
                .column("title", BookDto::title)
                .column("subtitle", BookDto::subtitle)
                .column("author", BookDto::author)
                .column("publisher", BookDto::publisher)
                .column("isbn", BookDto::isbn)
                .column("description", BookDto::description)
                .write(stream);
    }

    public static ExcelReadHandler<BookReadDto> getReadHandler(InputStream inputStream, Validator validator) {
        return new ExcelReader<>(BookReadDto::new, validator)
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
