package io.github.dornol.excelkit.example.app.book.adapter.out.file;

import io.github.dornol.excelkit.example.app.book.application.port.out.FileExportPort;
import io.github.dornol.excelkit.example.app.book.application.port.out.FileImportPort;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookDto;
import io.github.dornol.excelkit.example.app.book.application.port.out.BookReadDto;
import io.github.dornol.excelkit.example.app.book.application.port.out.ImportResult;
import io.github.dornol.excelkit.example.app.book.application.port.out.StreamingContent;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.*;
import jakarta.validation.Validator;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.function.Consumer;
import java.util.stream.Stream;

/**
 * File export/import adapter using excel-kit.
 * This is the only place in the book domain that depends on excel-kit.
 */
@Component
class BookExcelAdapter implements FileExportPort, FileImportPort {

    private final Validator validator;

    public BookExcelAdapter(Validator validator) {
        this.validator = validator;
    }

    @Override
    public StreamingContent exportExcel(Stream<BookDto> data) {
        var handler = createExcelHandler(data);
        return handler::consumeOutputStream;
    }

    @Override
    public StreamingContent exportExcelWithPassword(Stream<BookDto> data, String password) {
        var handler = createExcelHandler(data);
        return out -> handler.consumeOutputStreamWithPassword(out, password);
    }

    @Override
    public StreamingContent exportCsv(Stream<BookDto> data) {
        var handler = new CsvWriter<BookDto>()
                .column("no", (rowData, cursor) -> cursor.getCurrentTotal())
                .column("id", BookDto::id)
                .column("title", BookDto::title)
                .column("subtitle", BookDto::subtitle)
                .column("author", BookDto::author)
                .column("publisher", BookDto::publisher)
                .column("isbn", BookDto::isbn)
                .column("description", BookDto::description)
                .write(data);
        return handler::consumeOutputStream;
    }

    @Override
    public void readExcel(InputStream is, Consumer<ImportResult<BookReadDto>> consumer) {
        createExcelReadHandler(is).read(result ->
                consumer.accept(new ImportResult<>(result.data(), result.success(), result.messages())));
    }

    @Override
    public void readCsv(InputStream is, Consumer<ImportResult<BookReadDto>> consumer) {
        new CsvReader<>(BookReadDto::new, validator)
                .column((r, d) -> {})
                .column((r, d) -> r.setId(d.asLong()))
                .column((r, d) -> r.setTitle(d.asString()))
                .column((r, d) -> r.setSubtitle(d.asString()))
                .column((r, d) -> r.setAuthor(d.asString()))
                .column((r, d) -> r.setPublisher(d.asString()))
                .column((r, d) -> r.setIsbn(d.asString()))
                .column((r, d) -> r.setDescription(d.asString()))
                .build(is)
                .read(result ->
                        consumer.accept(new ImportResult<>(result.data(), result.success(), result.messages())));
    }

    private ExcelHandler createExcelHandler(Stream<BookDto> data) {
        return new ExcelWriter<BookDto>(ExcelColor.of(0xCC, 0xFF, 0x99))
                .tabColor(ExcelColor.STEEL_BLUE)
                .column("no", (rowData, cursor) -> cursor.getCurrentTotal())
                    .type(ExcelDataType.LONG)
                    .fontColor(ExcelColor.GRAY)
                .column("id", BookDto::id)
                    .type(ExcelDataType.LONG)
                .column("title", BookDto::title)
                    .bold(true)
                    .fontColor(ExcelColor.BLUE)
                    .underline()
                .column("subtitle", BookDto::subtitle)
                .column("author", BookDto::author)
                    .rotation(0)
                    .borderBottom(ExcelBorderStyle.MEDIUM)
                .column("publisher", BookDto::publisher)
                .column("isbn", BookDto::isbn)
                    .validation(ExcelValidation.textLength(10, 13)
                            .errorTitle("Invalid ISBN")
                            .errorMessage("ISBN must be 10–13 characters"))
                .column("description", BookDto::description)
                    .fontColor(ExcelColor.GRAY)
                    .strikethrough(false)
                .afterData(ctx -> {
                    if (ctx.getCurrentRow() > 2) {
                        ctx.groupRows(1, ctx.getCurrentRow() - 1);
                    }
                    return ctx.getCurrentRow();
                })
                .write(data);
    }

    private ExcelReadHandler<BookReadDto> createExcelReadHandler(InputStream is) {
        return new ExcelReader<>(BookReadDto::new, validator)
                .column((r, d) -> {})
                .column((r, d) -> r.setId(d.asLong()))
                .column((r, d) -> r.setTitle(d.asString()))
                .column((r, d) -> r.setSubtitle(d.asString()))
                .column((r, d) -> r.setAuthor(d.asString()))
                .column((r, d) -> r.setPublisher(d.asString()))
                .column((r, d) -> r.setIsbn(d.asString()))
                .column((r, d) -> r.setDescription(d.asString()))
                .build(is);
    }

}
