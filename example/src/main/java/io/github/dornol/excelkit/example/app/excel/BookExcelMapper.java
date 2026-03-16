package io.github.dornol.excelkit.example.app.excel;

import io.github.dornol.excelkit.example.app.dto.BookDto;
import io.github.dornol.excelkit.example.app.dto.BookReadDto;
import io.github.dornol.excelkit.excel.*;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.stream.Stream;

/**
 * Example mapper for Book entities to Excel.
 * Demonstrates how to use ExcelWriter for exports and ExcelReader for imports.
 */
public class BookExcelMapper {

    private BookExcelMapper() {
        /* empty */
    }

    /**
     * Configures and returns an ExcelHandler for exporting BookDto data.
     * <p>
     * Demonstrates: tab color, font color, strikethrough, underline, rotation,
     * per-side borders, advanced data validation, and row grouping.
     *
     * @param stream Stream of BookDto data
     * @return ExcelHandler for outputting the file
     */
    public static ExcelHandler getHandler(Stream<BookDto> stream) {
        return new ExcelWriter<BookDto>(0xCC, 0xFF, 0x99)
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
                    // Group all data rows for collapsible view
                    if (ctx.getCurrentRow() > 2) {
                        ctx.groupRows(1, ctx.getCurrentRow() - 1);
                    }
                    return ctx.getCurrentRow();
                })
                .write(stream);
    }

    /**
     * Configures and returns an ExcelReadHandler for importing BookReadDto data.
     *
     * @param inputStream Excel file input stream
     * @param validator   Optional bean validator
     * @return ExcelReadHandler for reading the file
     */
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
