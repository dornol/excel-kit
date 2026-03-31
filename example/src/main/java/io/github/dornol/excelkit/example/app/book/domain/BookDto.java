package io.github.dornol.excelkit.example.app.book.domain;

public record BookDto(
        Long id,
        String title,
        String subtitle,
        String author,
        String publisher,
        String isbn,
        String description) {
}
