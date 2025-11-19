package io.github.dornol.excelkit.example.app.dto;

public record BookDto(
        Long id,
        String title,
        String subtitle,
        String author,
        String publisher,
        String isbn,
        String description) {
}
