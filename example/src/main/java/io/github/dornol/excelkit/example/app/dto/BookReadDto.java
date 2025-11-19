package io.github.dornol.excelkit.example.app.dto;

import jakarta.validation.constraints.AssertTrue;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import jakarta.validation.constraints.Size;

import java.util.Objects;
import java.util.stream.Stream;

public final class BookReadDto {
    @NotNull(message = "ID must not be null")
    private Long id;

    @NotBlank(message = "Title must not be blank")
    @Size(min = 3, max = 195, message = "Title must be between 3 and 195 characters")
    private String title;

    @NotBlank(message = "Subtitle must not be blank")
    @Size(min = 3, max = 195, message = "Subtitle must be between 3 and 195 characters")
    private String subtitle;

    @NotBlank(message = "Author must not be blank")
    @Size(min = 3, max = 195, message = "Author must be between 3 and 195 characters")
    private String author;

    @NotBlank(message = "Publisher must not be blank")
    @Size(min = 3, max = 195, message = "Publisher must be between 3 and 195 characters")
    private String publisher;

    @NotBlank(message = "ISBN must not be blank")
    @Size(min = 3, max = 195, message = "ISBN must be between 3 and 195 characters")
    private String isbn;

    @NotBlank(message = "Description must not be blank")
    @Size(min = 3, max = 195, message = "Description must be between 3 and 195 characters")
    private String description;

    @AssertTrue(message = "@@@@@ contains dhkim!! @@@@@")
    public boolean isCustomValidation() {
        String restricted = "dhkim";
        return Stream.of(title, subtitle, author, publisher, isbn, description)
                .filter(Objects::nonNull)
                .noneMatch(field -> field.contains(restricted));
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSubtitle() {
        return subtitle;
    }

    public void setSubtitle(String subtitle) {
        this.subtitle = subtitle;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getPublisher() {
        return publisher;
    }

    public void setPublisher(String publisher) {
        this.publisher = publisher;
    }

    public String getIsbn() {
        return isbn;
    }

    public void setIsbn(String isbn) {
        this.isbn = isbn;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    @Override
    public String toString() {
        return "BookReadDto{" +
                "id=" + id +
                ", title='" + title + '\'' +
                ", subtitle='" + subtitle + '\'' +
                ", author='" + author + '\'' +
                ", publisher='" + publisher + '\'' +
                ", isbn='" + isbn + '\'' +
                ", description='" + description + '\'' +
                '}';
    }
}
