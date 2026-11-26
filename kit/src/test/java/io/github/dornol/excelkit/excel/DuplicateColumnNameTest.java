package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class DuplicateColumnNameTest {

    // ========================================================================
    // Duplicate column name validation
    // ========================================================================
    @Test
    void duplicateColumnName_shouldThrowInExcelWriter() {
        var writer = ExcelWriter.<String>create()
                .column("Name", s -> s)
                .column("Name", s -> s);
        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void duplicateColumnName_shouldThrowInExcelSheetWriter() {
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            var sheet = wb.<String>sheet("Test")
                    .column("Name", s -> s)
                    .column("Name", s -> s);
            assertThrows(ExcelWriteException.class, () -> sheet.write(Stream.of("test")));
        }
    }

    @Test
    void uniqueColumnNames_shouldNotThrow() {
        assertDoesNotThrow(() ->
                ExcelWriter.<String>create()
                        .column("Name", s -> s)
                        .column("Age", s -> s)
                        .write(Stream.of("test"))
                        .writeTo(java.io.OutputStream.nullOutputStream()));
    }

    @Test
    void sameNameInDifferentGroups_isAllowed() {
        assertDoesNotThrow(() -> ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.group("A"))
                .column("Name", s -> s, c -> c.group("B"))
                .write(Stream.of("test"))
                .writeTo(java.io.OutputStream.nullOutputStream()));
    }

    @Test
    void sameNameInSameGroup_stillThrows() {
        var writer = ExcelWriter.<String>create()
                .column("Name", s -> s, c -> c.group("A"))
                .column("Name", s -> s, c -> c.group("A"));
        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void sameNameInDifferentGroups_inTable_throws() {
        var writer = ExcelWriter.<String>create()
                .table("People")
                .column("Name", s -> s, c -> c.group("A"))
                .column("Name", s -> s, c -> c.group("B"));
        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("test")));
    }
}
