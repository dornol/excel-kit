package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.Consumer;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * A reader copies its input to a temp file when it opens it. If the reader's own configuration is
 * rejected afterwards, no handler is returned, so that copy must be released immediately.
 */
class ReadHandlerConstructionFailureTest {
    private static final Pattern UUID_NAMED =
            Pattern.compile("[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}.*");

    static class Row {
        String name;
    }

    private static long stagedInputCount() throws IOException {
        try (Stream<Path> entries = Files.list(Path.of(System.getProperty("java.io.tmpdir")))) {
            return entries.filter(path -> UUID_NAMED.matcher(path.getFileName().toString()).matches()).count();
        }
    }

    private static void assertRejectedWithoutLeavingTheInputCopy(Runnable rejected) throws IOException {
        long before = stagedInputCount();
        assertThrows(IllegalArgumentException.class, rejected::run);
        assertEquals(before, stagedInputCount());
    }

    private static ByteArrayInputStream csv() {
        return new ByteArrayInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
    }

    private static final Consumer<ReadResult<Row>> IGNORE = result -> { };

    @Test
    void csvNegativeHeaderRowIndex() throws IOException {
        assertRejectedWithoutLeavingTheInputCopy(() -> CsvReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .headerRowIndex(-1).read(csv(), IGNORE));
    }

    @Test
    void excelNegativeHeaderRowIndex() throws IOException {
        assertRejectedWithoutLeavingTheInputCopy(() -> ExcelReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .headerRowIndex(-1).read(csv(), IGNORE));
    }

    @Test
    void excelNegativeSheetIndex() throws IOException {
        assertRejectedWithoutLeavingTheInputCopy(() -> ExcelReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .sheetIndex(-1).read(csv(), IGNORE));
    }

    @Test
    void excelHeaderRowsBelowOne() throws IOException {
        assertRejectedWithoutLeavingTheInputCopy(() -> ExcelReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .headerRows(0).read(csv(), IGNORE));
    }
}
