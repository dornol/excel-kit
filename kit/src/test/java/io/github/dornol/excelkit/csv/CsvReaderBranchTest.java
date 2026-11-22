package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.read.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class CsvReaderBranchTest {

    static class Person {
        String name;
        int age;
    }

    @Test
    void csvReader_skipColumn_shouldSkipPositionally() {
        String csv = "Name,Skip,Age\nAlice,ignored,30\n";
        List<ReadResult<Person>> results = new ArrayList<>();
        new CsvReader<>(Person::new, null)
                .column((p, cell) -> p.name = cell.asString())
                .skipColumn()
                .column((p, cell) -> p.age = cell.asInt())
                .read(new ByteArrayInputStream(csv.getBytes()), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name);
        assertEquals(30, results.get(0).data().age);
    }

    @Test
    void csvReader_skipColumns_shouldSkipMultiple() {
        String csv = "A,B,C,D\n1,2,3,4\n";
        List<ReadResult<Person>> results = new ArrayList<>();
        new CsvReader<>(Person::new, null)
                .skipColumns(3)
                .column((p, cell) -> p.name = cell.asString())
                .read(new ByteArrayInputStream(csv.getBytes()), results::add);

        assertEquals(1, results.size());
        assertEquals("4", results.get(0).data().name);
    }

    @Test
    void csvReader_skipColumns_negativeCount_throwsException() {
        assertThrows(IllegalArgumentException.class, () ->
                new CsvReader<>(Person::new, null).skipColumns(-1));
    }

    @Test
    void csvReader_dialect_PIPE() {
        String csv = "Name|Age\nAlice|30\n";
        List<ReadResult<Person>> results = new ArrayList<>();
        new CsvReader<>(Person::new, null)
                .dialect(CsvDialect.PIPE)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Age", (p, cell) -> p.age = cell.asInt())
                .read(new ByteArrayInputStream(csv.getBytes()), results::add);

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).data().name);
        assertEquals(30, results.get(0).data().age);
    }

    @Test
    void csvReader_headerNotFound_throwsException() {
        String csv = "Name,Age\nAlice,30\n";
        var reader = new CsvReader<>(Person::new, null)
                .column("NonExistent", (p, cell) -> p.name = cell.asString());

        assertThrows(ExcelKitException.class, () ->
                reader.read(new ByteArrayInputStream(csv.getBytes()), r -> {}));
    }

    @Test
    void csvReader_emptyFile_throwsCsvReadException() {
        var reader = new CsvReader<>(Person::new, null)
                .column("Name", (p, cell) -> p.name = cell.asString());

        assertThrows(CsvReadException.class, () ->
                reader.read(new ByteArrayInputStream(new byte[0]), r -> {}));
    }
}
