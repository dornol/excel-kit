package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for header name-based column mapping in CsvReader.
 */
class CsvNameBasedReadTest {

    @Test
    void readByName_shouldMatchByHeaderName() {
        String csv = "Name,Age,City\nAlice,30,Seoul\nBob,25,Busan\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();
        new CsvReader<>(TestPerson::new, null)
                .column("Name", (TestPerson p, CellData cell) -> p.name = cell.asString())
                .column("Age", (TestPerson p, CellData cell) -> p.age = cell.asInt())
                .column("City", (TestPerson p, CellData cell) -> p.city = cell.asString())
                .build(is)
                .read(r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Seoul", results.get(0).city);
    }

    @Test
    void readByName_shouldWorkWithDifferentColumnOrder() {
        // CSV has columns in different order
        String csv = "City,Age,Name\nSeoul,30,Alice\nBusan,25,Bob\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();
        new CsvReader<>(TestPerson::new, null)
                .column("Name", (TestPerson p, CellData cell) -> p.name = cell.asString())
                .column("Age", (TestPerson p, CellData cell) -> p.age = cell.asInt())
                .column("City", (TestPerson p, CellData cell) -> p.city = cell.asString())
                .build(is)
                .read(r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Seoul", results.get(0).city);
        assertEquals("Bob", results.get(1).name);
    }

    @Test
    void readByName_shouldReadSubsetOfColumns() {
        String csv = "Name,Age,City,Email\nAlice,30,Seoul,alice@test.com\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();
        new CsvReader<>(TestPerson::new, null)
                .column("Name", (TestPerson p, CellData cell) -> p.name = cell.asString())
                .column("City", (TestPerson p, CellData cell) -> p.city = cell.asString())
                .build(is)
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals("Seoul", results.get(0).city);
        assertNull(results.get(0).age);
    }

    @Test
    void readByName_shouldThrowWhenHeaderNotFound() {
        String csv = "Name,Age\nAlice,30\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        CsvReadHandler<TestPerson> handler = new CsvReader<>(TestPerson::new, null)
                .column("Name", (TestPerson p, CellData cell) -> p.name = cell.asString())
                .column("NonExistent", (TestPerson p, CellData cell) -> p.city = cell.asString())
                .build(is);

        assertThrows(CsvReadException.class, () -> handler.read(r -> {}));
    }

    @Test
    void readByName_shouldWorkWithAddColumnMethod() {
        String csv = "City,Name\nSeoul,Alice\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestPerson> results = new ArrayList<>();
        new CsvReader<>(TestPerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .addColumn("City", (p, cell) -> p.city = cell.asString())
                .build(is)
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals("Seoul", results.get(0).city);
    }

    @Test
    void readByNameAsStream_shouldWorkWithDifferentOrder() {
        String csv = "City,Name\nSeoul,Alice\nBusan,Bob\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<String> names = new CsvReader<>(TestPerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .build(is)
                .readAsStream()
                .map(r -> r.data().name)
                .toList();

        assertEquals(2, names.size());
        assertEquals("Alice", names.get(0));
        assertEquals("Bob", names.get(1));
    }

    public static class TestPerson {
        String name;
        Integer age;
        String city;
    }
}
